var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

// Relationship types that may appear at most once per document. When we merge
// files rendered from the same template, every file carries the same singleton
// rels (styles, theme, fontTable, …) — we keep file 0's ids unchanged for these
// so the merge phase de-dupes them naturally instead of producing duplicate
// "package-relationship" entries that Word flags during Open-and-Repair.
var SINGLETON_REL_TYPES = {
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/glossaryDocument': true,
    'http://schemas.openxmlformats.org/officeDocument/2006/relationships/people': true
};

var sectionTypeMap = {
    'section-continuous': 'continuous',
    'section-newpage': 'nextPage',
    'section-evenpage': 'evenPage',
    'section-oddpage': 'oddPage'
};

// ---- Section property helpers (unchanged behaviour) ------------------------

function extractSectionProperties(xml) {
    var sectPrStartIndex = xml.lastIndexOf('<w:sectPr');
    if (sectPrStartIndex === -1) return null;
    var sectPrEndIndex = xml.indexOf('</w:sectPr>', sectPrStartIndex);
    if (sectPrEndIndex === -1) return null;
    sectPrEndIndex += 11;
    return xml.slice(sectPrStartIndex, sectPrEndIndex);
}

function wrapSectionProperties(sectPr) {
    return '<w:p><w:pPr>' + sectPr + '</w:pPr></w:p>';
}

function setSectionType(sectPr, type) {
    var baseSectPr = sectPr || '<w:sectPr/>';
    if (baseSectPr.endsWith('/>')) {
        baseSectPr = baseSectPr.slice(0, -2) + '></w:sectPr>';
    }
    if (/<w:type\b[^>]*w:val="[^"]*"\s*\/>/.test(baseSectPr)) {
        return baseSectPr.replace(/<w:type\b([^>]*)w:val="[^"]*"([^>]*)\/>/, '<w:type$1w:val="' + type + '"$2/>');
    }
    if (/<w:type\b[^>]*>.*?<\/w:type>/.test(baseSectPr)) {
        return baseSectPr.replace(/<w:type\b[^>]*>.*?<\/w:type>/, '<w:type w:val="' + type + '"/>');
    }
    return baseSectPr.replace('</w:sectPr>', '<w:type w:val="' + type + '"/></w:sectPr>');
}

function resolveBreakMode(options) {
    if (options && options.breakMode) return options.breakMode;
    if (options && typeof options.pageBreak !== 'undefined') return options.pageBreak ? 'template' : 'none';
    return 'template';
}

// ---- Path / part helpers ---------------------------------------------------

// Parts whose content varies per source file and that must be kept unique when
// merging multiple docx files. Other parts (styles, numbering, theme, settings,
// fontTable, comments, footnotes, endnotes, glossary) are treated as global
// singletons and de-duplicated.
function isPerFilePart(absolutePath) {
    if (!absolutePath) return false;
    if (/^word\/header\d*\.xml$/i.test(absolutePath)) return true;
    if (/^word\/footer\d*\.xml$/i.test(absolutePath)) return true;
    if (/^word\/media\//.test(absolutePath)) return true;
    if (/^word\/embeddings\//.test(absolutePath)) return true;
    if (/^word\/charts\//.test(absolutePath)) return true;
    if (/^word\/diagrams\//.test(absolutePath)) return true;
    if (/^word\/activeX\//.test(absolutePath)) return true;
    if (/^customXml\//.test(absolutePath)) return true;
    return false;
}

// Resolve a Target attribute relative to the directory of the rels file that
// contains it. Returns an absolute (package-root-relative) path string.
function resolveTarget(relsFilePath, target) {
    if (/^https?:\/\//i.test(target)) return target;
    // rels file `dir/_rels/foo.rels` describes `dir/foo`; targets are relative to `dir/`.
    var parts = relsFilePath.split('/');
    parts.pop(); // strip filename
    parts.pop(); // strip _rels
    var base = parts.join('/');
    var combined = base ? base + '/' + target : target;
    var stack = [];
    combined.split('/').forEach(function (seg) {
        if (seg === '' || seg === '.') return;
        if (seg === '..') stack.pop();
        else stack.push(seg);
    });
    return stack.join('/');
}

// Inverse of resolveTarget: produce a Target attribute (relative path) for a
// given absolute package path and a rels file location.
function makeRelativeTarget(relsFilePath, absolutePath) {
    var parts = relsFilePath.split('/');
    parts.pop();
    parts.pop();
    var base = parts.join('/');
    if (!base) return absolutePath;
    var prefix = base + '/';
    if (absolutePath.indexOf(prefix) === 0) {
        return absolutePath.slice(prefix.length);
    }
    return absolutePath;
}

// Append a suffix before the extension.
function suffixPath(path, suffix) {
    var slash = path.lastIndexOf('/');
    var dir = slash >= 0 ? path.slice(0, slash) : '';
    var name = slash >= 0 ? path.slice(slash + 1) : path;
    var dot = name.lastIndexOf('.');
    var base = dot >= 0 ? name.slice(0, dot) : name;
    var ext = dot >= 0 ? name.slice(dot) : '';
    return (dir ? dir + '/' : '') + base + suffix + ext;
}

// Map a part path to its associated rels file path.
//   word/header1.xml -> word/_rels/header1.xml.rels
//   customXml/item1.xml -> customXml/_rels/item1.xml.rels
function partToRelsPath(partPath) {
    var slash = partPath.lastIndexOf('/');
    var dir = slash >= 0 ? partPath.slice(0, slash) : '';
    var name = slash >= 0 ? partPath.slice(slash + 1) : partPath;
    var relsDir = dir ? dir + '/_rels' : '_rels';
    return relsDir + '/' + name + '.rels';
}

function listRelsFiles(zip) {
    var out = [];
    var files = zip.files;
    for (var f in files) {
        if (!Object.prototype.hasOwnProperty.call(files, f)) continue;
        if (/\.rels$/i.test(f) && !files[f].dir) out.push(f);
    }
    return out;
}

// Compact content signature for a media file: length plus a djb2 hash over
// three 64-byte windows (head/middle/tail). Designed for media files where
// collisions across this signature are astronomically unlikely. ~190 bytes of
// hashing per file instead of comparing 100KB–1MB binaries in full.
function mediaSignature(bin) {
    var bytes = bin.asUint8Array();
    var len = bytes.length;
    if (len === 0) return '0:0';
    var h = 5381;
    var sampleStarts = [0, Math.floor(len / 2), Math.max(0, len - 64)];
    for (var s = 0; s < sampleStarts.length; s++) {
        var start = sampleStarts[s];
        var end = Math.min(start + 64, len);
        for (var i = start; i < end; i++) {
            h = ((h << 5) + h + bytes[i]) | 0;
        }
    }
    return len + ':' + (h >>> 0).toString(16);
}

function escapeRegex(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function escapeXmlAttr(s) {
    return String(s)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;');
}

// Replace rId attribute values in an XML blob using a rename map. Restricted to
// attribute context (`="rId..."`) so text content is never touched, and sorted
// longest-key-first so rId1 doesn't accidentally match the leading characters
// of rId10.
function applyRIdMap(xml, map) {
    var keys = Object.keys(map);
    if (keys.length === 0) return xml;
    keys.sort(function (a, b) { return b.length - a.length; });
    keys.forEach(function (oldId) {
        var re = new RegExp('="' + escapeRegex(oldId) + '"', 'g');
        xml = xml.replace(re, '="' + map[oldId] + '"');
    });
    return xml;
}

// ---- Per-file prep ---------------------------------------------------------

// For files[index > 0], rename per-file parts (headers, footers, media, ...),
// rewrite their Target references in every rels file, suffix rIds in the main
// document.xml.rels (so they don't collide with other files' rIds after merge)
// and update document.xml's rId references accordingly.
//
// `mediaIndex` is mutable cross-file state mapping a content signature →
// canonical absolute path in the merged output. Media files whose signature
// already exists in the index are de-duplicated: their `Target` attributes are
// remapped to the canonical path, and the local file isn't copied to the base.
// This handles both template-baked logos (identical across every render) and
// generated images that happen to share content across renders.
//
// Returns the file's content-type defaults + overrides keyed by post-rename
// PartName so the caller can merge them across files.
function prepareFile(zip, index, mediaIndex) {
    var suffix = index === 0 ? '' : '_f' + index;
    var serializer = new XMLSerializer();
    var pathRenameMap = {}; // resolved oldPath -> new path (file moves to new path, gets copied to base)
    var pathRemapMap = {};  // resolved oldPath -> canonical path elsewhere (file is *not* copied; Target attribute is redirected)

    // Pre-register base file's media under their original paths so subsequent
    // files can deduplicate against them without renaming.
    if (index === 0) {
        for (var p in zip.files) {
            if (!Object.prototype.hasOwnProperty.call(zip.files, p)) continue;
            if (zip.files[p].dir) continue;
            if (!/^word\/media\//.test(p)) continue;
            var b0 = zip.file(p);
            if (!b0) continue;
            var sig0 = mediaSignature(b0);
            if (!mediaIndex[sig0]) mediaIndex[sig0] = p;
        }
        return collectContentTypes(zip, {});
    }

    // 1. Scan rels for per-file targets. For media, consult `mediaIndex` to
    //    dedupe against any previously processed file (base or another render).
    var relsFiles = listRelsFiles(zip);
    relsFiles.forEach(function (relsPath) {
        var bin = zip.file(relsPath);
        if (!bin) return;
        var xmlString = bin.asText();
        var dom = new DOMParser().parseFromString(xmlString, 'text/xml');
        var rels = dom.getElementsByTagName('Relationship');
        for (var i = 0; i < rels.length; i++) {
            var rel = rels[i];
            if (!rel.getAttribute) continue;
            var target = rel.getAttribute('Target');
            var targetMode = rel.getAttribute('TargetMode');
            if (!target || targetMode === 'External') continue;
            var resolved = resolveTarget(relsPath, target);
            if (!isPerFilePart(resolved)) continue;
            if (pathRenameMap[resolved] || pathRemapMap[resolved]) continue;

            if (/^word\/media\//.test(resolved)) {
                var mediaBin = zip.file(resolved);
                if (mediaBin) {
                    var sig = mediaSignature(mediaBin);
                    if (mediaIndex[sig]) {
                        pathRemapMap[resolved] = mediaIndex[sig];
                        continue;
                    }
                    var newPath = suffixPath(resolved, suffix);
                    mediaIndex[sig] = newPath;
                    pathRenameMap[resolved] = newPath;
                    continue;
                }
            }

            pathRenameMap[resolved] = suffixPath(resolved, suffix);
        }
    });

    // 2. Rewrite Target attributes in every rels file. `pathRenameMap` points
    //    at the suffixed local path; `pathRemapMap` points at a canonical path
    //    that already exists in the merged output.
    relsFiles.forEach(function (relsPath) {
        var bin = zip.file(relsPath);
        if (!bin) return;
        var xmlString = bin.asText();
        var dom = new DOMParser().parseFromString(xmlString, 'text/xml');
        var rels = dom.getElementsByTagName('Relationship');
        var changed = false;
        for (var i = 0; i < rels.length; i++) {
            var rel = rels[i];
            if (!rel.getAttribute) continue;
            var target = rel.getAttribute('Target');
            var targetMode = rel.getAttribute('TargetMode');
            if (!target || targetMode === 'External') continue;
            var resolved = resolveTarget(relsPath, target);
            var canonical = pathRenameMap[resolved] || pathRemapMap[resolved];
            if (canonical && canonical !== resolved) {
                rel.setAttribute('Target', makeRelativeTarget(relsPath, canonical));
                changed = true;
            }
        }
        if (changed) {
            var start = xmlString.indexOf('<Relationships');
            var rewritten = start >= 0
                ? xmlString.slice(0, start) + serializer.serializeToString(dom.documentElement)
                : serializer.serializeToString(dom.documentElement);
            zip.file(relsPath, rewritten);
        }
    });

    // 3. Suffix rIds in the main document rels and update document.xml.
    var rIdMap = suffixMainRelsIds(zip, suffix, serializer);
    if (Object.keys(rIdMap).length > 0) {
        var docXml = zip.file('word/document.xml').asText();
        zip.file('word/document.xml', applyRIdMap(docXml, rIdMap));
    }

    // 4. Move the actually-renamed part bytes (and their rels) to new paths.
    Object.keys(pathRenameMap).forEach(function (oldPath) {
        var newPath = pathRenameMap[oldPath];
        if (newPath === oldPath) return;

        var bin = zip.file(oldPath);
        if (bin) {
            var isXml = /\.(xml|rels)$/i.test(oldPath);
            if (isXml) {
                zip.file(newPath, bin.asText());
            } else {
                zip.file(newPath, bin.asUint8Array());
            }
            zip.remove(oldPath);
        }

        var oldRelsPath = partToRelsPath(oldPath);
        var newRelsPath = partToRelsPath(newPath);
        if (oldRelsPath !== newRelsPath) {
            var relsBin = zip.file(oldRelsPath);
            if (relsBin) {
                zip.file(newRelsPath, relsBin.asText());
                zip.remove(oldRelsPath);
            }
        }
    });

    // 5. Drop deduplicated media so copyAuxiliaryParts won't ship them.
    Object.keys(pathRemapMap).forEach(function (oldPath) {
        zip.remove(oldPath);
    });

    return collectContentTypes(zip, pathRenameMap);
}

function suffixMainRelsIds(zip, suffix, serializer) {
    var path = 'word/_rels/document.xml.rels';
    var bin = zip.file(path);
    if (!bin) return {};
    var xmlString = bin.asText();
    var dom = new DOMParser().parseFromString(xmlString, 'text/xml');
    var rels = dom.getElementsByTagName('Relationship');
    var map = {};
    for (var i = 0; i < rels.length; i++) {
        var rel = rels[i];
        if (!rel.getAttribute) continue;
        var oldId = rel.getAttribute('Id');
        if (!oldId) continue;
        // Singleton-type rels keep their id so the merge phase de-dupes them
        // against file 0's rel of the same type. document.xml in this file
        // already references the original id, so no rewrite needed for it.
        var type = rel.getAttribute('Type');
        if (SINGLETON_REL_TYPES[type]) continue;
        var newId = oldId + suffix;
        rel.setAttribute('Id', newId);
        map[oldId] = newId;
    }
    var start = xmlString.indexOf('<Relationships');
    var rewritten = start >= 0
        ? xmlString.slice(0, start) + serializer.serializeToString(dom.documentElement)
        : serializer.serializeToString(dom.documentElement);
    zip.file(path, rewritten);
    return map;
}

function collectContentTypes(zip, pathRenameMap) {
    var defaults = {};
    var overrides = {};
    var bin = zip.file('[Content_Types].xml');
    if (!bin) return { defaults: defaults, overrides: overrides };

    var xmlString = bin.asText();
    var dom = new DOMParser().parseFromString(xmlString, 'text/xml');

    var defs = dom.getElementsByTagName('Default');
    for (var i = 0; i < defs.length; i++) {
        if (!defs[i].getAttribute) continue;
        var ext = defs[i].getAttribute('Extension');
        var ct = defs[i].getAttribute('ContentType');
        if (ext) defaults[ext.toLowerCase()] = ct;
    }

    var ovs = dom.getElementsByTagName('Override');
    for (var j = 0; j < ovs.length; j++) {
        if (!ovs[j].getAttribute) continue;
        var pn = ovs[j].getAttribute('PartName');
        var ct2 = ovs[j].getAttribute('ContentType');
        if (!pn) continue;
        var pathKey = pn.replace(/^\//, '');
        if (pathRenameMap && pathRenameMap[pathKey]) {
            pn = '/' + pathRenameMap[pathKey];
        }
        overrides[pn] = ct2;
    }

    return { defaults: defaults, overrides: overrides };
}

// ---- Merging ---------------------------------------------------------------

function DocxMerger(options, files) {
    options = options || {};

    this._body = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    this._breakMode = resolveBreakMode(options);
    this._files = [];
    var self = this;
    (files || []).forEach(function (file) { self._files.push(new JSZip(file)); });

    this._mergedRels = [];                  // Relationship DOM nodes for the merged document rels
    this._mergedRelsIds = {};               // Id -> true (dedup)
    this._contentTypeDefaults = {};         // ext -> contentType (first wins)
    this._contentTypeOverrides = {};        // PartName -> contentType (first wins)
    this._mergeStyles = options.mergeStyles === true;  // opt-in; default false (templates are identical)

    this._builder = this._body;

    this.insertPageBreak = function () {
        this._builder.push('<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
    };

    this.insertRaw = function (xml) {
        this._builder.push(xml);
    };

    this.mergeBody = function (files) {
        var self = this;
        this._builder = this._body;

        // Phase 1 — prep each file (renames, rId suffixing, content types,
        // cross-file media dedup).
        var mediaIndex = {};
        files.forEach(function (zip, index) {
            var ct = prepareFile(zip, index, mediaIndex);
            Object.keys(ct.defaults).forEach(function (ext) {
                if (!self._contentTypeDefaults[ext]) {
                    self._contentTypeDefaults[ext] = ct.defaults[ext];
                }
            });
            Object.keys(ct.overrides).forEach(function (pn) {
                if (!self._contentTypeOverrides[pn]) {
                    self._contentTypeOverrides[pn] = ct.overrides[pn];
                }
            });
        });

        // Phase 2 — optional styles + numbering renaming.
        //
        // For our primary use case (rendering the same template N times) the
        // styles.xml and numbering.xml are byte-identical across files, so we
        // keep file 0's copies and skip the per-file styleId/numId suffixing
        // entirely. This avoids the bloat of N-copies of every style and
        // sidesteps Word's "Formatvorlagen"-Repair which triggers when the
        // merged styles.xml contains odd combinations of duplicated entries.
        //
        // Pass `mergeStyles: true` to fall back to the legacy renaming path
        // (e.g. when merging documents from genuinely different templates).
        if (this._mergeStyles) {
            var Style = require('./merge-styles');
            var bulletsNumbering = require('./merge-bullets-numberings');
            bulletsNumbering.prepareNumbering(files);
            bulletsNumbering.mergeNumbering(files, this._numbering);
            Style.prepareStyles(files, this._style);
            Style.mergeStyles(files, this._style);
            this._styleHelpers = { Style: Style, bulletsNumbering: bulletsNumbering };
        }

        // Phase 3 — collect all Relationship nodes from each file's main rels.
        files.forEach(function (zip) {
            var bin = zip.file('word/_rels/document.xml.rels');
            if (!bin) return;
            var dom = new DOMParser().parseFromString(bin.asText(), 'text/xml');
            var rels = dom.getElementsByTagName('Relationship');
            for (var i = 0; i < rels.length; i++) {
                var rel = rels[i];
                if (!rel.getAttribute) continue;
                var id = rel.getAttribute('Id');
                if (!id || self._mergedRelsIds[id]) continue;
                self._mergedRelsIds[id] = true;
                self._mergedRels.push(rel.cloneNode());
            }
        });

        // Phase 4 — merge body XML.
        files.forEach(function (zip, index) {
            var xml = zip.file('word/document.xml').asText();
            var sectPr = extractSectionProperties(xml);

            if (index === 0) self._baseZip = zip;
            if (sectPr) self._finalSectPr = sectPr;

            xml = xml.substring(xml.indexOf('<w:body>') + 8);
            xml = xml.substring(0, xml.indexOf('</w:body>'));

            var sectPrIndex = xml.lastIndexOf('<w:sectPr');
            if (sectPrIndex !== -1) xml = xml.substring(0, sectPrIndex);

            self.insertRaw(xml);

            if (index < files.length - 1) {
                switch (self._breakMode) {
                    case 'template':
                        if (sectPr) self.insertRaw(wrapSectionProperties(sectPr));
                        else self.insertPageBreak();
                        break;
                    case 'none':
                        break;
                    case 'section-continuous':
                    case 'section-newpage':
                    case 'section-evenpage':
                    case 'section-oddpage':
                        self.insertRaw(wrapSectionProperties(setSectionType(sectPr, sectionTypeMap[self._breakMode])));
                        break;
                    case 'pagebreak':
                        self.insertPageBreak();
                        break;
                    default:
                        if (sectPr) self.insertRaw(wrapSectionProperties(sectPr));
                        else self.insertPageBreak();
                }
            }
        });
    };

    this.save = function (type, callback) {
        var zip = this._baseZip;

        // 1. Replace body content in the base document.xml.
        var xml = zip.file('word/document.xml').asText();
        var startIndex = xml.indexOf('<w:body>') + 8;
        var endIndex = xml.lastIndexOf('<w:sectPr');
        xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(''));

        if (this._finalSectPr) {
            var finalSectPrStart = xml.lastIndexOf('<w:sectPr');
            if (finalSectPrStart !== -1) {
                var finalSectPrEnd = xml.indexOf('</w:sectPr>', finalSectPrStart);
                if (finalSectPrEnd !== -1) {
                    finalSectPrEnd += 11;
                    xml = xml.slice(0, finalSectPrStart) + this._finalSectPr + xml.slice(finalSectPrEnd);
                }
            }
        }
        zip.file('word/document.xml', xml);

        // 2. Copy renamed per-file parts (and their rels) from non-base files.
        copyAuxiliaryParts(zip, this._files);

        // 3. Write merged [Content_Types].xml.
        writeMergedContentTypes(zip, this._contentTypeDefaults, this._contentTypeOverrides);

        // 4. Write merged word/_rels/document.xml.rels.
        writeMergedRels(zip, this._mergedRels);

        if (this._styleHelpers) {
            this._styleHelpers.bulletsNumbering.generateNumbering(zip, this._numbering);
            this._styleHelpers.Style.generateStyles(zip, this._style);
        }

        // DEFLATE level 1 (fast) is the default — XML/rels content compresses
        // well even at low levels, and shaving compression time matters more
        // than ~5% extra bytes for in-browser report generation. Callers can
        // override via `compressionLevel` on the constructor options.
        callback(zip.generate({
            type: type,
            compression: 'DEFLATE',
            compressionOptions: { level: options.compressionLevel || 1 }
        }));
    };

    if (this._files.length > 0) {
        this.mergeBody(this._files);
    }
}

// Skip-list of parts that are shared singletons and must not be overwritten
// when copying auxiliary content from non-base files.
var SHARED_SINGLETONS = {
    'word/document.xml': true,
    'word/_rels/document.xml.rels': true,
    '[Content_Types].xml': true,
    '_rels/.rels': true,
    'word/styles.xml': true,
    'word/numbering.xml': true,
    'word/fontTable.xml': true,
    'word/settings.xml': true,
    'word/webSettings.xml': true,
    'word/footnotes.xml': true,
    'word/endnotes.xml': true,
    'word/comments.xml': true,
    'word/people.xml': true,
    'word/glossary/document.xml': true,
    'word/theme/theme1.xml': true
};

function isSharedSingleton(path) {
    if (SHARED_SINGLETONS[path]) return true;
    if (/^docProps\//.test(path)) return true;
    return false;
}

function copyAuxiliaryParts(baseZip, allFiles) {
    for (var i = 1; i < allFiles.length; i++) {
        var srcZip = allFiles[i];
        var files = srcZip.files;
        for (var p in files) {
            if (!Object.prototype.hasOwnProperty.call(files, p)) continue;
            var entry = files[p];
            if (entry.dir) continue;
            if (isSharedSingleton(p)) continue;
            if (baseZip.file(p)) continue;

            var bin = srcZip.file(p);
            if (!bin) continue;
            if (/\.(xml|rels)$/i.test(p)) {
                baseZip.file(p, bin.asText());
            } else {
                // Media files (jpeg/png/gif/etc.) are already compressed —
                // re-deflating them costs CPU for no size benefit, so we
                // store them uncompressed when possible.
                baseZip.file(p, bin.asUint8Array(), { binary: true, compression: 'STORE' });
            }
        }
    }
}

function writeMergedContentTypes(zip, defaults, overrides) {
    var out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
    Object.keys(defaults).forEach(function (ext) {
        out += '<Default Extension="' + escapeXmlAttr(ext) + '" ContentType="' + escapeXmlAttr(defaults[ext]) + '"/>';
    });
    Object.keys(overrides).forEach(function (pn) {
        out += '<Override PartName="' + escapeXmlAttr(pn) + '" ContentType="' + escapeXmlAttr(overrides[pn]) + '"/>';
    });
    out += '</Types>';
    zip.file('[Content_Types].xml', out);
}

function writeMergedRels(zip, rels) {
    var out = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
    rels.forEach(function (rel) {
        var attrs = [];
        if (rel.attributes) {
            for (var i = 0; i < rel.attributes.length; i++) {
                var a = rel.attributes[i];
                attrs.push(a.name + '="' + escapeXmlAttr(a.value) + '"');
            }
        }
        out += '<Relationship ' + attrs.join(' ') + '/>';
    });
    out += '</Relationships>';
    zip.file('word/_rels/document.xml.rels', out);
}

module.exports = DocxMerger;
