var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

var Style = require('./merge-styles');
var Media = require('./merge-media');
var RelContentType = require('./merge-relations-and-content-type');
var bulletsNumbering = require('./merge-bullets-numberings');

const sectionTypeMap = {
    'section-continuous': 'continuous',
    'section-newpage': 'nextPage',
    'section-evenpage': 'evenPage',
    'section-oddpage': 'oddPage'
};

function extractSectionProperties(xml) {
    var sectPrStartIndex = xml.lastIndexOf('<w:sectPr');

    if (sectPrStartIndex === -1) {
        return null;
    }

    var sectPrEndIndex = xml.indexOf('</w:sectPr>', sectPrStartIndex);

    if (sectPrEndIndex === -1) {
        return null;
    }

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
    if (options && options.breakMode) {
        return options.breakMode;
    }

    if (options && typeof options.pageBreak !== 'undefined') {
        return options.pageBreak ? 'template' : 'none';
    }

    return 'template';
}

function DocxMerger(options, files) {

    this._body = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    this._breakMode = resolveBreakMode(options);
    this._files = [];
    var self = this;
    (files || []).forEach(function(file) {
        self._files.push(new JSZip(file));
    });
    this._contentTypes = {};

    this._media = {};
    this._rel = {};

    this._builder = this._body;

    this.insertPageBreak = function() {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        this._builder.push(pb);
    };



    this.insertRaw = function(xml) {

        this._builder.push(xml);
    };

    this.mergeBody = function(files) {
        var self = this;
        this._builder = this._body;
    
        RelContentType.mergeContentTypes(files, this._contentTypes);
        Media.prepareMediaFiles(files, this._media);
        RelContentType.mergeRelations(files, this._rel);
    
        bulletsNumbering.prepareNumbering(files);
        bulletsNumbering.mergeNumbering(files, this._numbering);
    
        Style.prepareStyles(files, this._style);
        Style.mergeStyles(files, this._style);
    
        files.forEach(function(zip, index) {
            var xml = zip.file('word/document.xml').asText();
            var sectPr = extractSectionProperties(xml);

            if (index === 0) {
                self._baseZip = zip;
            }

            if (sectPr) {
                self._finalSectPr = sectPr;
            }

            xml = xml.substring(xml.indexOf("<w:body>") + 8);
            xml = xml.substring(0, xml.indexOf("</w:body>"));

            var sectPrIndex = xml.lastIndexOf('<w:sectPr');
            if (sectPrIndex !== -1) {
                xml = xml.substring(0, sectPrIndex);
            }

            self.insertRaw(xml);

            if (index < files.length - 1) {
                switch (self._breakMode) {
                    case 'template':
                        if (sectPr) {
                            self.insertRaw(wrapSectionProperties(sectPr));
                        } else {
                            self.insertPageBreak();
                        }
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
                        if (sectPr) {
                            self.insertRaw(wrapSectionProperties(sectPr));
                        } else {
                            self.insertPageBreak();
                        }
                }
            }
        });
    };

    this.save = function(type, callback) {
        var zip = this._baseZip;
    
        var xml = zip.file("word/document.xml").asText();
        var startIndex = xml.indexOf("<w:body>") + 8;
        var endIndex = xml.lastIndexOf("<w:sectPr");
    
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
    
        RelContentType.generateContentTypes(zip, this._contentTypes);
        Media.copyMediaFiles(zip, this._media, this._files);
        RelContentType.generateRelations(zip, this._rel);
        bulletsNumbering.generateNumbering(zip, this._numbering);
        Style.generateStyles(zip, this._style);
    
        zip.file("word/document.xml", xml);
    
        callback(zip.generate({
            type: type,
            compression: "DEFLATE",
            compressionOptions: {
                level: 4
            }
        }));
    };


    if (this._files.length > 0) {

        this.mergeBody(this._files);
    }
}


module.exports = DocxMerger;
