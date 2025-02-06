var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

var Style = require('./merge-styles');
var Media = require('./merge-media');
var RelContentType = require('./merge-relations-and-content-type');
var bulletsNumbering = require('./merge-bullets-numberings');

function DocxMerger(options, files) {

    this._body = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
    this._files = [];
    var self = this;
    (files || []).forEach(function (file) {
        self._files.push(new JSZip(file));
    });
    this._contentTypes = {};

    this._media = {};
    this._rel = {};

    this._builder = this._body;

    this.insertPageBreak = function () {
        var pb = '<w:p> \
					<w:r> \
						<w:br w:type="page"/> \
					</w:r> \
				  </w:p>';

        this._builder.push(pb);
    };

    this.insertRaw = function (xml) {

        this._builder.push(xml);
    };

    this.mergeBody = function (files) {
        var self = this;
        this._builder = this._body;

        RelContentType.mergeContentTypes(files, this._contentTypes);
        Media.prepareMediaFiles(files, this._media);
        RelContentType.mergeRelations(files, this._rel);

        bulletsNumbering.prepareNumbering(files);
        bulletsNumbering.mergeNumbering(files, this._numbering);

        Style.prepareStyles(files, this._style);
        Style.mergeStyles(files, this._style);

        var sectPr;
        files.forEach(function (zip, index) {
            var xml = zip.file("word/document.xml").asText();
            var bodyStart = xml.indexOf("<w:body>");
            var bodyEnd = xml.indexOf("</w:body>", bodyStart);
            var bodyContent = xml.substring(bodyStart + 8, bodyEnd);

            // Remove the existing sectPr from the body content
            var sectPrIndex = bodyContent.lastIndexOf("<w:sectPr");
            if (sectPrIndex !== -1) {
                bodyContent = bodyContent.substring(0, sectPrIndex);
            }

            self.insertRaw(bodyContent);

            if (index === 0) {
                self._baseZip = zip;

                // Extract sectPr from the original document.xml of the first file
                var sectPrStartIndex = xml.lastIndexOf("<w:sectPr");
                if (sectPrStartIndex !== -1) {
                    var sectPrEndIndex = xml.indexOf("</w:sectPr>", sectPrStartIndex);
                    if (sectPrEndIndex !== -1) {
                        sectPrEndIndex += "</w:sectPr>".length;
                        sectPr = xml.substring(sectPrStartIndex, sectPrEndIndex);
                    }
                }
            } else {
                // Insert a section break or page break after each file except the last one
                if (self._pageBreak && index < files.length - 1) {
                    if (sectPr) {
                        self.insertRaw('<w:p><w:pPr>' + sectPr + '</w:pPr></w:p>');
                    } else {
                        self.insertPageBreak();
                    }
                }
            }
        });
    };

    this.save = function (type, callback) {
        var zip = this._baseZip;

        var xml = zip.file("word/document.xml").asText();
        var startIndex = xml.indexOf("<w:body>") + 8;
        var endIndex = xml.lastIndexOf("<w:sectPr");

        xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(''));

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