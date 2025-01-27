var JSZip = require('jszip');
var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

var Style = require('./merge-styles');
var Media = require('./merge-media');
var RelContentType = require('./merge-relations-and-content-type');
var bulletsNumbering = require('./merge-bullets-numberings');
var headersFooters = require('./merge-headers-footers');

function DocxMerger(options, files) {

    this._body = [];
    this._header = [];
    this._footer = [];
    this._Basestyle = options.style || 'source';
    this._style = [];
    this._numbering = [];
    this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
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

    this.insertHeadersAndFooters = function (headerFooterRefs) {
        headerFooterRefs.forEach((ref) => {
            // Insert the header/footer XML into the appropriate place
            if (ref.type === "headerReference") {
                this._header.push(ref.xml);
            } else if (ref.type === "footerReference") {
                this._footer.push(ref.xml);
            }
        });
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
    
        files.forEach(function (zip, index) {
            // Extract document XML
            var xml = zip.file("word/document.xml").asText();
    
            // Extract <w:body> content
            var bodyStartIndex = xml.indexOf("<w:body>") + 8;
            var bodyEndIndex = xml.lastIndexOf("<w:sectPr");
    
            // Extract section properties (<w:sectPr>)
            var sectPrStartIndex = xml.lastIndexOf("<w:sectPr");
            var sectPrEndIndex = xml.indexOf("</w:sectPr>", sectPrStartIndex) + 11;
            var sectPr = xml.slice(sectPrStartIndex, sectPrEndIndex);
    
            // Extract content inside <w:body> excluding <w:sectPr>
            var content = xml.slice(bodyStartIndex, bodyEndIndex);
    
            // Merge headers and footers
            var headerFooterRefs = headersFooters.extractHeadersFooters(zip, sectPr);
            self.insertHeadersAndFooters(headerFooterRefs);
    
            // Append content to the builder
            self.insertRaw(content);
    
            // Append the section properties for this section
            if (self._pageBreak && index < files.length - 1) {
                self.insertRaw(`<w:p><w:pPr>${sectPr}</w:pPr></w:p>`);
            }
        });
    };
    

    this.save = function(type, callback) {
        var zip = this._files[0];
    
        var xml = zip.file("word/document.xml").asText();
        var startIndex = xml.indexOf("<w:body>") + 8;
        var endIndex = xml.lastIndexOf("<w:sectPr");
    
        xml = xml.replace(xml.slice(startIndex, endIndex), this._body.join(''));
    
        RelContentType.generateContentTypes(zip, this._contentTypes);
        Media.copyMediaFiles(zip, this._media, this._files);
        RelContentType.generateRelations(zip, this._rel);
        bulletsNumbering.generateNumbering(zip, this._numbering);
        Style.generateStyles(zip, this._style);
    
        // Generate header and footer files
        headersFooters.generateHeaders(zip, this._header);
        headersFooters.generateFooters(zip, this._footer);
    
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
