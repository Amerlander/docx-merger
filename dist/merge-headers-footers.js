var DOMParser = require('xmldom').DOMParser;
var XMLSerializer = require('xmldom').XMLSerializer;

module.exports = {
    /**
     * Extract header and footer references from section properties (<w:sectPr>)
     * @param {JSZip} zip - The zip object of the Word document
     * @param {string} sectPr - The section properties XML
     * @returns {Array} An array of header and footer references
     */
    extractHeadersFooters: function (zip, sectPr) {
        var headerFooterRefs = [];
        var doc = new DOMParser().parseFromString(sectPr, 'text/xml');

        // Extract header references
        var headerRefs = doc.getElementsByTagName('w:headerReference');
        for (var i = 0; i < headerRefs.length; i++) {
            var refId = headerRefs[i].getAttribute('r:id');
            var headerFile = zip.file(`word/${refId}.xml`);
            if (headerFile) {
                headerFooterRefs.push({
                    type: 'headerReference',
                    xml: headerFile.asText()
                });
            }
        }

        // Extract footer references
        var footerRefs = doc.getElementsByTagName('w:footerReference');
        for (var i = 0; i < footerRefs.length; i++) {
            var refId = footerRefs[i].getAttribute('r:id');
            var footerFile = zip.file(`word/${refId}.xml`);
            if (footerFile) {
                headerFooterRefs.push({
                    type: 'footerReference',
                    xml: footerFile.asText()
                });
            }
        }

        return headerFooterRefs || [];
    },

    /**
     * Generate header files in the zip object
     * @param {JSZip} zip - The zip object of the Word document
     * @param {Array} headers - An array of header XML content
     */
    generateHeaders: function (zip, headers) {
        headers.forEach((header, index) => {
            zip.file(`word/header${index + 1}.xml`, header);
        });
    },

    /**
     * Generate footer files in the zip object
     * @param {JSZip} zip - The zip object of the Word document
     * @param {Array} footers - An array of footer XML content
     */
    generateFooters: function (zip, footers) {
        footers.forEach((footer, index) => {
            zip.file(`word/footer${index + 1}.xml`, footer);
        });
    }
};