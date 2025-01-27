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
    
        // Parse the XML to extract <w:hdrReference> and <w:ftrReference> from document.xml
        var xmlString = zip.file("word/document.xml").asText();
        var xmlDoc = new DOMParser().parseFromString(xmlString, 'text/xml');
    
        // Look for header and footer references inside <w:sectPr>
        var sectPrStartIndex = xmlString.indexOf('<w:sectPr');
        if (sectPrStartIndex !== -1) {
            var sectPrXML = xmlString.slice(sectPrStartIndex, xmlString.indexOf('</w:sectPr>') + 11);
            var sectDoc = new DOMParser().parseFromString(sectPrXML, 'text/xml');
    
            // Extract header reference
            var headerRef = sectDoc.getElementsByTagName('w:hdrReference');
            if (headerRef.length > 0) {
                headerFooterRefs.push({
                    type: 'headerReference',
                    xml: headerRef[0].outerHTML
                });
            }
    
            // Extract footer reference
            var footerRef = sectDoc.getElementsByTagName('w:ftrReference');
            if (footerRef.length > 0) {
                headerFooterRefs.push({
                    type: 'footerReference',
                    xml: footerRef[0].outerHTML
                });
            }
        }
    
        return headerFooterRefs;
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
    },
};
