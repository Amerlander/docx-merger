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
    
        // Parse the document.xml to find the section properties (<w:sectPr>) correctly
        var xmlString = zip.file("word/document.xml").asText();
        var xmlDoc = new DOMParser().parseFromString(xmlString, 'text/xml');
    
        // Find the position of <w:sectPr> in the XML (start and end)
        var sectPrStartIndex = xmlString.indexOf('<w:sectPr');
        if (sectPrStartIndex === -1) {
            console.error("Section properties <w:sectPr> not found in document.xml");
            return headerFooterRefs; // Early return if not found
        }
        var sectPrEndIndex = xmlString.indexOf('</w:sectPr>', sectPrStartIndex) + 11;
        
        if (sectPrEndIndex === -1) {
            console.error("End of section properties </w:sectPr> not found");
            return headerFooterRefs; // Early return if not found
        }
    
        var sectPrXML = xmlString.slice(sectPrStartIndex, sectPrEndIndex);
    
        // Parse the section properties XML to find header/footer references
        var sectDoc = new DOMParser().parseFromString(sectPrXML, 'text/xml');
    
        // Debugging: Output the section properties XML to check the structure
        console.log("sectPrXML:", sectPrXML, sectDoc);
    
        // Find the header reference (<w:headerReference>)
        var headerRef = sectDoc.getElementsByTagName('w:headerReference');
        if (headerRef.length > 0) {
            // Log header reference if found
            console.log("Header Reference Found:", headerRef[0].outerHTML);
            headerFooterRefs.push({
                type: 'headerReference',
                xml: headerRef[0].outerHTML
            });
        }
    
        // Find the footer reference (<w:footerReference>)
        var footerRef = sectDoc.getElementsByTagName('w:footerReference');
        if (footerRef.length > 0) {
            // Log footer reference if found
            console.log("Footer Reference Found:", footerRef[0].outerHTML);
            headerFooterRefs.push({
                type: 'footerReference',
                xml: footerRef[0].outerHTML
            });
        }

        console.log('REFS', headerFooterRefs)
    
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
