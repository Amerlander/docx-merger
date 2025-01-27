var XMLSerializer = require('xmldom').XMLSerializer;
var DOMParser = require('xmldom').DOMParser;

var mergeContentTypes = function (files, _contentTypes) {
    files.forEach(function (zip) {
        var xmlString = zip.file("[Content_Types].xml").asText();
        var xml = new DOMParser().parseFromString(xmlString, "text/xml");

        var childNodes = xml.getElementsByTagName("Types")[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var contentType = childNodes[node].getAttribute("ContentType");
                if (!_contentTypes[contentType]) {
                    _contentTypes[contentType] = childNodes[node].cloneNode();
                }
            }
        }
    });
};

var mergeRelations = function (files, _rel) {
    files.forEach(function (zip) {
        var xmlString = zip.file("word/_rels/document.xml.rels").asText();
        var xml = new DOMParser().parseFromString(xmlString, "text/xml");

        var childNodes = xml.getElementsByTagName("Relationships")[0].childNodes;

        for (var node in childNodes) {
            if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
                var Id = childNodes[node].getAttribute("Id");
                if (!_rel[Id]) {
                    _rel[Id] = childNodes[node].cloneNode();
                }
            }
        }
    });
};

var generateContentTypes = function(zip, _contentTypes, _headerFooterRelationships) {
    var xmlString = zip.file("[Content_Types].xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    // Add the ContentType for headers and footers if any
    _headerFooterRelationships.forEach(function(rel) {
        var contentTypeNode = xml.createElement('Default');
        contentTypeNode.setAttribute('Extension', 'xml');
        contentTypeNode.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml');
        types.appendChild(contentTypeNode);

        contentTypeNode = xml.createElement('Default');
        contentTypeNode.setAttribute('Extension', 'xml');
        contentTypeNode.setAttribute('ContentType', 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml');
        types.appendChild(contentTypeNode);
    });

    // Add all other content types
    for (var node in _contentTypes) {
        types.appendChild(_contentTypes[node]);
    }

    var startIndex = xmlString.indexOf("<Types");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("[Content_Types].xml", xmlString);
};

var generateRelations = function(zip, _rel, _headerFooterRefs) {
    var xmlString = zip.file("word/_rels/document.xml.rels").asText();
    var xml = new DOMParser().parseFromString(xmlString, 'text/xml');
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    // Add relationships for headers and footers if any
    _headerFooterRefs.forEach(function(ref) {
        if (ref.type === "headerReference") {
            var relNode = xml.createElement('Relationship');
            relNode.setAttribute('Id', ref.xml);
            relNode.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/header');
            relNode.setAttribute('Target', 'headers/' + ref.xml); // Assuming the header XML is located in the headers folder
            types.appendChild(relNode);
        }
        if (ref.type === "footerReference") {
            var relNode = xml.createElement('Relationship');
            relNode.setAttribute('Id', ref.xml);
            relNode.setAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer');
            relNode.setAttribute('Target', 'footers/' + ref.xml); // Assuming the footer XML is located in the footers folder
            types.appendChild(relNode);
        }
    });

    // Add all other relationships
    for (var node in _rel) {
        types.appendChild(_rel[node]);
    }

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(types));

    zip.file("word/_rels/document.xml.rels", xmlString);
};

var generateHeaders = function (zip, headers) {
    headers.forEach((content, index) => {
        zip.file(`word/header${index + 1}.xml`, content);
    });
};

var generateFooters = function (zip, footers) {
    footers.forEach((content, index) => {
        zip.file(`word/footer${index + 1}.xml`, content);
    });
};

// Main export
module.exports = {
    mergeContentTypes: mergeContentTypes,
    mergeRelations: mergeRelations,
    generateContentTypes: generateContentTypes,
    generateRelations: generateRelations,
    generateHeaders: generateHeaders,
    generateFooters: generateFooters,
};
