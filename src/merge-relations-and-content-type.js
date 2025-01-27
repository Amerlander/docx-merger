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

var generateContentTypes = function (zip, _contentTypes, headers, footers) {
    var xmlString = zip.file("[Content_Types].xml").asText();
    var xml = new DOMParser().parseFromString(xmlString, "text/xml");
    var serializer = new XMLSerializer();

    var types = xml.documentElement.cloneNode();

    // Add merged content types
    for (var node in _contentTypes) {
        types.appendChild(_contentTypes[node]);
    }

    // Add header content types
    headers.forEach((_, index) => {
        const override = xml.createElement("Override");
        override.setAttribute("PartName", `/word/header${index + 1}.xml`);
        override.setAttribute(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"
        );
        types.appendChild(override);
    });

    // Add footer content types
    footers.forEach((_, index) => {
        const override = xml.createElement("Override");
        override.setAttribute("PartName", `/word/footer${index + 1}.xml`);
        override.setAttribute(
            "ContentType",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"
        );
        types.appendChild(override);
    });

    var startIndex = xmlString.indexOf("<Types");
    xmlString = xmlString.replace(
        xmlString.slice(startIndex),
        serializer.serializeToString(types)
    );

    zip.file("[Content_Types].xml", xmlString);
};

var generateRelations = function (zip, _rel, headers, footers) {
    var xmlString = zip.file("word/_rels/document.xml.rels").asText();
    var xml = new DOMParser().parseFromString(xmlString, "text/xml");
    var serializer = new XMLSerializer();

    var relationships = xml.documentElement.cloneNode();

    // Add merged relationships
    for (var node in _rel) {
        relationships.appendChild(_rel[node]);
    }

    // Add header relationships
    headers.forEach((_, index) => {
        const relationship = xml.createElement("Relationship");
        relationship.setAttribute("Id", `rIdHeader${index + 1}`);
        relationship.setAttribute(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
        );
        relationship.setAttribute("Target", `header${index + 1}.xml`);
        relationships.appendChild(relationship);
    });

    // Add footer relationships
    footers.forEach((_, index) => {
        const relationship = xml.createElement("Relationship");
        relationship.setAttribute("Id", `rIdFooter${index + 1}`);
        relationship.setAttribute(
            "Type",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
        );
        relationship.setAttribute("Target", `footer${index + 1}.xml`);
        relationships.appendChild(relationship);
    });

    var startIndex = xmlString.indexOf("<Relationships");
    xmlString = xmlString.replace(
        xmlString.slice(startIndex),
        serializer.serializeToString(relationships)
    );

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
