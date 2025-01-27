import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

async function prepareNumbering(files, _numbering) {
    const merge = files.map(async (zip) => {
        let xmlString = await zip.file('word/numbering.xml').async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const childNodes = xml.documentElement.childNodes;

        for (let node = 0; node < childNodes.length; node++) {
            if (childNodes[node].nodeType === 1) { // Element node
                const abstractNumId = childNodes[node].getAttribute('w:abstractNumId');
                if (!_numbering[abstractNumId])
                    _numbering[abstractNumId] = childNodes[node].cloneNode(true);
            }
        }
        _numbering.push(xmlString);
    });
    return Promise.all(merge);
}

async function mergeNumbering(files, _numbering) {
    const merge = files.map(async (zip) => {
        let xmlString = await zip.file('word/numbering.xml').async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const serializer = new XMLSerializer();

        const numbering = xml.documentElement.cloneNode();

        for (const node in _numbering) {
            numbering.appendChild(_numbering[node]);
        }

        const startIndex = xmlString.indexOf('<w:numbering');
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(numbering));

        zip.file('word/numbering.xml', xmlString);
    });
    return Promise.all(merge);
}

async function generateNumbering(zip, _numbering) {
    let xmlBin = zip.file("word/numbering.xml");
    if (!xmlBin) {
        throw new Error('Numbering file not found in the zip');
    }

    let xmlString = await xmlBin.async('string');
    const startIndex = xmlString.indexOf("<w:abstractNum ");
    const endIndex = xmlString.indexOf("</w:numbering>");

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), _numbering.join(''));

    zip.file("word/numbering.xml", xmlString);
}

export { prepareNumbering, mergeNumbering, generateNumbering };