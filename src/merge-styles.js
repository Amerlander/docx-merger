import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

const prepareStyles = async function(files, _style) {
    const merge = files.map(async (zip) => {
        let xmlString = await zip.file('word/styles.xml').async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const childNodes = xml.documentElement.childNodes;

        for (let node = 0; node < childNodes.length; node++) {
            if (childNodes[node].nodeType === 1) { // Element node
                const styleId = childNodes[node].getAttribute('w:styleId');
                if (!_style[styleId])
                    _style[styleId] = childNodes[node].cloneNode(true);
            }
        }
    });
    return Promise.all(merge);
};

const mergeStyles = async function(files, _style) {
    const merge = files.map(async (zip) => {
        let xmlString = await zip.file('word/styles.xml').async('string');
        const xml = new DOMParser().parseFromString(xmlString, 'text/xml');
        const serializer = new XMLSerializer();

        const styles = xml.documentElement.cloneNode();

        for (const node in _style) {
            styles.appendChild(_style[node]);
        }

        const startIndex = xmlString.indexOf('<w:styles');
        xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(styles));

        zip.file('word/styles.xml', xmlString);
    });
    return Promise.all(merge);
};

const updateStyleRel_Content = async function(zip, fileIndex, styleId) {
    let xmlString = await zip.file('word/document.xml').async('string');
    xmlString = xmlString.replace(new RegExp('w:val="' + styleId + '"', 'g'), 'w:val="' + styleId + '_' + fileIndex + '"');
    zip.file('word/document.xml', xmlString);
};

const generateStyles = async function(zip, _style) {
    let xmlString = await zip.file('word/styles.xml').async('string');
    const startIndex = xmlString.indexOf('<w:style ');
    const endIndex = xmlString.indexOf('</w:styles>');

    xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), _style.join(''));

    zip.file('word/styles.xml', xmlString);
};

export { prepareStyles, mergeStyles, updateStyleRel_Content, generateStyles };