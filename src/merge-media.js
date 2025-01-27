import JSZip from 'jszip';
import { DOMParser, XMLSerializer } from '@xmldom/xmldom';

async function prepareMediaFiles(files, media) {
    for (const file of files) {
        const zip = await new JSZip().loadAsync(file);
        const mediaFiles = zip.folder('word/media');
        if (mediaFiles) {
            mediaFiles.forEach((relativePath, file) => {
                media[relativePath] = file;
            });
        }
    }
}

const updateMediaRelations = async function(zip, count, _media) {
    let xmlString = await zip.file('word/_rels/document.xml.rels').async('string');
    let xml = new DOMParser().parseFromString(xmlString, 'text/xml');

    const childNodes = xml.getElementsByTagName('Relationships')[0].childNodes;
    const serializer = new XMLSerializer();

    for (const node in childNodes) {
        if (/^\d+$/.test(node) && childNodes[node].getAttribute) {
            const target = childNodes[node].getAttribute('Target');
            if ('word/' + target === _media[count].oldTarget) {
                _media[count].oldRelID = childNodes[node].getAttribute('Id');
                childNodes[node].setAttribute('Target', _media[count].newTarget);
                childNodes[node].setAttribute('Id', _media[count].oldRelID + '_' + count);
            }
        }
    }

    const startIndex = xmlString.indexOf('<Relationships');
    xmlString = xmlString.replace(xmlString.slice(startIndex), serializer.serializeToString(xml.documentElement));

    zip.file('word/_rels/document.xml.rels', xmlString);
};

const updateMediaContent = async function(zip, count, _media) {
    let xmlString = await zip.file('word/document.xml').async('string');
    xmlString = xmlString.replace(new RegExp(_media[count].oldRelID + '"', 'g'), _media[count].oldRelID + '_' + count + '"');
    zip.file('word/document.xml', xmlString);
};

async function copyMediaFiles(zip, media, files) {
    const mediaFolder = zip.folder('word/media');
    if (!mediaFolder) {
        throw new Error('Media folder not found in the zip');
    }

    for (const [relativePath, file] of Object.entries(media)) {
        const content = await file.async('blob');
        mediaFolder.file(relativePath, content);
    }
}

export { prepareMediaFiles, updateMediaRelations, updateMediaContent, copyMediaFiles };