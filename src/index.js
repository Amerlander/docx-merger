import JSZip from 'jszip';
import * as Style from './merge-styles.js';
import * as Media from './merge-media.js';
import * as RelContentType from './merge-relations-and-content-type.js';
import * as bulletsNumbering from './merge-bullets-numberings.js';

// Check if running in a browser environment
const isBrowser = typeof window !== 'undefined' && typeof window.document !== 'undefined';

// Use the appropriate XML parser and serializer based on the environment
const XMLSerializer = isBrowser ? window.XMLSerializer : require('@xmldom/xmldom').XMLSerializer;
const DOMParser = isBrowser ? window.DOMParser : require('@xmldom/xmldom').DOMParser;

class DocxMerger {
    constructor () {
        this._body = [];
        this._header = [];
        this._footer = [];
        this._pageBreak = true;
        this._Basestyle = 'source';
        this._style = [];
        this._numbering = [];
        this._files = [];
        this._contentTypes = {};
        this._media = {};
        this._rel = {};
        this._builder = this._body;
    }

    async initialize(options = {}, files) {
        files = files || [];
        this._pageBreak = typeof options.pageBreak !== 'undefined' ? !!options.pageBreak : true;
        this._Basestyle = options.style || 'source';
    
        for (const file of files) {
            // Convert Uint8Array to ArrayBuffer if necessary
            const arrayBuffer = file instanceof Uint8Array ? file.buffer : file;
            console.log('arrayBuffer', arrayBuffer);
            const zip = await new JSZip().loadAsync(arrayBuffer);
            this._files.push(zip);
        }
    
        if (this._files.length > 0) {
            await this.mergeBody(this._files);
        }
    }

    insertPageBreak() {
        const pb = '<w:p> \
                    <w:r> \
                        <w:br w:type="page"/> \
                    </w:r> \
                </w:p>';
        this._builder.push(pb);
    }

    insertSectionBreak() {
        const sb = '<w:p> \
                    <w:pPr> \
                        <w:sectPr> \
                            <w:type w:val="nextPage"/> \
                        </w:sectPr> \
                    </w:pPr> \
                </w:p>';
        this._builder.push(sb);
    }

    insertRaw(xml) {
        this._builder.push(xml);
    }

    async mergeBody(files) {
        this._builder = this._body;
        await RelContentType.mergeContentTypes(files, this._contentTypes);
        await Media.prepareMediaFiles(files, this._media);
        await RelContentType.mergeRelations(files, this._rel);
        await bulletsNumbering.prepareNumbering(files, this._numbering);
        await bulletsNumbering.mergeNumbering(files, this._numbering);
        await Style.prepareStyles(files, this._style);
        await Style.mergeStyles(files, this._style);
        const merge = files.map(async(zip, index) => {
            let xmlString = await zip.file('word/document.xml').async('string');
            xmlString = xmlString.substring(xmlString.indexOf('<w:body>') + 8);
            xmlString = xmlString.substring(0, xmlString.indexOf('</w:body>'));
            xmlString = xmlString.substring(0, xmlString.lastIndexOf('<w:sectPr'));
            this.insertRaw(xmlString);
            if (this._pageBreak && index < files.length-1){
                this.insertSectionBreak();
                // this.insertPageBreak();
            }
        });
        return Promise.all(merge).then(() => {});
    }

    async save(type, callback) {
        const zip = this._files[0];
        
        if (!zip || !zip.file) {
            throw new Error('JSZip file not properly loaded');
        }
    
        let xmlString = await zip.file('word/document.xml').async('string');
    
        const startIndex = xmlString.indexOf('<w:body>') + 8;
        const endIndex = xmlString.lastIndexOf('<w:sectPr');
    
        xmlString = xmlString.replace(xmlString.slice(startIndex, endIndex), this._body.join(''));
    
        await RelContentType.generateContentTypes(zip, this._contentTypes);
        await Media.copyMediaFiles(zip, this._media, this._files);
        await RelContentType.generateRelations(zip, this._rel);
        await bulletsNumbering.generateNumbering(zip, this._numbering);
        await Style.generateStyles(zip, this._style);
    
        zip.file('word/document.xml', xmlString);
    
        const generatedFile = await zip.generateAsync({
            type: type,
            compression: 'DEFLATE',
            compressionOptions: {
                level: 4
            }
        });

        if (callback) {
            callback(generatedFile);
        }

        return generatedFile;
    }
}

export default DocxMerger;
