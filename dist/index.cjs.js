var F=Object.create;var g=Object.defineProperty;var z=Object.getOwnPropertyDescriptor;var A=Object.getOwnPropertyNames;var D=Object.getPrototypeOf,L=Object.prototype.hasOwnProperty;var k=(i,e)=>{for(var t in e)g(i,t,{get:e[t],enumerable:!0})},f=(i,e,t,r)=>{if(e&&typeof e=="object"||typeof e=="function")for(let n of A(e))!L.call(i,n)&&n!==t&&g(i,n,{get:()=>e[n],enumerable:!(r=z(e,n))||r.enumerable});return i};var y=(i,e,t)=>(t=i!=null?F(D(i)):{},f(e||!i||!i.__esModule?g(t,"default",{value:i,enumerable:!0}):t,i)),X=i=>f(g({},"__esModule",{value:!0}),i);var $={};k($,{default:()=>q});module.exports=X($);var E=y(require("jszip"));var d=require("@xmldom/xmldom"),u=async function(i,e){let t=i.map(async r=>{let n=await r.file("word/styles.xml").async("string"),l=new d.DOMParser().parseFromString(n,"text/xml").documentElement.childNodes;for(let s=0;s<l.length;s++)if(l[s].nodeType===1){let a=l[s].getAttribute("w:styleId");e[a]||(e[a]=l[s].cloneNode(!0))}});return Promise.all(t)},x=async function(i,e){let t=i.map(async r=>{let n=await r.file("word/styles.xml").async("string"),o=new d.DOMParser().parseFromString(n,"text/xml"),l=new d.XMLSerializer,s=o.documentElement.cloneNode();for(let w in e)s.appendChild(e[w]);let a=n.indexOf("<w:styles");n=n.replace(n.slice(a),l.serializeToString(s)),r.file("word/styles.xml",n)});return Promise.all(t)};var h=async function(i,e){let t=await i.file("word/styles.xml").async("string"),r=t.indexOf("<w:style "),n=t.indexOf("</w:styles>");t=t.replace(t.slice(r,n),e.join("")),i.file("word/styles.xml",t)};var b=y(require("jszip")),S=require("@xmldom/xmldom");async function _(i,e){for(let t of i){let n=(await new b.default().loadAsync(t)).folder("word/media");n&&n.forEach((o,l)=>{e[o]=l})}}async function N(i,e,t){let r=i.folder("word/media");if(!r)throw new Error("Media folder not found in the zip");for(let[n,o]of Object.entries(e)){let l=await o.async("blob");r.file(n,l)}}var c=require("@xmldom/xmldom"),T=function(i,e){let t=i.map(async r=>{let n=await r.file("[Content_Types].xml").async("string"),l=new c.DOMParser().parseFromString(n,"text/xml").getElementsByTagName("Types")[0].childNodes;for(let s in l)if(/^\d+$/.test(s)&&l[s].getAttribute){let a=l[s].getAttribute("ContentType");e[a]||(e[a]=l[s].cloneNode())}});return Promise.all(t)},I=async function(i,e){let t=i.map(async r=>{let n=await r.file("word/_rels/document.xml.rels").async("string"),l=new c.DOMParser().parseFromString(n,"text/xml").documentElement.childNodes;for(let s=0;s<l.length;s++)if(l[s].nodeType===1){let a=l[s].getAttribute("Id");e[a]||(e[a]=l[s].cloneNode())}});return Promise.all(t)},M=async function(i,e){let t=await i.file("[Content_Types].xml").async("string"),r=new c.DOMParser().parseFromString(t,"text/xml"),n=new c.XMLSerializer,o=r.documentElement.cloneNode();for(let s in e)o.appendChild(e[s]);let l=t.indexOf("<Types");t=t.replace(t.slice(l),n.serializeToString(o)),i.file("[Content_Types].xml",t)},O=async function(i,e){let t=await i.file("word/_rels/document.xml.rels").async("string"),r=new c.DOMParser().parseFromString(t,"text/xml"),n=new c.XMLSerializer,o=r.documentElement.cloneNode();for(let s in e)o.appendChild(e[s]);let l=t.indexOf("<Relationships");t=t.replace(t.slice(l),n.serializeToString(o)),i.file("word/_rels/document.xml.rels",t)};var G=require("jszip"),m=require("@xmldom/xmldom");async function P(i,e){let t=i.map(async r=>{let n=await r.file("word/numbering.xml").async("string"),l=new m.DOMParser().parseFromString(n,"text/xml").documentElement.childNodes;for(let s=0;s<l.length;s++)if(l[s].nodeType===1){let a=l[s].getAttribute("w:abstractNumId");e[a]||(e[a]=l[s].cloneNode(!0))}e.push(n)});return Promise.all(t)}async function R(i,e){let t=i.map(async r=>{let n=await r.file("word/numbering.xml").async("string"),o=new m.DOMParser().parseFromString(n,"text/xml"),l=new m.XMLSerializer,s=o.documentElement.cloneNode();for(let w in e)s.appendChild(e[w]);let a=n.indexOf("<w:numbering");n=n.replace(n.slice(a),l.serializeToString(s)),r.file("word/numbering.xml",n)});return Promise.all(t)}async function C(i,e){let t=i.file("word/numbering.xml");if(!t)throw new Error("Numbering file not found in the zip");let r=await t.async("string"),n=r.indexOf("<w:abstractNum "),o=r.indexOf("</w:numbering>");r=r.replace(r.slice(n,o),e.join("")),i.file("word/numbering.xml",r)}var B=typeof window!="undefined"&&typeof window.document!="undefined",K=B?window.XMLSerializer:require("@xmldom/xmldom").XMLSerializer,Q=B?window.DOMParser:require("@xmldom/xmldom").DOMParser,p=class{constructor(){this._body=[],this._header=[],this._footer=[],this._pageBreak=!0,this._Basestyle="source",this._style=[],this._numbering=[],this._files=[],this._contentTypes={},this._media={},this._rel={},this._builder=this._body}async initialize(e={},t){t=t||[],this._pageBreak=typeof e.pageBreak!="undefined"?!!e.pageBreak:!0,this._Basestyle=e.style||"source";for(let r of t){let n=r instanceof Uint8Array?r.buffer:r,o=await new E.default().loadAsync(n);this._files.push(o)}this._files.length>0&&await this.mergeBody(this._files)}insertPageBreak(){let e='<w:p>                     <w:r>                         <w:br w:type="page"/>                     </w:r>                 </w:p>';this._builder.push(e)}insertSectionBreak(){let e='<w:p>                     <w:pPr>                         <w:sectPr>                             <w:type w:val="nextPage"/>                         </w:sectPr>                     </w:pPr>                 </w:p>';this._builder.push(e)}insertRaw(e){this._builder.push(e)}async mergeBody(e){this._builder=this._body,await T(e,this._contentTypes),await _(e,this._media),await I(e,this._rel),await P(e,this._numbering),await R(e,this._numbering),await u(e,this._style),await x(e,this._style);let t=e.map(async(r,n)=>{let o=await r.file("word/document.xml").async("string");o=o.substring(o.indexOf("<w:body>")+8),o=o.substring(0,o.indexOf("</w:body>")),o=o.substring(0,o.lastIndexOf("<w:sectPr")),this.insertRaw(o),this._pageBreak&&n<e.length-1&&this.insertSectionBreak()});return Promise.all(t).then(()=>{})}async save(e,t){let r=this._files[0];if(!r||!r.file)throw new Error("JSZip file not properly loaded");let n=await r.file("word/document.xml").async("string"),o=n.indexOf("<w:body>")+8,l=n.lastIndexOf("<w:sectPr");n=n.replace(n.slice(o,l),this._body.join("")),await M(r,this._contentTypes),await N(r,this._media,this._files),await O(r,this._rel),await C(r,this._numbering),await h(r,this._style),r.file("word/document.xml",n);let s=await r.generateAsync({type:e,compression:"DEFLATE",compressionOptions:{level:4}});return t&&t(s),s}},q=p;
//# sourceMappingURL=index.cjs.js.map
