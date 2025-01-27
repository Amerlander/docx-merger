import DocxMerger from './../src/index';

import fs from 'fs';
import path from 'path';

(async () => {
    const file1 = fs.readFileSync(path.resolve(__dirname, 'template-0.docx'), 'binary');
    const file2 = fs.readFileSync(path.resolve(__dirname, 'template-1.docx'), 'binary');
    const docx = new DocxMerger();
    await docx.initialize({},[file1,file2]);
    //SAVING THE DOCX FILE
    const data = await docx.save('nodebuffer');
    fs.writeFileSync(path.resolve(__dirname, "output.docx"), data);
})()



