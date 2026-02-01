import { readFile, writeFile } from 'fs/promises';
import { processDocx } from './docx/processor.js';

async function test() {
    const inputPath = process.argv[2] || './40. Phong bế mặt phẳng cơ ngực lớn.docx';
    const outputPath = inputPath.replace(/\.docx$/i, '_processed.docx');

    console.log('Reading:', inputPath);
    const inputBuffer = await readFile(inputPath);
    console.log('Input size:', inputBuffer.length);

    console.log('Processing...');
    const outputBuffer = await processDocx(inputBuffer);
    console.log('Output size:', outputBuffer.length);

    console.log('Writing:', outputPath);
    await writeFile(outputPath, outputBuffer);
    console.log('Done!');

    // Verify the output is a valid ZIP
    const JSZip = (await import('jszip')).default;
    try {
        const zip = await JSZip.loadAsync(outputBuffer);
        console.log('Output is valid ZIP with files:', Object.keys(zip.files).length);

        // Check document.xml
        const docXml = await zip.file('word/document.xml').async('string');
        console.log('document.xml starts with:', docXml.substring(0, 100));
    } catch (err) {
        console.error('Output is NOT a valid ZIP:', err.message);
    }
}

test().catch(console.error);
