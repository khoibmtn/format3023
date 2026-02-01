import { readFile, writeFile } from 'fs/promises';
import { processDocx } from './docx/processor.js';
import JSZip from 'jszip';

async function test() {
    // Read the test_download.docx (which is an already processed file)
    // We need an original file to test properly
    // Let's create a simple test by processing the existing file again
    console.log('=== DOCX Processing Verification Test ===\n');

    const testFile = './test_download.docx';
    console.log('Reading:', testFile);

    const inputBuffer = await readFile(testFile);
    console.log('Input size:', inputBuffer.length);

    // Process the file
    console.log('\nProcessing...');
    const outputBuffer = await processDocx(inputBuffer);
    console.log('Output size:', outputBuffer.length);

    // Extract and compare document.xml
    const inputZip = await JSZip.loadAsync(inputBuffer);
    const outputZip = await JSZip.loadAsync(outputBuffer);

    const inputDoc = await inputZip.file('word/document.xml').async('string');
    const outputDoc = await outputZip.file('word/document.xml').async('string');

    console.log('\n=== Input document.xml ===');
    console.log('Length:', inputDoc.length);

    console.log('\n=== Output document.xml ===');
    console.log('Length:', outputDoc.length);

    // Check for specific formatting attributes in output
    console.log('\n=== Formatting Verification ===');

    // Check for line spacing attribute
    const lineMatch = outputDoc.match(/w:line="(\d+)"/);
    console.log('w:line attribute:', lineMatch ? lineMatch[1] : 'NOT FOUND');

    // Check for before spacing
    const beforeMatch = outputDoc.match(/w:before="(\d+)"/);
    console.log('w:before attribute:', beforeMatch ? beforeMatch[1] : 'NOT FOUND');

    // Check for after spacing
    const afterMatch = outputDoc.match(/w:after="(\d+)"/);
    console.log('w:after attribute:', afterMatch ? afterMatch[1] : 'NOT FOUND');

    // Check for right indent
    const rightMatch = outputDoc.match(/w:right="(\d+)"/);
    console.log('w:right attribute:', rightMatch ? rightMatch[1] : 'NOT FOUND');

    // Check for widowControl
    console.log('w:widowControl with val:', outputDoc.includes('w:widowControl') && outputDoc.includes('w:val="1"'));

    // Show first paragraph with w:pPr and w:spacing
    console.log('\n=== Sample formatted paragraph ===');
    const spacingParagraph = outputDoc.match(/<w:p[^>]*>.*?<w:spacing[^>]*>.*?<\/w:p>/s);
    if (spacingParagraph) {
        console.log(spacingParagraph[0].substring(0, 800));
    } else {
        console.log('No paragraph with w:spacing found!');

        // Try to find any w:spacing
        const spacingMatch = outputDoc.match(/<w:spacing[^>]*>/);
        if (spacingMatch) {
            console.log('w:spacing found:', spacingMatch[0]);
        }
    }

    // Save the output for manual inspection
    await writeFile('./test_verify_output.docx', outputBuffer);
    console.log('\nSaved to: ./test_verify_output.docx');

    // Also save just the document.xml for inspection
    await writeFile('./document_output.xml', outputDoc);
    console.log('Saved XML to: ./document_output.xml');

    console.log('\n=== COMPARE: Input vs Output (first w:spacing) ===');
    const inputSpacing = inputDoc.match(/<w:spacing[^/>]*\/?>/);
    const outputSpacing = outputDoc.match(/<w:spacing[^/>]*\/?>/);
    console.log('Input spacing:', inputSpacing ? inputSpacing[0] : 'none');
    console.log('Output spacing:', outputSpacing ? outputSpacing[0] : 'none');
}

test().catch(console.error);
