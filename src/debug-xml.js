import { readFile, writeFile } from 'fs/promises';
import JSZip from 'jszip';

async function debug() {
    const testFile = process.argv[2] || './test_download.docx';

    console.log('=== DEBUG DOCX Processing ===\n');
    console.log('Reading:', testFile);

    const buffer = await readFile(testFile);
    console.log('Size:', buffer.length, 'bytes');

    // Extract and show document.xml sample
    const zip = await JSZip.loadAsync(buffer);
    const docXml = await zip.file('word/document.xml').async('string');
    const stylesXml = await zip.file('word/styles.xml').async('string');

    console.log('\n=== document.xml ===');
    console.log('Length:', docXml.length);
    console.log('First 2000 chars:\n');
    console.log(docXml.substring(0, 2000));

    console.log('\n=== Sample w:p (paragraph) ===');
    // Find first paragraph with w:pPr
    const pPrMatch = docXml.match(/<w:p[^>]*>[\s\S]*?<w:pPr[\s\S]*?<\/w:pPr>[\s\S]*?<\/w:p>/);
    if (pPrMatch) {
        console.log(pPrMatch[0].substring(0, 1500));
    } else {
        console.log('No w:pPr found in paragraphs!');
        // Find first paragraph
        const pMatch = docXml.match(/<w:p[^>]*>[\s\S]*?<\/w:p>/);
        if (pMatch) {
            console.log('First w:p:\n', pMatch[0].substring(0, 1000));
        }
    }

    console.log('\n=== styles.xml sample ===');
    console.log('Length:', stylesXml.length);

    // Check for w:docDefaults
    const docDefaultsMatch = stylesXml.match(/<w:docDefaults[\s\S]*?<\/w:docDefaults>/);
    if (docDefaultsMatch) {
        console.log('\nw:docDefaults:\n', docDefaultsMatch[0].substring(0, 1000));
    } else {
        console.log('No w:docDefaults found!');
    }

    // Check for Normal style
    const normalStyleMatch = stylesXml.match(/<w:style[^>]*w:styleId="Normal"[\s\S]*?<\/w:style>/);
    if (normalStyleMatch) {
        console.log('\n=== Normal style ===');
        console.log(normalStyleMatch[0].substring(0, 1000));
    }

    // Check for specific formatting attributes
    console.log('\n=== Formatting Check ===');
    console.log('Contains w:spacing:', docXml.includes('<w:spacing'));
    console.log('Contains w:ind:', docXml.includes('<w:ind'));
    console.log('Contains w:widowControl:', docXml.includes('<w:widowControl'));
    console.log('Contains w:line=', docXml.includes('w:line='));
    console.log('Contains w:before=', docXml.includes('w:before='));
    console.log('Contains w:after=', docXml.includes('w:after='));

    // Count paragraphs
    const paragraphCount = (docXml.match(/<w:p[ >]/g) || []).length;
    console.log('\nParagraph count:', paragraphCount);
}

debug().catch(console.error);
