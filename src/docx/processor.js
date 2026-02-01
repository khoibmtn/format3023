import { extractDocx, repackageDocx, getDocumentXml, getStylesXml } from './extractor.js';
import { replaceText } from './text-replacer.js';
import { formatSpecialParagraphs, formatTitleParagraph } from './special-formatter.js';
import { formatParagraphs } from './paragraph-formatter.js';
import { formatRuns } from './run-formatter.js';
import { normalizeFonts } from './font-normalizer.js';

/**
 * Process a DOCX file with all transformations
 * 
 * @param {Buffer|ArrayBuffer} inputBuffer - Input DOCX file
 * @returns {Promise<Buffer>} Processed DOCX file
 */
export async function processDocx(inputBuffer) {
    console.log('[PROCESSOR] Starting DOCX processing...');

    // Extract DOCX contents
    const { zip, files } = await extractDocx(inputBuffer);
    console.log('[PROCESSOR] Extracted', Object.keys(files).length, 'files');

    const modifiedFiles = {};

    // 1. Process document.xml
    const documentXml = getDocumentXml(files);
    if (documentXml) {
        console.log('[PROCESSOR] document.xml length:', documentXml.length);
        let processedDoc = documentXml;

        try {
            // Apply text replacements (line breaks → paragraph breaks, en dash → hyphen)
            processedDoc = replaceText(processedDoc);
            console.log('[PROCESSOR] After text replace:', processedDoc.length);

            // Apply special paragraph formatting (title, blank lines before sections)
            processedDoc = formatSpecialParagraphs(processedDoc);
            console.log('[PROCESSOR] After special format:', processedDoc.length);

            // Apply paragraph formatting
            processedDoc = formatParagraphs(processedDoc);
            console.log('[PROCESSOR] After paragraph format:', processedDoc.length);

            // Apply title formatting (AFTER paragraph formatting to avoid being overwritten)
            processedDoc = formatTitleParagraph(processedDoc);
            console.log('[PROCESSOR] After title format:', processedDoc.length);

            // Apply run formatting (font 14pt, Times New Roman)
            processedDoc = formatRuns(processedDoc);
            console.log('[PROCESSOR] After run format:', processedDoc.length);

            modifiedFiles['word/document.xml'] = processedDoc;
        } catch (err) {
            console.error('[PROCESSOR] Error:', err);
            throw err;
        }
    }

    // 2. Process styles.xml for font normalization
    const stylesXml = getStylesXml(files);
    if (stylesXml) {
        console.log('[PROCESSOR] styles.xml length:', stylesXml.length);
        try {
            const processedStyles = normalizeFonts(stylesXml);
            console.log('[PROCESSOR] After font normalize:', processedStyles.length);
            modifiedFiles['word/styles.xml'] = processedStyles;
        } catch (err) {
            console.error('[PROCESSOR] Styles error:', err);
        }
    }

    // Repackage with modifications
    console.log('[PROCESSOR] Repackaging with', Object.keys(modifiedFiles).length, 'modified files');
    const outputBuffer = await repackageDocx(zip, modifiedFiles);
    console.log('[PROCESSOR] Output size:', outputBuffer.length);

    return outputBuffer;
}

/**
 * Process multiple DOCX files
 * 
 * @param {Array<{name: string, buffer: Buffer}>} files - Array of files to process
 * @returns {Promise<Array<{name: string, buffer: Buffer}>>} Processed files
 */
export async function processMultipleDocx(files) {
    const results = [];

    for (const file of files) {
        try {
            const processedBuffer = await processDocx(file.buffer);
            results.push({
                name: file.name.replace(/\.docx$/i, '_processed.docx'),
                buffer: processedBuffer,
                success: true
            });
        } catch (error) {
            results.push({
                name: file.name,
                error: error.message,
                success: false
            });
        }
    }

    return results;
}
