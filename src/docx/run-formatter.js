import { DEFAULT_FONT, FONT_SIZE_14PT } from '../utils/constants.js';

/**
 * Apply font and size formatting to all runs in document
 * Sets Times New Roman 14pt for all text
 * 
 * @param {string} documentXml - The word/document.xml content
 * @returns {string} Modified document XML
 */
export function formatRuns(documentXml) {
    let result = documentXml;

    // Process each w:r element to set font and size
    result = processFontAndSize(result);

    return result;
}

/**
 * Process w:r elements to add/update font and size
 */
function processFontAndSize(xml) {
    let result = xml;

    // First, update existing w:sz elements to 14pt
    result = result.replace(/<w:sz w:val="[^"]*"/g, `<w:sz w:val="${FONT_SIZE_14PT}"`);
    result = result.replace(/<w:szCs w:val="[^"]*"/g, `<w:szCs w:val="${FONT_SIZE_14PT}"`);

    // Update existing w:rFonts to Times New Roman
    result = result.replace(
        /<w:rFonts([^>]*)w:ascii="[^"]*"/g,
        `<w:rFonts$1w:ascii="${DEFAULT_FONT}"`
    );
    result = result.replace(
        /<w:rFonts([^>]*)w:hAnsi="[^"]*"/g,
        `<w:rFonts$1w:hAnsi="${DEFAULT_FONT}"`
    );
    result = result.replace(
        /<w:rFonts([^>]*)w:cs="[^"]*"/g,
        `<w:rFonts$1w:cs="${DEFAULT_FONT}"`
    );

    // For runs without w:rPr, add one with font settings
    // This is complex, so we'll focus on runs that already have w:rPr

    // For w:rPr that don't have w:sz, add it
    // Find </w:rPr> that don't have w:sz before them
    result = addMissingSz(result);

    return result;
}

/**
 * Add w:sz and w:szCs to w:rPr elements that don't have them
 */
function addMissingSz(xml) {
    // This is a simplified approach - add sz before </w:rPr> if not present
    // We need to be careful not to add duplicates

    let result = xml;

    // Pattern: Find w:rPr blocks and check if they have w:sz
    result = result.replace(/<w:rPr>([\s\S]*?)<\/w:rPr>/g, (match, content) => {
        let newContent = content;

        // Add w:sz if not present
        if (!content.includes('<w:sz')) {
            newContent += `<w:sz w:val="${FONT_SIZE_14PT}"/>`;
        }

        // Add w:szCs if not present
        if (!content.includes('<w:szCs')) {
            newContent += `<w:szCs w:val="${FONT_SIZE_14PT}"/>`;
        }

        // Add w:rFonts if not present
        if (!content.includes('<w:rFonts')) {
            newContent = `<w:rFonts w:ascii="${DEFAULT_FONT}" w:hAnsi="${DEFAULT_FONT}" w:cs="${DEFAULT_FONT}"/>` + newContent;
        }

        return `<w:rPr>${newContent}</w:rPr>`;
    });

    return result;
}
