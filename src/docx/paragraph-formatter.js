import {
    LINE_SPACING_1_1,
    SPACE_BEFORE_6PT,
    SPACE_AFTER_0PT,
    RIGHT_INDENT_0_25CM,
    FIRST_LINE_INDENT_1CM
} from '../utils/constants.js';

/**
 * Apply paragraph-level formatting to document using regex
 * - Line spacing: Multiple 1.1
 * - Space before: 6pt, Space after: 0pt
 * - Right indent: 0.25cm
 * - First line indent: 1cm
 * - Left alignment (not justify)
 * 
 * @param {string} documentXml - The word/document.xml content
 * @returns {string} Modified document XML
 */
export function formatParagraphs(documentXml) {
    let result = documentXml;

    // 1. Replace or add w:spacing with our values in all w:pPr elements
    result = processSpacing(result);

    // 2. Add/update w:ind (indentation) in all w:pPr elements
    result = processIndentation(result);

    // 3. Add/update left alignment
    result = processAlignment(result);

    return result;
}

/**
 * Process w:spacing elements - replace existing or add new
 */
function processSpacing(xml) {
    const spacingAttrs = `w:line="${LINE_SPACING_1_1}" w:lineRule="auto" w:before="${SPACE_BEFORE_6PT}" w:after="${SPACE_AFTER_0PT}"`;

    let result = xml;

    // Remove existing w:spacing elements (both empty and with content)
    result = result.replace(/<w:spacing[^>]*\/?>/g, '');
    result = result.replace(/<w:spacing[^>]*>[\s\S]*?<\/w:spacing>/g, '');
    result = result.replace(/<\/w:spacing>/g, '');  // orphan closing tags

    // Add new w:spacing at the end of each w:pPr (before </w:pPr>)
    result = result.replace(/<\/w:pPr>/g, `<w:spacing ${spacingAttrs}/></w:pPr>`);

    return result;
}

/**
 * Process w:ind elements - add first line indent (1cm), hanging=0, right indent
 */
function processIndentation(xml) {
    let result = xml;

    // Remove existing w:ind elements
    result = result.replace(/<w:ind[^>]*\/?>/g, '');
    result = result.replace(/<w:ind[^>]*>[\s\S]*?<\/w:ind>/g, '');
    result = result.replace(/<\/w:ind>/g, '');  // orphan closing tags

    // Add new w:ind with left=0, firstLine=1cm, right=0.25cm before w:spacing
    // w:left="0" explicitly resets any inherited left indent
    const indentAttrs = `w:left="0" w:firstLine="${FIRST_LINE_INDENT_1CM}" w:right="${RIGHT_INDENT_0_25CM}"`;

    result = result.replace(/<w:spacing/g, `<w:ind ${indentAttrs}/><w:spacing`);

    return result;
}

/**
 * Process justification - set to LEFT alignment
 */
function processAlignment(xml) {
    let result = xml;

    // Remove existing w:jc elements
    result = result.replace(/<w:jc[^>]*\/?>/g, '');
    result = result.replace(/<w:jc[^>]*>[\s\S]*?<\/w:jc>/g, '');
    result = result.replace(/<\/w:jc>/g, '');  // orphan closing tags

    // Add LEFT alignment before w:ind
    // w:val="left" for left alignment
    result = result.replace(/<w:ind/g, `<w:jc w:val="left"/><w:ind`);

    return result;
}
