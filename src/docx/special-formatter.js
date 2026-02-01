/**
 * Special paragraph formatting for document structure:
 * - Insert blank lines before key sections
 * - Title formatting handled separately
 * 
 * @param {string} documentXml - The word/document.xml content
 * @returns {string} Modified document XML
 */
export function formatSpecialParagraphs(documentXml) {
    let result = documentXml;

    // 1. Insert blank line AFTER the first paragraph (title)
    result = insertBlankLineAfterTitleParagraph(result);

    // 2. Insert blank line before "TÀI LIỆU THAM KHẢO" if not already present
    result = insertBlankLineBeforeText(result, 'TÀI LIỆU THAM KHẢO');

    return result;
}

/**
 * Format the title paragraph (MUST run after formatParagraphs)
 * Title should be: center aligned, left=0, firstLine=0, right=0
 * 
 * @param {string} documentXml - The word/document.xml content
 * @returns {string} Modified document XML
 */
export function formatTitleParagraph(documentXml) {
    const bodyStart = documentXml.indexOf('<w:body');
    if (bodyStart === -1) return documentXml;

    const afterBody = documentXml.substring(bodyStart);
    const firstPStart = afterBody.indexOf('<w:p');
    if (firstPStart === -1) return documentXml;

    const titleStart = bodyStart + firstPStart;

    // Find end of first paragraph
    let endPos = titleStart;
    let depth = 0;
    while (endPos < documentXml.length) {
        if (documentXml.substring(endPos, endPos + 4) === '<w:p') {
            const nc = documentXml[endPos + 4];
            if (nc === '>' || nc === ' ') depth++;
        } else if (documentXml.substring(endPos, endPos + 6) === '</w:p>') {
            depth--;
            if (depth === 0) {
                endPos += 6;
                break;
            }
        }
        endPos++;
    }

    let titleParagraph = documentXml.substring(titleStart, endPos);

    // Replace w:jc with center
    titleParagraph = titleParagraph.replace(/<w:jc w:val="[^"]*"\/>/g, '<w:jc w:val="center"/>');

    // Replace w:ind with all zeros (explicitly set left, firstLine, right to 0)
    titleParagraph = titleParagraph.replace(
        /<w:ind[^>]*\/>/g,
        '<w:ind w:left="0" w:firstLine="0" w:right="0"/>'
    );

    return documentXml.substring(0, titleStart) + titleParagraph + documentXml.substring(endPos);
}

/**
 * Insert blank line after the first paragraph (title) in the document body
 */
function insertBlankLineAfterTitleParagraph(xml) {
    const bodyStart = xml.indexOf('<w:body');
    if (bodyStart === -1) return xml;

    const afterBody = xml.substring(bodyStart);
    const firstPEnd = afterBody.indexOf('</w:p>');
    if (firstPEnd === -1) return xml;

    const insertPosition = bodyStart + firstPEnd + 6;

    // Check if there's already a blank/empty paragraph after title
    const nextContent = xml.substring(insertPosition, insertPosition + 500);
    const nextParagraphMatch = nextContent.match(/^\s*<w:p[^>]*>([\s\S]*?)<\/w:p>/);

    if (nextParagraphMatch) {
        const pContent = nextParagraphMatch[1];
        if (!pContent.includes('<w:t')) {
            return xml; // Already has blank
        }
    }

    const blankParagraph = '<w:p><w:pPr></w:pPr></w:p>';
    return xml.substring(0, insertPosition) + blankParagraph + xml.substring(insertPosition);
}

/**
 * Insert blank line before a paragraph containing specific text
 */
function insertBlankLineBeforeText(xml, searchText) {
    const textIndex = xml.indexOf(searchText);
    if (textIndex === -1) return xml;

    let searchStart = textIndex;
    let pStartIndex = -1;

    while (searchStart > 0) {
        const checkStr = xml.substring(searchStart, searchStart + 4);
        if (checkStr === '<w:p') {
            const nextChar = xml[searchStart + 4];
            if (nextChar === '>' || nextChar === ' ') {
                pStartIndex = searchStart;
                break;
            }
        }
        searchStart--;
    }

    if (pStartIndex === -1) return xml;

    const beforeTarget = xml.substring(0, pStartIndex);
    const lastPEndIndex = beforeTarget.lastIndexOf('</w:p>');

    if (lastPEndIndex === -1) return xml;

    let prevPStart = lastPEndIndex;
    while (prevPStart > 0) {
        if (xml.substring(prevPStart, prevPStart + 4) === '<w:p') {
            const nextChar = xml[prevPStart + 4];
            if (nextChar === '>' || nextChar === ' ') {
                break;
            }
        }
        prevPStart--;
    }

    const prevParagraph = xml.substring(prevPStart, lastPEndIndex + 6);
    const hasText = prevParagraph.match(/<w:t[^>]*>([^<]+)<\/w:t>/);

    if (!hasText || hasText[1].trim() === '') {
        return xml; // Previous paragraph is blank
    }

    const blankParagraph = '<w:p><w:pPr></w:pPr></w:p>';
    return xml.substring(0, pStartIndex) + blankParagraph + xml.substring(pStartIndex);
}
