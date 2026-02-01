import { parseXml, buildXml, findAllElements, deepClone, getAttribute } from './xml-utils.js';
import { EN_DASH, HYPHEN } from '../utils/constants.js';

/**
 * Replace all manual line breaks with paragraph breaks
 * and replace en dashes with hyphens
 * 
 * @param {string} documentXml - The word/document.xml content
 * @returns {string} Modified document XML
 */
export function replaceText(documentXml) {
    const parsed = parseXml(documentXml);

    // Process the document body
    processNode(parsed);

    return buildXml(parsed);
}

/**
 * Recursively process nodes to find paragraphs and perform replacements
 */
function processNode(node) {
    if (Array.isArray(node)) {
        // Process arrays in place, handling potential splits
        let i = 0;
        while (i < node.length) {
            const item = node[i];

            // Check if this is a paragraph that needs splitting
            if (item && typeof item === 'object' && 'w:p' in item) {
                const newParagraphs = processParagraph(item);
                if (newParagraphs.length > 1) {
                    // Replace single paragraph with multiple
                    node.splice(i, 1, ...newParagraphs);
                    i += newParagraphs.length;
                    continue;
                }
            } else {
                processNode(item);
            }
            i++;
        }
    } else if (typeof node === 'object' && node !== null) {
        for (const key of Object.keys(node)) {
            if (key.startsWith('@_') || key === '#text' || key === ':@') continue;
            processNode(node[key]);
        }
    }
}

/**
 * Process a single paragraph: replace en dashes and split on line breaks
 * @param {Object} paragraph - Paragraph object with 'w:p' key
 * @returns {Array} Array of paragraph objects (may be more than one if split)
 */
function processParagraph(paragraph) {
    const pContent = paragraph['w:p'];
    if (!Array.isArray(pContent)) return [paragraph];

    // First pass: replace en dashes in all text nodes
    replaceEnDashesInParagraph(pContent);

    // Second pass: find line breaks and split
    const breakIndices = findLineBreakIndices(pContent);

    if (breakIndices.length === 0) {
        return [paragraph];
    }

    // Split the paragraph at each break
    return splitParagraphAtBreaks(paragraph, pContent, breakIndices);
}

/**
 * Replace en dashes with hyphens in all text nodes of a paragraph
 */
function replaceEnDashesInParagraph(pContent) {
    for (const item of pContent) {
        if ('w:r' in item) {
            replaceEnDashesInRun(item['w:r']);
        }
    }
}

/**
 * Replace en dashes in a run's text nodes
 */
function replaceEnDashesInRun(runContent) {
    if (!Array.isArray(runContent)) return;

    for (const item of runContent) {
        if ('w:t' in item && Array.isArray(item['w:t'])) {
            for (const textItem of item['w:t']) {
                if ('#text' in textItem && typeof textItem['#text'] === 'string') {
                    textItem['#text'] = textItem['#text'].replace(new RegExp(EN_DASH, 'g'), HYPHEN);
                }
            }
        }
    }
}

/**
 * Find indices of runs containing line breaks
 * @returns {Array<{runIndex: number, breakIndex: number}>}
 */
function findLineBreakIndices(pContent) {
    const breaks = [];

    for (let runIdx = 0; runIdx < pContent.length; runIdx++) {
        const item = pContent[runIdx];
        if (!('w:r' in item)) continue;

        const runContent = item['w:r'];
        if (!Array.isArray(runContent)) continue;

        for (let brIdx = 0; brIdx < runContent.length; brIdx++) {
            const runItem = runContent[brIdx];
            if ('w:br' in runItem) {
                // Check if it's a text wrapping break (not page break)
                const breakType = getBreakType(runItem);
                if (breakType !== 'page' && breakType !== 'column') {
                    breaks.push({ runIndex: runIdx, breakIndex: brIdx });
                }
            }
        }
    }

    return breaks;
}

/**
 * Get the type of break element
 */
function getBreakType(breakElement) {
    const brContent = breakElement['w:br'];
    if (!Array.isArray(brContent)) return 'textWrapping';

    const attrObj = brContent.find(c => ':@' in c);
    if (attrObj && attrObj[':@'] && attrObj[':@']['@_w:type']) {
        return attrObj[':@']['@_w:type'];
    }
    return 'textWrapping';
}

/**
 * Split paragraph at break points
 */
function splitParagraphAtBreaks(originalParagraph, pContent, breakIndices) {
    const paragraphs = [];

    // Get paragraph properties to copy
    const pPr = pContent.find(item => 'w:pPr' in item);

    let currentParagraphContent = [];
    let contentStartIdx = 0;

    for (const { runIndex, breakIndex } of breakIndices) {
        // Copy content up to the break
        for (let i = contentStartIdx; i < runIndex; i++) {
            currentParagraphContent.push(deepClone(pContent[i]));
        }

        // Handle the run containing the break - split it
        const breakRun = pContent[runIndex];
        const runContent = breakRun['w:r'];

        if (breakIndex > 0) {
            // There's content before the break in this run
            const beforeBreakRun = createPartialRun(breakRun, runContent.slice(0, breakIndex));
            if (beforeBreakRun) {
                currentParagraphContent.push(beforeBreakRun);
            }
        }

        // Create paragraph from accumulated content
        paragraphs.push(createParagraph(pPr, currentParagraphContent));

        // Reset for next paragraph
        currentParagraphContent = [];

        // Content after the break in this run
        if (breakIndex + 1 < runContent.length) {
            const afterBreakRun = createPartialRun(breakRun, runContent.slice(breakIndex + 1));
            if (afterBreakRun) {
                currentParagraphContent.push(afterBreakRun);
            }
        }

        contentStartIdx = runIndex + 1;
    }

    // Add remaining content to final paragraph
    for (let i = contentStartIdx; i < pContent.length; i++) {
        if (!('w:pPr' in pContent[i])) {
            currentParagraphContent.push(deepClone(pContent[i]));
        }
    }

    if (currentParagraphContent.length > 0 || paragraphs.length === 0) {
        paragraphs.push(createParagraph(pPr, currentParagraphContent));
    }

    return paragraphs;
}

/**
 * Create a partial run with specified content
 */
function createPartialRun(originalRun, contentSlice) {
    if (!contentSlice || contentSlice.length === 0) return null;

    const runContent = originalRun['w:r'];
    const rPr = runContent.find(item => 'w:rPr' in item);

    const newRunContent = [];
    if (rPr) {
        newRunContent.push(deepClone(rPr));
    }
    newRunContent.push(...contentSlice.map(c => deepClone(c)));

    return { 'w:r': newRunContent };
}

/**
 * Create a new paragraph with given properties and content
 */
function createParagraph(pPr, content) {
    const newPContent = [];

    // Add paragraph properties first
    if (pPr) {
        newPContent.push(deepClone(pPr));
    }

    // Add content
    newPContent.push(...content);

    return { 'w:p': newPContent };
}
