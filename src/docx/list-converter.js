/**
 * Convert automatic bullets and numbering to plain text
 * 
 * Handles:
 * - Bullet lists: converts to "- " or "+ " prefix
 * - Numbered lists: converts to "1. ", "2. ", etc.
 * 
 * @param {string} documentXml - The word/document.xml content
 * @param {string} numberingXml - The word/numbering.xml content (optional)
 * @returns {string} Modified document XML
 */
export function convertListsToText(documentXml, numberingXml) {
    console.log('[LIST-CONVERTER] Starting list conversion...');

    // Parse numbering definitions if available
    const bulletTypes = parseNumberingDefinitions(numberingXml);
    console.log('[LIST-CONVERTER] Found bullet types:', Object.keys(bulletTypes).length);

    // Track counters for each numId and ilvl
    const counters = {};

    let result = documentXml;

    // Find all paragraphs with w:numPr and process them
    result = processNumberedParagraphs(result, bulletTypes, counters);

    console.log('[LIST-CONVERTER] Conversion complete');
    return result;
}

/**
 * Parse numbering.xml to determine bullet/number types
 * @param {string} numberingXml - The word/numbering.xml content
 * @returns {Object} Map of numId -> type info
 */
function parseNumberingDefinitions(numberingXml) {
    const definitions = {};

    if (!numberingXml) {
        console.log('[LIST-CONVERTER] No numbering.xml found');
        return definitions;
    }

    // Extract abstract numbering definitions
    // w:abstractNum contains the actual format definitions
    // w:num references w:abstractNum via w:abstractNumId

    // First, parse abstract numbering definitions
    const abstractDefs = {};
    const abstractMatches = numberingXml.matchAll(/<w:abstractNum[^>]*w:abstractNumId="(\d+)"[^>]*>([\s\S]*?)<\/w:abstractNum>/g);

    for (const match of abstractMatches) {
        const abstractNumId = match[1];
        const content = match[2];

        // Parse levels
        const levels = {};
        const levelMatches = content.matchAll(/<w:lvl[^>]*w:ilvl="(\d+)"[^>]*>([\s\S]*?)<\/w:lvl>/g);

        for (const lvlMatch of levelMatches) {
            const ilvl = lvlMatch[1];
            const lvlContent = lvlMatch[2];

            // Determine if bullet or number
            const numFmtMatch = lvlContent.match(/<w:numFmt[^>]*w:val="([^"]+)"/);
            const lvlTextMatch = lvlContent.match(/<w:lvlText[^>]*w:val="([^"]*)"/);

            const numFmt = numFmtMatch ? numFmtMatch[1] : 'decimal';
            const lvlText = lvlTextMatch ? lvlTextMatch[1] : '';

            levels[ilvl] = {
                numFmt: numFmt,
                lvlText: lvlText,
                isBullet: numFmt === 'bullet',
                isNumber: numFmt === 'decimal' || numFmt === 'lowerLetter' || numFmt === 'upperLetter' ||
                    numFmt === 'lowerRoman' || numFmt === 'upperRoman'
            };
        }

        abstractDefs[abstractNumId] = levels;
    }

    // Now parse w:num to map numId to abstractNumId
    const numMatches = numberingXml.matchAll(/<w:num[^>]*w:numId="(\d+)"[^>]*>([\s\S]*?)<\/w:num>/g);

    for (const match of numMatches) {
        const numId = match[1];
        const content = match[2];

        const abstractRefMatch = content.match(/<w:abstractNumId[^>]*w:val="(\d+)"/);
        if (abstractRefMatch) {
            const abstractNumId = abstractRefMatch[1];
            if (abstractDefs[abstractNumId]) {
                definitions[numId] = abstractDefs[abstractNumId];
            }
        }
    }

    return definitions;
}

/**
 * Process all paragraphs with numbering
 */
function processNumberedParagraphs(xml, bulletTypes, counters) {
    let result = xml;

    // Find paragraphs with w:numPr
    // We need to process them in order, tracking position

    const paragraphRegex = /<w:p[^>]*>([\s\S]*?)<\/w:p>/g;
    let offset = 0;

    // Collect all paragraphs first to process in order
    const paragraphs = [];
    let match;
    while ((match = paragraphRegex.exec(xml)) !== null) {
        paragraphs.push({
            fullMatch: match[0],
            content: match[1],
            index: match.index
        });
    }

    // Process each paragraph
    for (const para of paragraphs) {
        // Check if paragraph has w:numPr
        const numPrMatch = para.content.match(/<w:numPr>([\s\S]*?)<\/w:numPr>/);
        if (!numPrMatch) continue;

        const numPrContent = numPrMatch[1];

        // Extract numId and ilvl
        const ilvlMatch = numPrContent.match(/<w:ilvl[^>]*w:val="(\d+)"/);
        const numIdMatch = numPrContent.match(/<w:numId[^>]*w:val="(\d+)"/);

        if (!numIdMatch) continue;

        const numId = numIdMatch[1];
        const ilvl = ilvlMatch ? ilvlMatch[1] : '0';

        // Get bullet/number type
        const levelDef = bulletTypes[numId]?.[ilvl];

        // Determine prefix to add
        let prefix = '';

        if (levelDef) {
            if (levelDef.isBullet) {
                // Check bullet character
                const bulletChar = levelDef.lvlText;
                if (bulletChar === '-' || bulletChar === '–' || bulletChar === '—' || bulletChar === '') {
                    prefix = '- ';
                } else {
                    prefix = '+ ';
                }
            } else if (levelDef.isNumber) {
                // Get counter for this numId/ilvl
                const counterKey = `${numId}-${ilvl}`;
                if (!counters[counterKey]) {
                    counters[counterKey] = 0;
                }
                counters[counterKey]++;

                // Format number based on numFmt
                const num = counters[counterKey];
                prefix = formatNumber(num, levelDef.numFmt, levelDef.lvlText);
            }
        } else {
            // Default: assume bullet dash
            prefix = '- ';
        }

        // Create modified paragraph
        let modifiedContent = para.content;

        // Remove w:numPr
        modifiedContent = modifiedContent.replace(/<w:numPr>[\s\S]*?<\/w:numPr>/g, '');

        // Add prefix to the first w:t element
        if (prefix) {
            modifiedContent = addPrefixToText(modifiedContent, prefix);
        }

        const modifiedPara = para.fullMatch.replace(para.content, modifiedContent);

        // Replace in result
        const adjustedIndex = para.index + offset;
        result = result.substring(0, adjustedIndex) + modifiedPara + result.substring(adjustedIndex + para.fullMatch.length);

        // Adjust offset for next iteration
        offset += modifiedPara.length - para.fullMatch.length;
    }

    return result;
}

/**
 * Format number based on numFmt
 */
function formatNumber(num, numFmt, lvlText) {
    let formatted;

    switch (numFmt) {
        case 'lowerLetter':
            formatted = String.fromCharCode(96 + ((num - 1) % 26) + 1);
            break;
        case 'upperLetter':
            formatted = String.fromCharCode(64 + ((num - 1) % 26) + 1);
            break;
        case 'lowerRoman':
            formatted = toRoman(num).toLowerCase();
            break;
        case 'upperRoman':
            formatted = toRoman(num);
            break;
        default:
            formatted = String(num);
    }

    // Apply lvlText template if present (e.g., "%1." becomes "1.")
    if (lvlText) {
        return lvlText.replace(/%\d+/, formatted) + ' ';
    }

    return formatted + '. ';
}

/**
 * Convert number to Roman numerals
 */
function toRoman(num) {
    const romanNumerals = [
        ['M', 1000], ['CM', 900], ['D', 500], ['CD', 400],
        ['C', 100], ['XC', 90], ['L', 50], ['XL', 40],
        ['X', 10], ['IX', 9], ['V', 5], ['IV', 4], ['I', 1]
    ];

    let result = '';
    for (const [letter, value] of romanNumerals) {
        while (num >= value) {
            result += letter;
            num -= value;
        }
    }
    return result;
}

/**
 * Add prefix text to the first w:t element in a paragraph
 */
function addPrefixToText(paragraphContent, prefix) {
    // Find the first w:t element
    const firstTMatch = paragraphContent.match(/<w:t(\s[^>]*)?>([^<]*)<\/w:t>/);

    if (firstTMatch) {
        const originalT = firstTMatch[0];
        const attrs = firstTMatch[1] || '';
        const text = firstTMatch[2];

        // Ensure xml:space="preserve" to keep spaces
        let newAttrs = attrs;
        if (!attrs.includes('xml:space')) {
            newAttrs = ' xml:space="preserve"' + attrs;
        }

        const newT = `<w:t${newAttrs}>${prefix}${text}</w:t>`;

        return paragraphContent.replace(originalT, newT);
    }

    // If no w:t found, need to create one in a w:r element
    // Find the first w:r element
    const firstRMatch = paragraphContent.match(/<w:r>([^<]*<\/w:r>|[\s\S]*?<\/w:r>)/);

    if (firstRMatch) {
        const originalR = firstRMatch[0];
        // Add w:t at the beginning of the run content
        const newR = originalR.replace(/<w:r>/, `<w:r><w:t xml:space="preserve">${prefix}</w:t>`);
        return paragraphContent.replace(originalR, newR);
    }

    return paragraphContent;
}
