import { parseXml, buildXml } from './xml-utils.js';
import { DEFAULT_FONT, FONT_SIZE_13PT } from '../utils/constants.js';

/**
 * Apply font normalization to document defaults
 * Sets Times New Roman 13pt as default without affecting run-level overrides
 * 
 * @param {string} stylesXml - The word/styles.xml content
 * @returns {string} Modified styles XML
 */
export function normalizeFonts(stylesXml) {
    // First, remove all contextualSpacing elements using regex (before parsing)
    // This ensures "Don't add a space between paragraphs of the same style" is disabled
    let cleanedStylesXml = stylesXml;
    cleanedStylesXml = cleanedStylesXml.replace(/<w:contextualSpacing[^>]*\/?>/g, '');
    cleanedStylesXml = cleanedStylesXml.replace(/<w:contextualSpacing[^>]*>[\s\S]*?<\/w:contextualSpacing>/g, '');

    const parsed = parseXml(cleanedStylesXml);

    // Find or create w:docDefaults
    processDocDefaults(parsed);

    return buildXml(parsed);
}

/**
 * Process document defaults to set font
 */
function processDocDefaults(parsed) {
    // Find w:styles root
    const stylesRoot = findStylesRoot(parsed);
    if (!stylesRoot) return;

    const stylesContent = stylesRoot['w:styles'];
    if (!Array.isArray(stylesContent)) return;

    // Find or create docDefaults
    let docDefaultsIndex = stylesContent.findIndex(item => 'w:docDefaults' in item);
    let docDefaults;

    if (docDefaultsIndex === -1) {
        docDefaults = { 'w:docDefaults': [] };
        // Insert after any namespace declarations
        stylesContent.splice(1, 0, docDefaults);
    } else {
        docDefaults = stylesContent[docDefaultsIndex];
    }

    const docDefaultsContent = docDefaults['w:docDefaults'];

    // Find or create rPrDefault
    let rPrDefaultIndex = docDefaultsContent.findIndex(item => 'w:rPrDefault' in item);
    let rPrDefault;

    if (rPrDefaultIndex === -1) {
        rPrDefault = { 'w:rPrDefault': [] };
        docDefaultsContent.push(rPrDefault);
    } else {
        rPrDefault = docDefaultsContent[rPrDefaultIndex];
    }

    const rPrDefaultContent = rPrDefault['w:rPrDefault'];

    // Find or create rPr
    let rPrIndex = rPrDefaultContent.findIndex(item => 'w:rPr' in item);
    let rPr;

    if (rPrIndex === -1) {
        rPr = { 'w:rPr': [] };
        rPrDefaultContent.push(rPr);
    } else {
        rPr = rPrDefaultContent[rPrIndex];
    }

    const rPrContent = rPr['w:rPr'];

    // Set font
    applyDefaultFont(rPrContent);

    // Set font size
    applyDefaultFontSize(rPrContent);
}

/**
 * Find the w:styles root element
 */
function findStylesRoot(parsed) {
    if (Array.isArray(parsed)) {
        for (const item of parsed) {
            if (item && typeof item === 'object' && 'w:styles' in item) {
                return item;
            }
            const result = findStylesRoot(item);
            if (result) return result;
        }
    } else if (typeof parsed === 'object' && parsed !== null) {
        if ('w:styles' in parsed) return parsed;
        for (const key of Object.keys(parsed)) {
            if (key.startsWith('@_') || key === '#text' || key === ':@') continue;
            const result = findStylesRoot(parsed[key]);
            if (result) return result;
        }
    }
    return null;
}

/**
 * Apply default font to rPr
 */
function applyDefaultFont(rPrContent) {
    // Find existing rFonts
    let rFontsIndex = rPrContent.findIndex(item => 'w:rFonts' in item);

    if (rFontsIndex === -1) {
        // Create new rFonts
        rPrContent.push({
            'w:rFonts': [{
                ':@': {
                    '@_w:ascii': DEFAULT_FONT,
                    '@_w:hAnsi': DEFAULT_FONT,
                    '@_w:cs': DEFAULT_FONT,
                    '@_w:eastAsia': DEFAULT_FONT
                }
            }]
        });
    } else {
        // Update existing rFonts
        const rFonts = rPrContent[rFontsIndex]['w:rFonts'];
        let attrObj = rFonts.find(c => ':@' in c);

        if (!attrObj) {
            attrObj = { ':@': {} };
            rFonts.unshift(attrObj);
        }

        attrObj[':@']['@_w:ascii'] = DEFAULT_FONT;
        attrObj[':@']['@_w:hAnsi'] = DEFAULT_FONT;
        attrObj[':@']['@_w:cs'] = DEFAULT_FONT;
        attrObj[':@']['@_w:eastAsia'] = DEFAULT_FONT;
    }
}

/**
 * Apply default font size to rPr
 */
function applyDefaultFontSize(rPrContent) {
    // Set sz (font size)
    let szIndex = rPrContent.findIndex(item => 'w:sz' in item);

    if (szIndex === -1) {
        rPrContent.push({
            'w:sz': [{
                ':@': { '@_w:val': String(FONT_SIZE_13PT) }
            }]
        });
    } else {
        const sz = rPrContent[szIndex]['w:sz'];
        let attrObj = sz.find(c => ':@' in c);
        if (!attrObj) {
            attrObj = { ':@': {} };
            sz.unshift(attrObj);
        }
        attrObj[':@']['@_w:val'] = String(FONT_SIZE_13PT);
    }

    // Set szCs (complex script font size)
    let szCsIndex = rPrContent.findIndex(item => 'w:szCs' in item);

    if (szCsIndex === -1) {
        rPrContent.push({
            'w:szCs': [{
                ':@': { '@_w:val': String(FONT_SIZE_13PT) }
            }]
        });
    } else {
        const szCs = rPrContent[szCsIndex]['w:szCs'];
        let attrObj = szCs.find(c => ':@' in c);
        if (!attrObj) {
            attrObj = { ':@': {} };
            szCs.unshift(attrObj);
        }
        attrObj[':@']['@_w:val'] = String(FONT_SIZE_13PT);
    }
}
