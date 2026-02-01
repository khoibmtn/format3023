import { XMLParser, XMLBuilder } from 'fast-xml-parser';

/**
 * XML Parser options for OpenXML processing
 * Configured to preserve structure and attributes
 */
const parserOptions = {
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    preserveOrder: true,
    commentPropName: '#comment',
    cdataPropName: '#cdata',
    textNodeName: '#text',
    trimValues: false,
    parseTagValue: false,
    parseAttributeValue: false,
    // Preserve processing instructions like <?xml ...?>
    ignorePiTags: false,
    // Handle standalone attribute
    allowBooleanAttributes: true
};

const builderOptions = {
    ignoreAttributes: false,
    attributeNamePrefix: '@_',
    preserveOrder: true,
    commentPropName: '#comment',
    cdataPropName: '#cdata',
    textNodeName: '#text',
    format: false,
    suppressEmptyNode: false,
    suppressBooleanAttributes: false
};

/**
 * Parse XML string to JavaScript object
 * Stores the XML declaration for later restoration
 * @param {string} xmlString - XML content
 * @returns {Array} Parsed XML structure
 */
export function parseXml(xmlString) {
    const parser = new XMLParser(parserOptions);
    return parser.parse(xmlString);
}

/**
 * Build XML string from JavaScript object
 * Restores XML declaration at the beginning
 * @param {Array} xmlObj - XML object structure
 * @returns {string} XML string
 */
export function buildXml(xmlObj) {
    const builder = new XMLBuilder(builderOptions);
    let xml = builder.build(xmlObj);

    // Ensure XML declaration is present
    if (!xml.startsWith('<?xml')) {
        xml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml;
    }

    return xml;
}

/**
 * Find element by tag name in parsed XML (preserveOrder format)
 * @param {Array|Object} node - XML node
 * @param {string} tagName - Tag name to find
 * @returns {Object|null} Found element or null
 */
export function findElement(node, tagName) {
    if (Array.isArray(node)) {
        for (const item of node) {
            const result = findElement(item, tagName);
            if (result) return result;
        }
        return null;
    }

    if (typeof node === 'object' && node !== null) {
        if (tagName in node) {
            return node;
        }
        for (const key of Object.keys(node)) {
            if (key.startsWith('@_') || key === '#text') continue;
            const result = findElement(node[key], tagName);
            if (result) return result;
        }
    }
    return null;
}

/**
 * Find all elements by tag name in parsed XML
 * @param {Array|Object} node - XML node
 * @param {string} tagName - Tag name to find
 * @returns {Array} Array of found elements
 */
export function findAllElements(node, tagName) {
    const results = [];

    function search(n) {
        if (Array.isArray(n)) {
            for (const item of n) {
                search(item);
            }
        } else if (typeof n === 'object' && n !== null) {
            if (tagName in n) {
                results.push(n);
            }
            for (const key of Object.keys(n)) {
                if (key.startsWith('@_') || key === '#text') continue;
                search(n[key]);
            }
        }
    }

    search(node);
    return results;
}

/**
 * Get or create a child element
 * @param {Object} parent - Parent element (in preserveOrder format)
 * @param {string} parentTag - Parent tag name
 * @param {string} childTag - Child tag to get or create
 * @returns {Array} Child element array
 */
export function getOrCreateChild(parent, parentTag, childTag) {
    if (!parent[parentTag]) {
        parent[parentTag] = [];
    }

    const children = parent[parentTag];
    let child = children.find(c => childTag in c);

    if (!child) {
        child = { [childTag]: [] };
        children.push(child);
    }

    return child;
}

/**
 * Set attribute on an element
 * @param {Object} element - Element object
 * @param {string} tagName - Tag name 
 * @param {string} attrName - Attribute name (without @_)
 * @param {string} value - Attribute value
 */
export function setAttribute(element, tagName, attrName, value) {
    if (!element[tagName]) {
        element[tagName] = [];
    }

    // Find or create :@ object for attributes
    let attrObj = element[tagName].find(c => ':@' in c);
    if (!attrObj) {
        attrObj = { ':@': {} };
        element[tagName].unshift(attrObj);
    }

    attrObj[':@'][`@_${attrName}`] = value;
}

/**
 * Get attribute value from element
 * @param {Object} element - Element object  
 * @param {string} tagName - Tag name
 * @param {string} attrName - Attribute name (without @_)
 * @returns {string|null} Attribute value or null
 */
export function getAttribute(element, tagName, attrName) {
    if (!element[tagName] || !Array.isArray(element[tagName])) {
        return null;
    }

    const attrObj = element[tagName].find(c => ':@' in c);
    if (attrObj && attrObj[':@']) {
        return attrObj[':@'][`@_${attrName}`] || null;
    }
    return null;
}

/**
 * Deep clone an object
 * @param {Object} obj - Object to clone
 * @returns {Object} Cloned object
 */
export function deepClone(obj) {
    return JSON.parse(JSON.stringify(obj));
}

/**
 * Check if element has a specific child
 * @param {Object} element - Parent element
 * @param {string} parentTag - Parent tag name
 * @param {string} childTag - Child tag to check
 * @returns {boolean} True if child exists
 */
export function hasChild(element, parentTag, childTag) {
    if (!element[parentTag] || !Array.isArray(element[parentTag])) {
        return false;
    }
    return element[parentTag].some(c => childTag in c);
}

/**
 * Remove child element by tag name
 * @param {Object} element - Parent element
 * @param {string} parentTag - Parent tag name
 * @param {string} childTag - Child tag to remove
 */
export function removeChild(element, parentTag, childTag) {
    if (!element[parentTag] || !Array.isArray(element[parentTag])) {
        return;
    }
    element[parentTag] = element[parentTag].filter(c => !(childTag in c));
}
