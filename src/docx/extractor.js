import JSZip from 'jszip';

/**
 * Extract DOCX file contents
 * @param {Buffer|ArrayBuffer} buffer - DOCX file buffer
 * @returns {Promise<{zip: JSZip, files: Object}>} Extracted ZIP and file map
 */
export async function extractDocx(buffer) {
    const zip = await JSZip.loadAsync(buffer);
    const files = {};

    // Extract all files
    const fileNames = Object.keys(zip.files);
    for (const fileName of fileNames) {
        const file = zip.files[fileName];
        if (!file.dir) {
            // Get content based on file type
            if (fileName.endsWith('.xml') || fileName.endsWith('.rels')) {
                files[fileName] = await file.async('string');
            } else {
                files[fileName] = await file.async('uint8array');
            }
        }
    }

    return { zip, files };
}

/**
 * Repackage modified files into a new DOCX
 * @param {JSZip} originalZip - Original ZIP object
 * @param {Object} modifiedFiles - Map of modified file paths to content
 * @returns {Promise<Buffer>} New DOCX file buffer
 */
export async function repackageDocx(originalZip, modifiedFiles) {
    const newZip = new JSZip();

    // Copy all files from original, replacing modified ones
    const fileNames = Object.keys(originalZip.files);

    for (const fileName of fileNames) {
        const file = originalZip.files[fileName];

        if (file.dir) {
            // Create directory
            newZip.folder(fileName);
        } else if (modifiedFiles[fileName] !== undefined) {
            // Use modified content
            newZip.file(fileName, modifiedFiles[fileName]);
        } else {
            // Copy original content
            const content = await file.async('uint8array');
            newZip.file(fileName, content);
        }
    }

    // Generate new DOCX
    const buffer = await newZip.generateAsync({
        type: 'nodebuffer',
        compression: 'DEFLATE',
        compressionOptions: { level: 6 }
    });

    return buffer;
}

/**
 * Get document.xml content from extracted files
 * @param {Object} files - Extracted files map
 * @returns {string} Document XML content
 */
export function getDocumentXml(files) {
    return files['word/document.xml'] || null;
}

/**
 * Get styles.xml content from extracted files
 * @param {Object} files - Extracted files map
 * @returns {string} Styles XML content
 */
export function getStylesXml(files) {
    return files['word/styles.xml'] || null;
}

/**
 * List all XML files in the DOCX
 * @param {Object} files - Extracted files map
 * @returns {string[]} Array of XML file paths
 */
export function listXmlFiles(files) {
    return Object.keys(files).filter(f => f.endsWith('.xml'));
}
