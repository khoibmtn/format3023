import { processDocx } from '../src/docx/processor.js';

// Vercel serverless handler
export const config = {
    api: {
        bodyParser: false,
    },
};

export default async function handler(req, res) {
    // Add CORS headers
    res.setHeader('Access-Control-Allow-Origin', '*');
    res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
    res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

    if (req.method === 'OPTIONS') {
        return res.status(200).end();
    }

    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    try {
        console.log('[API] Starting file processing...');

        // Parse multipart form data
        const { files } = await parseMultipartForm(req);
        console.log('[API] Files received:', files.length);

        if (!files || files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }

        // Process first file only for simplicity
        const file = files[0];
        console.log('[API] Processing:', file.name, 'size:', file.buffer.length);

        const processedBuffer = await processDocx(file.buffer);
        console.log('[API] Processed, output size:', processedBuffer.length);

        const outputName = file.name.replace(/\.docx$/i, '_processed.docx');

        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
        res.setHeader('Content-Disposition', `attachment; filename*=UTF-8''${encodeURIComponent(outputName)}`);

        return res.send(Buffer.from(processedBuffer));

    } catch (error) {
        console.error('[API] Processing error:', error);
        return res.status(500).json({
            error: 'Processing failed',
            details: error.message,
            stack: error.stack
        });
    }
}

/**
 * Parse multipart form data from request
 */
async function parseMultipartForm(req) {
    const busboy = (await import('busboy')).default;

    return new Promise((resolve, reject) => {
        const files = [];
        const bb = busboy({ headers: req.headers });

        bb.on('file', (name, file, info) => {
            const { filename } = info;
            console.log('[API] Receiving file:', filename);

            if (!filename.toLowerCase().endsWith('.docx')) {
                console.log('[API] Skipping non-docx file:', filename);
                file.resume();
                return;
            }

            const chunks = [];
            file.on('data', chunk => chunks.push(chunk));
            file.on('end', () => {
                console.log('[API] File received, chunks:', chunks.length);
                files.push({
                    name: filename,
                    buffer: Buffer.concat(chunks)
                });
            });
        });

        bb.on('close', () => {
            console.log('[API] Form parsing complete, files:', files.length);
            resolve({ files });
        });
        bb.on('error', (err) => {
            console.error('[API] Busboy error:', err);
            reject(err);
        });

        req.pipe(bb);
    });
}
