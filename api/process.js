import { processDocx, processMultipleDocx } from '../src/docx/processor.js';
import archiver from 'archiver';

// Vercel serverless handler
export const config = {
    api: {
        bodyParser: false,
    },
};

export default async function handler(req, res) {
    if (req.method !== 'POST') {
        return res.status(405).json({ error: 'Method not allowed' });
    }

    try {
        // Parse multipart form data
        const { files } = await parseMultipartForm(req);

        if (!files || files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }

        const results = await processMultipleDocx(files);
        const successfulResults = results.filter(r => r.success);
        const failedResults = results.filter(r => !r.success);

        if (successfulResults.length === 0) {
            return res.status(500).json({
                error: 'All files failed to process',
                failures: failedResults
            });
        }

        // If single file, return directly
        if (successfulResults.length === 1 && failedResults.length === 0) {
            const result = successfulResults[0];
            res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
            res.setHeader('Content-Disposition', `attachment; filename="${encodeURIComponent(result.name)}"`);
            return res.send(result.buffer);
        }

        // Multiple files: return as ZIP
        res.setHeader('Content-Type', 'application/zip');
        res.setHeader('Content-Disposition', 'attachment; filename="processed_documents.zip"');

        const archive = archiver('zip', { zlib: { level: 6 } });
        archive.pipe(res);

        for (const result of successfulResults) {
            archive.append(result.buffer, { name: result.name });
        }

        if (failedResults.length > 0) {
            const errorReport = failedResults.map(f => `${f.name}: ${f.error}`).join('\n');
            archive.append(errorReport, { name: 'errors.txt' });
        }

        await archive.finalize();

    } catch (error) {
        console.error('Processing error:', error);
        res.status(500).json({ error: error.message });
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
            const { filename, mimeType } = info;

            if (!filename.toLowerCase().endsWith('.docx')) {
                file.resume();
                return;
            }

            const chunks = [];
            file.on('data', chunk => chunks.push(chunk));
            file.on('end', () => {
                files.push({
                    name: filename,
                    buffer: Buffer.concat(chunks)
                });
            });
        });

        bb.on('close', () => resolve({ files }));
        bb.on('error', reject);

        req.pipe(bb);
    });
}
