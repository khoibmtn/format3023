import express from 'express';
import multer from 'multer';
import archiver from 'archiver';
import path from 'path';
import { fileURLToPath } from 'url';
import { processDocx, processMultipleDocx } from './docx/processor.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = process.env.PORT || 3000;

// Configure multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({
    storage,
    limits: { fileSize: 50 * 1024 * 1024 }, // 50MB limit
    fileFilter: (req, file, cb) => {
        if (file.mimetype === 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
            file.originalname.toLowerCase().endsWith('.docx')) {
            cb(null, true);
        } else {
            cb(new Error('Only .docx files are allowed'), false);
        }
    }
});

// Serve static files
app.use(express.static(path.join(__dirname, '../public')));

// Health check
app.get('/api/health', (req, res) => {
    res.json({ status: 'ok', timestamp: new Date().toISOString() });
});

// Process single or multiple DOCX files
app.post('/api/process', upload.array('files', 20), async (req, res) => {
    try {
        if (!req.files || req.files.length === 0) {
            return res.status(400).json({ error: 'No files uploaded' });
        }

        const files = req.files.map(f => ({
            name: f.originalname,
            buffer: f.buffer
        }));

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
            // RFC 5987 encoding for UTF-8 filenames
            const encodedFilename = encodeURIComponent(result.name).replace(/'/g, '%27');
            const asciiFilename = result.name.replace(/[^\x20-\x7E]/g, '_');

            res.set({
                'Content-Type': 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
                'Content-Disposition': `attachment; filename="${asciiFilename}"; filename*=UTF-8''${encodedFilename}`,
                'Content-Length': result.buffer.length
            });
            return res.send(result.buffer);
        }

        // Multiple files: return as ZIP
        res.set({
            'Content-Type': 'application/zip',
            'Content-Disposition': 'attachment; filename="processed_documents.zip"'
        });

        const archive = archiver('zip', { zlib: { level: 6 } });
        archive.pipe(res);

        for (const result of successfulResults) {
            archive.append(result.buffer, { name: result.name });
        }

        // Add error report if any failures
        if (failedResults.length > 0) {
            const errorReport = failedResults.map(f =>
                `${f.name}: ${f.error}`
            ).join('\n');
            archive.append(errorReport, { name: 'errors.txt' });
        }

        await archive.finalize();

    } catch (error) {
        console.error('Processing error:', error);
        res.status(500).json({ error: error.message });
    }
});

// Error handling middleware
app.use((err, req, res, next) => {
    console.error('Server error:', err);
    res.status(500).json({ error: err.message });
});

// Start server
app.listen(PORT, () => {
    console.log(`🚀 DOCX Processor running at http://localhost:${PORT}`);
});

export default app;
