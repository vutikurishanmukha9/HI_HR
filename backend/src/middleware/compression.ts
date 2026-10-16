/**
 * Native HTTP Response Compression Middleware
 * Uses Node.js built-in zlib to compress JSON and text responses
 * with zero external dependencies and zero node_modules overhead.
 */
import { Request, Response, NextFunction } from 'express';
import zlib from 'zlib';

const MIN_COMPRESS_SIZE = 1024; // 1 KB threshold

export function nativeCompressionMiddleware(req: Request, res: Response, next: NextFunction): void {
    const acceptEncoding = req.headers['accept-encoding'] as string || '';

    // If client does not accept compression or request is a HEAD, skip
    if (req.method === 'HEAD' || !acceptEncoding) {
        next();
        return;
    }

    const supportsGzip = acceptEncoding.includes('gzip');
    const supportsDeflate = acceptEncoding.includes('deflate');

    if (!supportsGzip && !supportsDeflate) {
        next();
        return;
    }

    // Capture standard response methods
    const originalSend = res.send.bind(res);
    const originalJson = res.json.bind(res);

    const compressAndSend = (body: any): Response => {
        // Ensure Vary header is set so CDNs and browsers cache correctly
        res.setHeader('Vary', 'Accept-Encoding');

        let data: Buffer;
        if (Buffer.isBuffer(body)) {
            data = body;
        } else if (typeof body === 'string') {
            data = Buffer.from(body, 'utf8');
        } else if (typeof body === 'object' && body !== null) {
            data = Buffer.from(JSON.stringify(body), 'utf8');
            if (!res.getHeader('Content-Type')) {
                res.setHeader('Content-Type', 'application/json; charset=utf-8');
            }
        } else {
            return originalSend(body);
        }

        // Do not compress small payloads (under 1KB) as overhead exceeds savings
        if (data.length < MIN_COMPRESS_SIZE) {
            return originalSend(data);
        }

        try {
            if (supportsGzip) {
                const compressed = zlib.gzipSync(data, { level: 6 });
                res.setHeader('Content-Encoding', 'gzip');
                res.removeHeader('Content-Length');
                return originalSend(compressed);
            } else if (supportsDeflate) {
                const compressed = zlib.deflateSync(data, { level: 6 });
                res.setHeader('Content-Encoding', 'deflate');
                res.removeHeader('Content-Length');
                return originalSend(compressed);
            }
        } catch {
            // Fallback to uncompressed on any compression error
            return originalSend(data);
        }

        return originalSend(data);
    };

    res.send = compressAndSend as any;
    res.json = compressAndSend as any;

    next();
}

export default nativeCompressionMiddleware;
