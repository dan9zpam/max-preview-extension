(function (root) {
    const extensions = new Set('pdf docx xlsx xls pptx odt rtf zip png jpeg jpg txt csv py c cpp java cs go php rb swift ts sh json xml md js css html mp3 wav ogg mp4 webm avi mov mkv'.split(' '));
    const mimeExtensions = {
        'application/pdf': 'pdf', 'application/zip': 'zip', 'application/x-zip-compressed': 'zip',
        'application/vnd.openxmlformats-officedocument.wordprocessingml.document': 'docx',
        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': 'xlsx',
        'application/vnd.ms-excel': 'xls',
        'application/vnd.openxmlformats-officedocument.presentationml.presentation': 'pptx',
        'application/vnd.oasis.opendocument.text': 'odt', 'application/rtf': 'rtf', 'text/rtf': 'rtf',
        'image/png': 'png', 'image/jpeg': 'jpg', 'text/plain': 'txt', 'text/csv': 'csv', 'text/html': 'html',
        'audio/mpeg': 'mp3', 'audio/wav': 'wav', 'audio/ogg': 'ogg',
        'video/mp4': 'mp4', 'video/webm': 'webm', 'video/ogg': 'ogg'
    };
    function clean(name) {
        return String(name || '').split(/[\\/]/).pop().normalize('NFC')
            .replace(/[<>:"|?*\x00-\x1f\x7f]/g, '_').replace(/[. ]+$/, '').trim();
    }
    function extension(name) {
        const ext = /\.([a-z0-9]+)$/i.exec(name || '')?.[1].toLowerCase();
        return extensions.has(ext) ? ext : '';
    }
    function repairEncoding(name) {
        const characters = Array.from(name);
        if (characters.every(char => char.codePointAt(0) <= 255)) {
            try {
                const repaired = new TextDecoder('utf-8', { fatal: true }).decode(Uint8Array.from(characters, char => char.charCodeAt(0)));
                if (repaired !== name) return repaired;
            } catch (_) {}
        }
        // Recover UTF-8 bytes interpreted as Windows-1252 or Windows-1251 by a server.
        for (const encoding of ['iso-8859-1', 'windows-1251']) {
            const decoder = new TextDecoder(encoding);
            const bytes = new Map();
            for (let i = 0; i < 256; i++) bytes.set(decoder.decode(new Uint8Array([i])), i);
            if (!characters.every(char => bytes.has(char))) continue;
            try {
                const repaired = new TextDecoder('utf-8', { fatal: true }).decode(Uint8Array.from(characters, char => bytes.get(char)));
                if (repaired !== name) return repaired;
            } catch (_) {}
        }
        return name;
    }
    function decode(name) {
        try { name = decodeURIComponent(name); } catch (_) {}
        // Some HTTP stacks encode an already misdecoded filename a second time.
        for (let depth = 0; depth < 3; depth++) {
            const repaired = repairEncoding(name);
            if (repaired === name) break;
            name = repaired;
        }
        return name;
    }
    function disposition(header) {
        let plain = '';
        for (const match of header.matchAll(/(?:^|;)\s*(filename\*?)\s*=\s*("(?:\\.|[^"\\])*"|[^;]*)/gi)) {
            const value = match[2].replace(/^"|"$/g, '').replace(/\\(["\\])/g, '$1').trim();
            if (match[1].endsWith('*')) {
                const encoded = /^([^']+)'[^']*'(.*)$/.exec(value);
                if (!encoded) continue;
                try {
                    // RFC 8187: the language field may be nonempty.
                    const data = Uint8Array.from(encoded[2].match(/%[\da-f]{2}|[^%]/gi) || [], part =>
                        part[0] === '%' ? parseInt(part.slice(1), 16) : part.charCodeAt(0));
                    if (/%(?![\da-f]{2})/i.test(encoded[2])) continue;
                    return clean(new TextDecoder(encoded[1], { fatal: true }).decode(data));
                } catch (_) {}
            } else plain = clean(decode(value));
        }
        return plain;
    }
    function fromUrl(url) {
        try {
            const parsed = new URL(url);
            for (const key of ['response-content-disposition', 'content-disposition']) {
                const name = disposition(parsed.searchParams.get(key) || '');
                if (name) return name;
            }
            for (const key of ['filename', 'fileName', 'name']) {
                const name = clean(decode(parsed.searchParams.get(key) || ''));
                if (extension(name)) return name;
            }
            return clean(decode(parsed.pathname.split('/').pop()));
        } catch (_) { return ''; }
    }
    function quality(name) {
        if (!name) return -10;
        const stem = extension(name) ? name.slice(0, name.lastIndexOf('.')) : name;
        const generic = /^(document|download|attachment|file)$/i.test(stem) || extensions.has(stem.toLowerCase());
        const opaque = /^[a-f\d-]{16,}$/i.test(stem) || /^[a-z\d_-]{32,}$/i.test(stem);
        return (generic || opaque ? 0 : 4) + (extension(name) ? 2 : 0);
    }
    function resolve({ provided = '', header = '', url = '', mime = '', detected = '' }) {
        const names = [clean(provided), disposition(header), fromUrl(url)];
        // Prefer a readable original name over a storage ID from HTTP headers.
        const best = quality(names[0]) >= 4 ? names[0] :
            names.reduce((selected, name) => quality(name) > quality(selected) ? name : selected, '');
        const ext = detected || extension(best) || names.map(extension).find(Boolean) || mimeExtensions[mime.split(';')[0]] || '';
        let result = best || 'Document';
        if (ext && extension(result) !== ext) {
            if (extension(result)) result = result.slice(0, result.lastIndexOf('.'));
            result += '.' + ext;
        }
        return clean(result) || 'Document';
    }
    function mimeFor(name) {
        const ext = extension(name);
        if (ext === 'jpeg') return 'image/jpeg';
        return Object.entries(mimeExtensions).find(([, value]) => value === ext)?.[0] || 'application/octet-stream';
    }
    const api = { clean, extension, decode, disposition, fromUrl, resolve, mimeFor };
    root.MAXFileNames = api;
    if (typeof module !== 'undefined') module.exports = api;
})(globalThis);
