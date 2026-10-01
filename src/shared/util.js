// Helpers shared by the content scripts, the popup and the service worker.
//
// Every script of the extension is a classic script wrapped in an IIFE that
// publishes its API on one global namespace, `TE`. There is deliberately no
// top-level const/let anywhere: Chrome re-runs content scripts in the same
// isolated world when the extension is reloaded or updated, and redeclaring a
// top-level binding would throw and leave the stale copy in charge.
(function (TE) {
    'use strict';

    // Normalize a date-ish value to ISO 8601 UTC. Returns '' for empty input
    // and the original string when it does not parse, so a locale-specific
    // date never turns into a bogus timestamp.
    function toIso(value) {
        if (value == null || value === '') return '';
        const parsed = value instanceof Date ? value : new Date(value);
        const ms = parsed.getTime();
        return ms === ms ? parsed.toISOString() : String(value);
    }

    function pad2(n) {
        return n < 10 ? '0' + n : String(n);
    }

    // Local calendar date as YYYY-MM-DD (export file name prefix).
    function dateStamp(d) {
        return d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
    }

    // Local time as HHMM.
    function timeStamp(d) {
        return pad2(d.getHours()) + pad2(d.getMinutes());
    }

    function formatDuration(totalSeconds) {
        const m = Math.floor(totalSeconds / 60);
        const s = totalSeconds % 60;
        return m > 0 ? m + 'm ' + s + 's' : s + 's';
    }

    // Index of the dot starting the extension, or -1. A leading dot
    // (".gitignore") is not an extension.
    function extensionDot(name) {
        const at = name.lastIndexOf('.');
        return at > 0 && at < name.length - 1 ? at : -1;
    }

    // Lower-cased extension without the dot, or ''.
    function fileExtension(name) {
        if (!name) return '';
        const dot = extensionDot(name);
        return dot === -1 ? '' : name.slice(dot + 1).toLowerCase();
    }

    const RESERVED_NAME_RE = /^(con|prn|aux|nul|com\d|lpt\d)(\..*)?$/i;

    // Make `name` usable as a file or folder name on Windows, macOS and Linux
    // while keeping it readable. Over-long names are cut before the extension.
    function safeFileName(name, fallback, maxLength) {
        const max = maxLength || 120;
        let s = String(name || '').normalize('NFC')
            .replace(/[<>:"/\\|?*\u0000-\u001f\u007f]/g, '_')
            .replace(/\s+/g, ' ')
            .replace(/_{2,}/g, '_')
            .replace(/^[\s.]+|[\s.]+$/g, '');
        if (RESERVED_NAME_RE.test(s)) s = '_' + s;
        if (s.length > max) {
            const dot = extensionDot(s);
            const ext = dot !== -1 && s.length - dot <= 16 ? s.slice(dot) : '';
            let cut = max - ext.length;
            // Never split a surrogate pair.
            const code = s.charCodeAt(cut - 1);
            if (code >= 0xd800 && code <= 0xdbff) cut--;
            s = s.slice(0, cut).replace(/[\s.]+$/, '') + ext;
        }
        return s || fallback || 'file';
    }

    // Register and return `name`, or "name (2).ext", "name (3).ext"… when it
    // is already in `used`. Case-insensitive, like Windows and macOS.
    function dedupeName(used, name) {
        let key = name.toLowerCase();
        let result = name;
        if (used.has(key)) {
            const cut = extensionDot(name);
            const stem = cut === -1 ? name : name.slice(0, cut);
            const tail = cut === -1 ? '' : name.slice(cut);
            for (let n = 2; used.has(key); n++) {
                result = stem + ' (' + n + ')' + tail;
                key = result.toLowerCase();
            }
        }
        used.add(key);
        return result;
    }

    // Bare MIME type of a fetch Response ("image/png", no parameters).
    function responseMime(res) {
        const raw = res.headers.get('content-type') || '';
        const semi = raw.indexOf(';');
        return (semi === -1 ? raw : raw.slice(0, semi)).trim().toLowerCase();
    }

    // Text-ish payloads are error pages, not file or image bytes.
    function isErrorPageMime(mime) {
        return mime.startsWith('text/') || /html|json|xml/.test(mime);
    }

    // Teams web hosts, including Defender for Cloud Apps proxies that suffix
    // the real host (teams.microsoft.com.mcas.ms, ….us2.mcas.ms).
    function isTeamsHost(hostname) {
        return /(^|\.)teams\.(microsoft\.com|live\.com|cloud\.microsoft)(\.|$)/i.test(hostname || '');
    }

    function isTeamsUrl(url) {
        try {
            return isTeamsHost(new URL(url).hostname);
        } catch (_) {
            return false;
        }
    }

    // --- Plain-text rendering of the message body model (see richtext.js) ---

    // A run is { text, b?, i?, u?, s?, mention?, href? }.
    function runsToText(runs) {
        let out = '';
        for (let i = 0; i < runs.length; i++) {
            const r = runs[i];
            if (r.mention) {
                out += '@' + r.text;
            } else if (r.href && r.text.trim() !== r.href) {
                out += r.text + ' (' + r.href + ')';
            } else {
                out += r.text;
            }
        }
        return out;
    }

    // Reply cards are left out: quotes have their own fields/columns.
    function blocksToText(blocks) {
        const out = [];
        appendBlocksText(blocks, '', out);
        return out.join('\n');
    }

    function appendBlocksText(blocks, prefix, out) {
        for (let i = 0; i < blocks.length; i++) {
            const block = blocks[i];
            switch (block.kind) {
                case 'text':
                    for (let l = 0; l < block.lines.length; l++) out.push(prefix + runsToText(block.lines[l]));
                    break;
                case 'code': {
                    const lines = block.text.split('\n');
                    for (let l = 0; l < lines.length; l++) out.push(prefix + lines[l]);
                    break;
                }
                case 'quote':
                    appendBlocksText(block.blocks, prefix + '> ', out);
                    break;
                case 'list':
                    for (let j = 0; j < block.items.length; j++) {
                        const marker = block.ordered ? (j + 1) + '. ' : '• ';
                        const indent = marker.length === 2 ? '  ' : '   ';
                        const itemLines = [];
                        appendBlocksText(block.items[j], '', itemLines);
                        for (let l = 0; l < itemLines.length; l++) {
                            out.push(prefix + (l === 0 ? marker : indent) + itemLines[l]);
                        }
                    }
                    break;
            }
        }
    }

    TE.util = {
        toIso: toIso,
        dateStamp: dateStamp,
        timeStamp: timeStamp,
        formatDuration: formatDuration,
        fileExtension: fileExtension,
        safeFileName: safeFileName,
        dedupeName: dedupeName,
        responseMime: responseMime,
        isErrorPageMime: isErrorPageMime,
        isTeamsHost: isTeamsHost,
        isTeamsUrl: isTeamsUrl,
        runsToText: runsToText,
        blocksToText: blocksToText
    };
})(globalThis.TE = globalThis.TE || {});
