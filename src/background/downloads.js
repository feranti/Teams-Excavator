// Exports and downloads, run by the service worker so they survive the popup
// closing.
//
// Delivery: the worker registers a body under /dl/<id> and hands the URL to
// the popup, which navigates a hidden iframe to it. The worker's fetch handler
// answers with Content-Disposition: attachment, so Chrome's download manager
// takes over and keeps reading from the worker even once the popup is gone
// (the StreamSaver.js technique). chrome.downloads.download() can't be used
// for this: it bypasses the worker's fetch handler for extension URLs.
//
// The Teams tab stays the source of truth: transcripts, attachment lists and
// captured image bytes are pulled from its content script on demand.
(function (TE) {
    'use strict';

    const util = TE.util;

    const FETCH_TIMEOUT_MS = 30000;
    const FETCH_ATTEMPTS = 3;
    const SERVE_TTL_MS = 5 * 60 * 1000;

    // --- the Teams tab ---

    async function askTab(tabId, message) {
        let res;
        try {
            res = await chrome.tabs.sendMessage(tabId, message);
        } catch (_) {
            res = null;
        }
        if (!res) throw new Error('The Teams tab did not answer. Keep it open and try again.');
        if (res.error) throw new Error(res.error);
        return res;
    }

    // Map<imageId, dataUrl> for every captured image referenced by `messages`.
    // One message per image keeps each reply far below the messaging limit.
    async function loadImages(tabId, messages) {
        const images = new Map();
        for (let i = 0; i < messages.length; i++) {
            const attachments = messages[i].attachments;
            for (let a = 0; a < attachments.length; a++) {
                const id = attachments[a].imageId;
                if (!id || images.has(id)) continue;
                try {
                    const res = await askTab(tabId, { type: 'getImage', id: id });
                    if (res.dataUrl) images.set(id, res.dataUrl);
                } catch (_) {
                    // Missing image: the export falls back to the file name.
                }
            }
        }
        return images;
    }

    // --- served bodies ---

    const served = new Map(); // id -> { filename, blob } | { filename, open: () => ReadableStream }

    function serve(entry) {
        const id = crypto.randomUUID();
        served.set(id, entry);
        setTimeout(() => served.delete(id), SERVE_TTL_MS);
        return { id: id, url: chrome.runtime.getURL('dl/' + id) };
    }

    function contentDisposition(filename) {
        const ascii = filename.replace(/[^\x20-\x7e]/g, '_').replace(/["\\]/g, '_');
        return 'attachment; filename="' + ascii + '"; filename*=UTF-8\'\'' + encodeURIComponent(filename);
    }

    // Fetch handler hook: a Response for /dl/<id>, or null when not ours.
    function handleFetch(url) {
        if (!url.pathname.startsWith('/dl/')) return null;
        const id = url.pathname.slice(4);
        const entry = served.get(id);
        if (!entry) return new Response('This download has expired.', { status: 404 });
        served.delete(id);
        const body = entry.blob || entry.open();
        return new Response(body, {
            headers: {
                'Content-Type': entry.blob ? (entry.blob.type || 'application/octet-stream') : 'application/zip',
                'Content-Disposition': contentDisposition(entry.filename),
                'Cache-Control': 'no-store'
            }
        });
    }

    // --- transcripts ---

    function exportName(conversationName) {
        return util.dateStamp(new Date()) + ' - ' + util.safeFileName(conversationName, 'Teams chat', 90);
    }

    function buildFiles(formats, conversation, opts, pathPrefix, files) {
        for (let i = 0; i < formats.length; i++) {
            const file = TE.exporters.build(formats[i], conversation, opts);
            files.push({ path: pathPrefix + '.' + file.ext, data: file.content, mime: file.mime });
        }
    }

    // One format: a single file. Several: a ZIP with one file per format.
    async function exportTranscript(tabId, formats, includeFiles, labels) {
        const result = await askTab(tabId, { type: 'getResult' });
        const images = includeFiles && formats.indexOf('html') !== -1 ? await loadImages(tabId, result.messages) : null;
        const conversation = { title: result.conversationName, messages: result.messages };
        const opts = { includeFiles: includeFiles, labels: labels, images: images };
        const base = exportName(result.conversationName);

        const files = [];
        buildFiles(formats, conversation, opts, base, files);
        if (files.length === 1) {
            return serve({ filename: files[0].path, blob: new Blob([files[0].data], { type: files[0].mime }) });
        }
        return serve({ filename: base + '.zip', blob: TE.zip.zipToBlob(files) });
    }

    // Every chat of the finished batch, one folder per chat, in one ZIP.
    async function exportBatch(tabId, formats, includeFiles, labels) {
        const summary = await askTab(tabId, { type: 'getBatchSummary' });
        const now = new Date();
        const root = 'Teams Batch Export - ' + util.dateStamp(now).replace(/-/g, '') + ' - ' + util.timeStamp(now);
        const usedFolders = new Set();
        const files = [];
        let exported = 0;

        for (let i = 0; i < summary.items.length; i++) {
            if (!summary.items[i].messageCount) continue;
            const item = await askTab(tabId, { type: 'getBatchItem', index: i });
            const images = includeFiles && formats.indexOf('html') !== -1 ? await loadImages(tabId, item.messages) : null;
            const folder = util.dedupeName(usedFolders, util.safeFileName(item.name, 'chat', 90));
            const conversation = { title: item.name, messages: item.messages };
            buildFiles(formats, conversation, { includeFiles: includeFiles, labels: labels, images: images },
                root + '/' + folder + '/' + folder, files);
            exported++;
        }
        if (!exported) return { url: null, count: 0 };

        const delivery = serve({ filename: root + '.zip', blob: TE.zip.zipToBlob(files) });
        askTab(tabId, { type: 'markBatchExported' }).catch(() => {});
        return { url: delivery.url, count: exported };
    }

    // --- attachments ---

    // Formats that are already compressed (or grow under DEFLATE): stored.
    const STORED_EXTENSIONS = new Set([
        'jpg', 'jpeg', 'png', 'gif', 'webp', 'heic', 'heif', 'avif', 'tif', 'tiff',
        'mkv', 'mp4', 'webm', 'mov', 'avi', 'wmv', 'm4v',
        'mp3', 'm4a', 'aac', 'opus', 'ogg', 'flac', 'wav',
        'zip', 'rar', '7z', 'gz', 'xz', 'bz2', 'tgz', 'tbz2',
        'pdf', 'docx', 'xlsx', 'pptx', 'odt', 'ods', 'odp', 'jar', 'apk', 'ipa'
    ]);

    // Matched against our own "HTTP nnn" errors and chrome.downloads
    // interrupt reasons. Narrow on purpose so network blips still retry.
    const PERMANENT_ERROR_RE = /forbidden|unauthorized|access denied|not found|invalid[_ ]url|invalid filename|name_too_long|bad_content|no_space/i;

    function isFatalFailure(message) {
        if (PERMANENT_ERROR_RE.test(message || '')) return true;
        const http = /http (\d{3})/i.exec(message || '');
        if (!http) return false;
        const status = Number(http[1]);
        // 4xx are final, except timeout and rate limiting.
        return status >= 400 && status < 500 && status !== 408 && status !== 429;
    }

    function delay(ms) {
        return new Promise((resolve) => setTimeout(resolve, ms));
    }

    // Fetch with a connection timeout and bounded retries. Only the request is
    // retried, never a body already being streamed into the archive.
    async function fetchWithRetry(url) {
        let lastError = null;
        for (let attempt = 1; attempt <= FETCH_ATTEMPTS; attempt++) {
            const abort = new AbortController();
            const timer = setTimeout(() => abort.abort(), FETCH_TIMEOUT_MS);
            try {
                const res = await fetch(url, { credentials: 'include', signal: abort.signal });
                clearTimeout(timer);
                if (!res.ok) throw new Error('HTTP ' + res.status);
                if (!res.body) throw new Error('Empty response');
                return res;
            } catch (e) {
                clearTimeout(timer);
                lastError = e.name === 'AbortError' ? new Error('Connection timed out') : e;
                if (attempt === FETCH_ATTEMPTS || isFatalFailure(lastError.message)) break;
                await delay(1500 * attempt);
            }
        }
        throw lastError;
    }

    let job = null; // the attachments download in progress or last finished
    let keepAliveTimer = null;

    function snapshot() {
        if (!job) return null;
        return {
            kind: job.kind,
            status: job.status,
            done: job.done,
            total: job.files.length,
            current: job.current,
            errors: job.errors
        };
    }

    function broadcast() {
        chrome.runtime.sendMessage({ type: 'dl:progress', job: snapshot() }).catch(() => {});
    }

    // An extension service worker is stopped after ~30 s without extension
    // events; an API call every 20 s keeps it up while a job is running.
    function updateKeepAlive() {
        const busy = job !== null && job.status === 'running';
        if (busy && !keepAliveTimer) {
            keepAliveTimer = setInterval(() => chrome.runtime.getPlatformInfo(() => {}), 20000);
        } else if (!busy && keepAliveTimer) {
            clearInterval(keepAliveTimer);
            keepAliveTimer = null;
        }
    }

    function finishJob(j, status) {
        if (j.status !== 'running') return;
        j.status = status;
        j.current = '';
        broadcast();
        updateKeepAlive();
    }

    async function resolveFileUrl(j, file) {
        if (!file.imageId) return file.url;
        const res = await askTab(j.tabId, { type: 'getImage', id: file.imageId });
        if (!res.dataUrl) throw new Error('Image no longer available');
        return res.dataUrl;
    }

    // Start downloading the current result's attachments, as one streamed ZIP
    // or as individual files in a dated folder.
    async function startFiles(tabId, asZip) {
        if (job && job.status === 'running') {
            const err = new Error('A file download is already in progress');
            err.code = 'busy';
            throw err;
        }
        const res = await askTab(tabId, { type: 'getFiles' });
        if (!res.files.length) throw new Error('No downloadable files');

        const base = exportName(res.conversationName);
        job = {
            kind: asZip ? 'zip' : 'files',
            tabId: tabId,
            files: res.files,
            status: 'running',
            done: 0,
            current: '',
            errors: [],
            cancelled: false,
            servedId: null,
            downloadId: null,
            started: false
        };
        const j = job;
        updateKeepAlive();
        broadcast();

        if (!asZip) {
            downloadEach(j, base);
            return { url: null };
        }

        const entry = serve({ filename: base + ' - files.zip', open: () => zipStream(j) });
        j.servedId = entry.id;
        // If the popup closed before navigating to the URL, nothing will ever
        // pull the stream: don't leave the job blocking the next one.
        setTimeout(() => { if (!j.started) finishJob(j, 'failed'); }, 60000);
        return { url: entry.url };
    }

    function stopFiles() {
        if (!job || job.status !== 'running') return;
        job.cancelled = true;
        if (job.downloadId !== null) chrome.downloads.cancel(job.downloadId).catch(() => {});
        if (job.wake) job.wake();
    }

    // Pair the browser download created by our /dl/ response with the job, so
    // "Stop" can cancel it.
    function onDownloadCreated(item) {
        if (!job || !job.servedId || !item.url || item.url.indexOf('/dl/' + job.servedId) === -1) return;
        job.downloadId = item.id;
    }

    // ReadableStream fed by the archive writer, with backpressure: writes wait
    // while the download manager is behind.
    function zipStream(j) {
        j.started = true;
        let controller = null;
        let resume = null;
        j.wake = () => {
            if (resume) {
                const r = resume;
                resume = null;
                r();
            }
        };
        const out = {
            async write(chunk) {
                if (j.cancelled) return;
                controller.enqueue(chunk);
                if (controller.desiredSize <= 0) await new Promise((r) => { resume = r; });
            }
        };
        const stream = new ReadableStream({
            start(c) { controller = c; },
            pull() { j.wake(); },
            cancel() {
                j.cancelled = true;
                j.wake();
            }
        }, new CountQueuingStrategy({ highWaterMark: 16 }));

        writeZip(j, out).then(() => {
            try { controller.close(); } catch (_) { /* already cancelled */ }
            finishJob(j, 'done');
        }, (e) => {
            try { controller.error(e); } catch (_) { /* already cancelled */ }
            finishJob(j, j.cancelled ? 'stopped' : 'failed');
        });
        return stream;
    }

    async function writeZip(j, out) {
        const zip = new TE.zip.ZipWriter();
        const usedNames = new Set();
        for (let i = 0; i < j.files.length; i++) {
            if (j.cancelled) break;
            const file = j.files[i];
            j.current = file.name;
            broadcast();
            try {
                await addToZip(j, zip, out, file, util.dedupeName(usedNames, util.safeFileName(file.name, 'file', 200)));
                j.done++;
            } catch (e) {
                if (j.cancelled) break;
                j.errors.push({ name: file.name, error: e.message || String(e) });
            }
        }
        if (j.cancelled) throw new Error('Cancelled');
        await out.write(zip.finish());
    }

    // The local header is written lazily, on the first byte of data: a file
    // whose request fails leaves no trace in the archive. A failure mid-body
    // still closes the entry, so the archive stays valid (that file is cut).
    async function addToZip(j, zip, out, file, name) {
        const res = await fetchWithRetry(await resolveFileUrl(j, file));
        const deflate = !STORED_EXTENSIONS.has(util.fileExtension(name));
        const entry = zip.openStreamEntry(name, deflate);
        let headerWritten = false;

        let source = res.body;
        if (deflate) {
            source = source
                .pipeThrough(new TransformStream({
                    transform(chunk, c) {
                        entry.consume(chunk);
                        c.enqueue(chunk);
                    }
                }))
                .pipeThrough(new CompressionStream('deflate-raw'));
        }
        const reader = source.getReader();
        try {
            for (;;) {
                if (j.cancelled) {
                    reader.cancel().catch(() => {});
                    throw new Error('Cancelled');
                }
                const { done, value } = await reader.read();
                if (done) break;
                if (!value || !value.length) continue;
                if (!deflate) entry.consume(value);
                if (!headerWritten) {
                    headerWritten = true;
                    await out.write(entry.header);
                }
                entry.produce(value.length);
                await out.write(value);
            }
        } catch (e) {
            if (!headerWritten) {
                entry.abort();
                throw e;
            }
            await out.write(entry.close());
            throw new Error((e.message || String(e)) + ' (file truncated)');
        }
        if (!headerWritten) await out.write(entry.header); // empty stored file
        await out.write(entry.close());
    }

    // One chrome.downloads call per file, sequentially, with retries.
    async function downloadEach(j, folder) {
        const usedNames = new Set();
        for (let i = 0; i < j.files.length; i++) {
            if (j.cancelled) break;
            const file = j.files[i];
            j.current = file.name;
            broadcast();
            const path = folder + '/' + util.dedupeName(usedNames, util.safeFileName(file.name, 'file', 200));
            let error = null;
            for (let attempt = 1; attempt <= FETCH_ATTEMPTS; attempt++) {
                try {
                    error = await downloadOne(await resolveFileUrl(j, file), path);
                } catch (e) {
                    error = e.message || String(e);
                }
                if (!error || j.cancelled || isFatalFailure(error)) break;
                if (attempt < FETCH_ATTEMPTS) await delay(2000);
            }
            if (error) j.errors.push({ name: file.name, error: error });
            else j.done++;
        }
        finishJob(j, j.cancelled ? 'stopped' : 'done');
    }

    // Resolves null on success, or the error message.
    function downloadOne(url, filename) {
        return new Promise((resolve) => {
            chrome.downloads.download({ url: url, filename: filename, conflictAction: 'uniquify' }, (id) => {
                if (chrome.runtime.lastError || id === undefined) {
                    resolve(chrome.runtime.lastError ? chrome.runtime.lastError.message : 'Download refused');
                    return;
                }
                const timer = setTimeout(() => done('Download timed out after 5 minutes'), 5 * 60 * 1000);
                function done(result) {
                    clearTimeout(timer);
                    chrome.downloads.onChanged.removeListener(listener);
                    resolve(result);
                }
                function listener(delta) {
                    if (delta.id !== id || !delta.state) return;
                    if (delta.state.current === 'complete') done(null);
                    else if (delta.state.current === 'interrupted') done(delta.error ? delta.error.current : 'Download interrupted');
                }
                chrome.downloads.onChanged.addListener(listener);
            });
        });
    }

    // --- inline image bytes for the content script ---

    function encodeBase64(bytes) {
        let binary = '';
        for (let i = 0; i < bytes.length; i += 0x8000) {
            binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 0x8000));
        }
        return btoa(binary);
    }

    // Fetch an image with the extension's host permissions and the user's
    // cookies (no page CORS here) and return its exact bytes as a data: URL.
    async function fetchImage(url) {
        if (typeof url !== 'string' || !url.startsWith('https://')) return { ok: false };
        const abort = new AbortController();
        const timer = setTimeout(() => abort.abort(), FETCH_TIMEOUT_MS);
        try {
            const res = await fetch(url, { credentials: 'include', signal: abort.signal });
            const type = util.responseMime(res);
            // asm.skype.com may serve images as octet-stream: only reject what is
            // clearly not an image (an HTML/JSON error page).
            if (!res.ok || util.isErrorPageMime(type)) return { ok: false };
            const bytes = new Uint8Array(await res.arrayBuffer());
            if (bytes.byteLength === 0) return { ok: false };
            const mime = /^image\/[a-z0-9.+-]+$/.test(type) ? type : 'image/png';
            return { ok: true, dataUrl: 'data:' + mime + ';base64,' + encodeBase64(bytes) };
        } catch (e) {
            return { ok: false, error: e.message };
        } finally {
            clearTimeout(timer);
        }
    }

    TE.downloads = {
        handleFetch: handleFetch,
        exportTranscript: exportTranscript,
        exportBatch: exportBatch,
        startFiles: startFiles,
        stopFiles: stopFiles,
        state: snapshot,
        onDownloadCreated: onDownloadCreated,
        fetchImage: fetchImage
    };
})(globalThis.TE = globalThis.TE || {});
