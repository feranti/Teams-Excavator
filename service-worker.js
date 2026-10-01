// Service worker: exports, downloads and the popup/content-script bridge.
//
// It sits at the extension root on purpose: its fetch handler must control
// the chrome-extension://<id>/dl/... URLs it serves downloads from, and a
// worker only controls URLs under its own directory.
//
// importScripts must run during the worker's first evaluation (MV3 refuses
// later imports of scripts it has not cached).
importScripts(
    'src/shared/util.js',
    'src/background/zip.js',
    'src/background/exporters.js',
    'src/background/downloads.js'
);

(function (TE) {
    'use strict';

    const downloads = TE.downloads;

    self.addEventListener('fetch', (event) => {
        const response = downloads.handleFetch(new URL(event.request.url));
        if (response) event.respondWith(response);
    });

    chrome.downloads.onCreated.addListener(downloads.onDownloadCreated);

    // Chrome only injects content scripts into pages loaded after install or
    // update. Inject into already-open Teams tabs so they work without a reload
    // (the scripts dispose of any stale copy themselves).
    chrome.runtime.onInstalled.addListener(async () => {
        const scripts = chrome.runtime.getManifest().content_scripts || [];
        for (let i = 0; i < scripts.length; i++) {
            const tabs = await chrome.tabs.query({ url: scripts[i].matches });
            for (let t = 0; t < tabs.length; t++) {
                chrome.scripting.executeScript({ target: { tabId: tabs[t].id }, files: scripts[i].js }).catch(() => {});
            }
        }
    });

    // Async handlers: resolve to the response, or throw to answer { error }.
    const handlers = {
        fetchImage: (msg) => downloads.fetchImage(msg.url),
        'export:transcript': (msg) => downloads.exportTranscript(msg.tabId, msg.formats, msg.includeFiles, msg.labels),
        'export:batch': (msg) => downloads.exportBatch(msg.tabId, msg.formats, msg.includeFiles, msg.labels),
        'files:start': (msg) => downloads.startFiles(msg.tabId, msg.zip),
        'files:stop': () => downloads.stopFiles(),
        'files:state': () => ({ job: downloads.state() })
    };

    chrome.runtime.onMessage.addListener((msg, sender, sendResponse) => {
        const handler = msg && handlers[msg.type];
        if (!handler) return false; // not for us (e.g. popup-bound broadcasts)
        Promise.resolve()
            .then(() => handler(msg, sender))
            .then(
                (result) => sendResponse(result || {}),
                (e) => sendResponse({ error: e.message || String(e), code: e.code })
            );
        return true;
    });
})(globalThis.TE);
