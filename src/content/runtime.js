// Content-script runtime: page gate, logger and throttle-proof sleep.
(function (TE) {
    'use strict';

    // The manifest has to match every *.mcas.ms host to reach Teams behind
    // Defender for Cloud Apps; gating on the host keeps the scripts inert on
    // the other proxied apps.
    TE.isTeamsPage = TE.util.isTeamsHost(location.hostname);

    // --- sleep ---
    // Chrome throttles the timers of a hidden tab to one wake-up per minute.
    // Extraction runs unattended for minutes while the user works elsewhere,
    // so every delay goes through a dedicated worker, whose timers are not
    // throttled. Falls back to window timers if the page CSP blocks blob:
    // workers. Kept across re-injections so workers don't pile up.
    if (!TE.sleep) TE.sleep = createSleep();

    function createSleep() {
        const pending = new Map();
        let nextId = 1;
        let worker = null;
        try {
            const src = 'onmessage=function(e){setTimeout(function(){postMessage(e.data.id)},e.data.ms)}';
            worker = new Worker(URL.createObjectURL(new Blob([src], { type: 'text/javascript' })));
            worker.onmessage = (e) => {
                const resolve = pending.get(e.data);
                if (!resolve) return;
                pending.delete(e.data);
                resolve();
            };
            worker.onerror = () => {
                worker = null;
                pending.forEach((resolve) => setTimeout(resolve, 0));
                pending.clear();
            };
        } catch (_) {
            worker = null;
        }
        return function sleep(ms) {
            return new Promise((resolve) => {
                if (!worker) {
                    setTimeout(resolve, ms);
                    return;
                }
                const id = nextId++;
                pending.set(id, resolve);
                worker.postMessage({ id: id, ms: ms });
            });
        };
    }

    // Resolves with the promise's value, or undefined after `ms`.
    TE.withTimeout = function (promise, ms) {
        return Promise.race([promise, TE.sleep(ms)]);
    };

    // --- logging ---
    // Mirrors to the page console and keeps a ring buffer the popup's
    // "Console" panel shows; the controller streams new entries to it.

    const LOG_LIMIT = 500;
    if (!TE.logEntries) TE.logEntries = [];

    function describe(value) {
        if (typeof value === 'string') return value;
        if (value instanceof Error) return value.message;
        try {
            return JSON.stringify(value, (key, v) =>
                typeof v === 'string' && v.length > 200 && v.startsWith('data:') ? v.slice(0, 40) + '…' : v);
        } catch (_) {
            return String(value);
        }
    }

    function write(level, args) {
        const sink = level === 'error' ? console.error : level === 'warn' ? console.warn : console.log;
        sink.apply(console, ['[Teams Excavator]'].concat(args));

        let text = '';
        for (let i = 0; i < args.length; i++) text += (i ? ' ' : '') + describe(args[i]);
        const entry = { time: Date.now(), level: level, text: text };
        const entries = TE.logEntries;
        entries.push(entry);
        if (entries.length > LOG_LIMIT) entries.splice(0, entries.length - LOG_LIMIT);
        if (TE.log.onEntry) TE.log.onEntry(entry);
    }

    TE.log = {
        onEntry: null,
        info: function () { write('log', Array.prototype.slice.call(arguments)); },
        warn: function () { write('warn', Array.prototype.slice.call(arguments)); },
        error: function () { write('error', Array.prototype.slice.call(arguments)); }
    };

    // Fire-and-forget message to the extension. Throws synchronously when this
    // copy of the script outlived an extension reload; nothing to do then.
    TE.send = function (message) {
        try {
            return chrome.runtime.sendMessage(message);
        } catch (_) {
            return Promise.resolve(null);
        }
    };
})(globalThis.TE = globalThis.TE || {});
