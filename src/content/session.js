// State of one extraction (one chat or channel) and the captured-image store.
(function (TE) {
    'use strict';

    // Captured inline image bytes, as data: URLs, keyed by id. Messages only
    // carry the id: the bytes stay in the tab until an export asks for them,
    // so no single extension message has to carry megabytes of base64.
    class ImageStore {
        constructor(limitBytes) {
            this.limit = limitBytes;
            this.map = new Map();
            this.bytes = 0;
            this.nextId = 1;
        }
        get full() {
            return this.bytes >= this.limit;
        }
        add(dataUrl) {
            const id = 'img' + this.nextId++;
            this.map.set(id, dataUrl);
            this.bytes += dataUrl.length;
            return id;
        }
        get(id) {
            return this.map.get(id) || null;
        }
        clear() {
            this.map.clear();
            this.bytes = 0;
        }
    }

    // 'YYYY-MM-DD' from <input type="date"> → local midnight in ms, or NaN.
    function parseThreshold(value) {
        const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(value || '');
        return m ? new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3])).getTime() : NaN;
    }

    function timeOf(message) {
        return message.timestamp ? Date.parse(message.timestamp) : NaN;
    }

    class Session {
        // config: { extractFiles, currentUserName, thresholdDate, unknownAuthor }
        constructor(config, images, onProgress) {
            this.config = config;
            this.images = images;
            this.onProgress = onProgress;
            this.userName = config.currentUserName || '';
            this.threshold = parseThreshold(config.thresholdDate);

            this.messages = new Map();     // id -> message, in discovery order
            this.skipped = new Set();      // ids that failed to parse: never retried
            this.fileKeys = new Set();     // unique attachments seen, for progress
            this.attachments = new Map();  // message id -> { list, keys }
            this.imageSignatures = new Map();
            this.imageSourcesTried = new Map();
            this.captures = [];            // in-flight image captures
            this.drillAttempts = new Map();
            this.threads = new Map();      // channel thread id -> { replies, latest }
            this.oldest = Infinity;
            this.stopRequested = false;
        }

        // Author text as shown by Teams, or '' when missing or when it is the
        // localized "unknown" label (compared in lower case).
        authorName(el) {
            const text = el ? el.textContent.trim() : '';
            const unknown = this.config.unknownAuthor;
            return unknown && text.toLowerCase() === unknown ? '' : text;
        }

        seen(id) {
            return this.messages.has(id) || this.skipped.has(id);
        }

        add(message) {
            this.messages.set(message.id, message);
            const t = timeOf(message);
            if (t < this.oldest) this.oldest = t;
            if (message.threadId) {
                let thread = this.threads.get(message.threadId);
                if (!thread) {
                    thread = { replies: 0, latest: -Infinity };
                    this.threads.set(message.threadId, thread);
                }
                if (message.isReply) thread.replies++;
                if (t > thread.latest) thread.latest = t;
            }
            if (this.onProgress) this.onProgress(this);
        }

        get halted() {
            return this.stopRequested;
        }

        // We scroll newest → oldest: once a message at/before the threshold has
        // been seen, there is nothing newer left to find.
        reachedThreshold() {
            return !isNaN(this.threshold) && this.oldest <= this.threshold;
        }

        keepAfterThreshold(messages) {
            if (isNaN(this.threshold)) return messages;
            const kept = [];
            for (let i = 0; i < messages.length; i++) {
                const t = timeOf(messages[i]);
                if (isNaN(t) || t > this.threshold) kept.push(messages[i]);
            }
            return kept;
        }

        // Per-message attachment state. It outlives a sighting: Teams recycles
        // message nodes and loads images lazily, so a message is re-scanned
        // every time it is seen and new finds are merged into the same list.
        attachmentState(messageId) {
            let state = this.attachments.get(messageId);
            if (!state) {
                state = { list: [], keys: new Set() };
                this.attachments.set(messageId, state);
            }
            return state;
        }

        // Run an async capture in the background, bounded so a stuck fetch or
        // decode can never hold up the end of the extraction.
        capture(promise) {
            this.captures.push(TE.withTimeout(promise.catch((e) => TE.log.warn('Image capture failed:', e)), 15000));
        }

        async settleCaptures() {
            if (!this.captures.length) return;
            TE.log.info('Waiting for ' + this.captures.length + ' image capture(s)…');
            await TE.withTimeout(Promise.allSettled(this.captures), 30000);
            if (this.images.full) TE.log.warn('Image memory cap reached: remaining images are kept as links only.');
        }
    }

    TE.ImageStore = ImageStore;
    TE.Session = Session;
})(globalThis.TE = globalThis.TE || {});
