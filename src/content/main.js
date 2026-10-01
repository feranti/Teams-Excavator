// Content-script controller: runs extractions in the Teams tab and holds
// their results, which makes the tab the single source of truth. The popup
// connects a port to follow state and logs live; the popup and the service
// worker send requests with chrome.tabs.sendMessage.
(function (TE) {
    'use strict';

    if (!TE.isTeamsPage) return;

    const IMAGE_MEMORY_CAP = 150 * 1024 * 1024;
    // Waiting for Teams to open a clicked sidebar item. Only time spent with
    // the tab visible counts: a hidden tab is throttled and Teams stalls.
    const SWITCH_POLL_MS = 1000;
    const SWITCH_RECLICK_MS = 10000;
    const SWITCH_GIVE_UP_MS = 60000;

    function normalizeConfig(c) {
        c = c || {};
        return {
            extractFiles: c.extractFiles !== false,
            currentUserName: typeof c.currentUserName === 'string' ? c.currentUserName.trim() : '',
            thresholdDate: typeof c.thresholdDate === 'string' ? c.thresholdDate : '',
            // Compared in lower case against author names and the chat header.
            unknownAuthor: typeof c.unknownAuthor === 'string' ? c.unknownAuthor.trim().toLowerCase() : ''
        };
    }

    class Controller {
        constructor() {
            this.images = new TE.ImageStore(IMAGE_MEMORY_CAP);
            this.busy = null;       // 'single' | 'batch' | null
            this.session = null;    // extraction in progress
            this.startedAt = 0;
            this.result = null;     // last single extraction
            this.error = '';
            this.batch = null;      // last or running batch
            this.ports = new Set();
            this.notifyPending = false;

            this.onMessage = this.onMessage.bind(this);
            this.onConnect = this.onConnect.bind(this);
            chrome.runtime.onMessage.addListener(this.onMessage);
            chrome.runtime.onConnect.addListener(this.onConnect);
            TE.log.onEntry = (entry) => this.post({ type: 'log', entry: entry });
        }

        // Called on the previous copy when the scripts are injected again
        // (extension reload/update): stop its work and detach it.
        dispose() {
            if (this.session) this.session.stopRequested = true;
            if (this.batch) this.batch.stopRequested = true;
            try {
                chrome.runtime.onMessage.removeListener(this.onMessage);
                chrome.runtime.onConnect.removeListener(this.onConnect);
            } catch (_) { /* extension context already gone */ }
            this.ports.forEach((port) => {
                try { port.disconnect(); } catch (_) { /* gone */ }
            });
            this.ports.clear();
        }

        // --- popup port ---

        onConnect(port) {
            if (port.name !== 'popup') return;
            this.ports.add(port);
            port.onDisconnect.addListener(() => this.ports.delete(port));
            port.postMessage({ type: 'state', state: this.publicState() });
            port.postMessage({ type: 'logs', entries: TE.logEntries });
        }

        post(message) {
            this.ports.forEach((port) => {
                try {
                    port.postMessage(message);
                } catch (_) {
                    this.ports.delete(port);
                }
            });
        }

        // Throttled: progress fires on every extracted message.
        notify() {
            if (this.notifyPending || !this.ports.size) return;
            this.notifyPending = true;
            TE.sleep(250).then(() => {
                this.notifyPending = false;
                this.post({ type: 'state', state: this.publicState() });
            });
        }

        notifyNow() {
            this.post({ type: 'state', state: this.publicState() });
        }

        publicState() {
            const s = this.session;
            const b = this.batch;
            let batch = null;
            if (b) {
                let exportable = 0;
                for (let i = 0; i < b.items.length; i++) {
                    if (b.items[i].messages && b.items[i].messages.length) exportable++;
                }
                const current = b.items[b.index];
                batch = {
                    total: b.items.length,
                    index: b.index,
                    current: current ? current.name : '',
                    currentStatus: current ? current.status : '',
                    finished: b.finished,
                    stopRequested: b.stopRequested,
                    exportable: exportable,
                    exported: b.exported
                };
            }
            return {
                busy: this.busy,
                stopping: !!s && s.stopRequested,
                startedAt: this.startedAt,
                progress: s ? { messages: s.messages.size, files: s.fileKeys.size } : null,
                result: this.result ? {
                    conversationName: this.result.conversationName,
                    messageCount: this.result.messages.length,
                    fileCount: this.result.fileCount,
                    durationMs: this.result.durationMs
                } : null,
                error: this.error,
                batch: batch
            };
        }

        // --- requests ---

        onMessage(msg, sender, sendResponse) {
            const handler = msg && this.handlers[msg.type];
            if (!handler) return false;
            let response;
            try {
                response = handler.call(this, msg);
            } catch (e) {
                response = { error: e.message || String(e) };
            }
            sendResponse(response || {});
            return false;
        }

        requireResult() {
            if (!this.result) throw new Error('Nothing extracted yet');
            return this.result;
        }

        requireBatch() {
            if (!this.batch) throw new Error('No batch extraction');
            return this.batch;
        }

        // --- runs ---

        // Forget previous results (and their captured images).
        reset() {
            this.result = null;
            this.error = '';
            this.batch = null;
            this.images.clear();
        }

        newSession(config) {
            const c = Object.assign({}, config);
            if (!c.currentUserName) c.currentUserName = TE.conversation.currentUserName();
            this.session = new TE.Session(c, this.images, () => this.notify());
            return this.session;
        }

        // Crawl the open chat or channel. Returns { name, messages }.
        async extract(session) {
            const messages = TE.channel.onChannelSurface() ? await TE.channel.crawl(session) : await TE.chat.crawl(session);
            TE.richtext.resolveQuotes(messages);
            const name = TE.conversation.name(session, messages);
            const kept = session.keepAfterThreshold(messages);
            if (kept.length < messages.length) {
                TE.log.info('Dropped ' + (messages.length - kept.length) + ' message(s) older than the "stop at date" threshold.');
            }
            return { name: name, messages: kept };
        }

        async runSingle(config) {
            this.reset();
            this.busy = 'single';
            this.startedAt = Date.now();
            const session = this.newSession(config);
            this.notifyNow();
            try {
                const extracted = await this.extract(session);
                this.result = {
                    conversationName: extracted.name,
                    messages: extracted.messages,
                    fileCount: TE.attachments.collectFiles(extracted.messages).length,
                    durationMs: Date.now() - this.startedAt
                };
                TE.log.info('Extracted ' + extracted.messages.length + ' message(s) from "' + extracted.name + '".');
            } catch (e) {
                this.error = e.message || String(e);
                TE.log.error('Extraction failed:', e);
            }
            this.session = null;
            this.busy = null;
            this.notifyNow();
        }

        // chats: [{ id, name }] from the sidebar listing.
        async runBatch(chats, config) {
            this.reset();
            this.busy = 'batch';
            this.startedAt = Date.now();
            const batch = { items: [], index: 0, finished: false, stopRequested: false, exported: false };
            for (let i = 0; i < chats.length; i++) {
                batch.items.push({ id: chats[i].id, name: chats[i].name || chats[i].id, status: 'pending', messages: null, error: '' });
            }
            this.batch = batch;
            // Message id -> index of the batch item holding it. A stalled Teams
            // can leave the previous conversation (or part of it) on screen.
            const owners = new Map();

            for (let i = 0; i < batch.items.length && !batch.stopRequested; i++) {
                const item = batch.items[i];
                batch.index = i;
                item.status = 'running';
                this.notifyNow();

                const opened = await this.openBatchItem(batch, item);
                if (batch.stopRequested) {
                    item.status = 'pending';
                    break;
                }
                if (!opened) {
                    item.status = 'error';
                    if (!item.error) item.error = 'Conversation did not open';
                    continue;
                }
                await TE.sleep(2000); // let Teams render the conversation
                if (batch.stopRequested) {
                    item.status = 'pending';
                    break;
                }
                try {
                    const extracted = await this.extract(this.newSession(config));
                    if (extracted.name) item.name = extracted.name;
                    if (this.keepBatchExtraction(batch, owners, i, extracted.messages)) {
                        item.messages = extracted.messages;
                        item.status = 'done';
                        TE.log.info('Batch: ' + extracted.messages.length + ' message(s) from "' + item.name + '".');
                    }
                } catch (e) {
                    item.status = 'error';
                    item.error = e.message || String(e);
                    TE.log.error('Batch: "' + item.name + '" failed:', e);
                }
                this.session = null;
            }
            batch.finished = true;
            this.busy = null;
            this.notifyNow();
        }

        // Click the item's sidebar entry and wait until Teams shows it. Waits
        // for as long as the tab stays hidden; gives up after a minute of
        // visible time without a switch.
        async openBatchItem(batch, item) {
            if (TE.conversation.isSidebarItemSelected(item.id)) return TE.conversation.openSidebarItem(item.id);
            const before = TE.conversation.openState();
            if (!TE.conversation.openSidebarItem(item.id)) {
                item.error = 'Not found in the sidebar';
                TE.log.warn('Batch: "' + item.name + '" is no longer in the sidebar.');
                return false;
            }
            const switched = () => TE.conversation.hasSwitched(before);
            let visibleMs = 0;
            let sinceClickMs = 0;
            let loggedHidden = false;
            while (!batch.stopRequested) {
                if (await TE.dom.waitFor(batch, switched, SWITCH_POLL_MS)) return true;
                if (document.hidden) {
                    if (!loggedHidden) TE.log.warn('Batch: the Teams tab is hidden, Teams is paused. Waiting for it to be visible again.');
                    loggedHidden = true;
                    continue;
                }
                loggedHidden = false;
                visibleMs += SWITCH_POLL_MS;
                sinceClickMs += SWITCH_POLL_MS;
                if (visibleMs >= SWITCH_GIVE_UP_MS) break;
                if (sinceClickMs >= SWITCH_RECLICK_MS) {
                    TE.log.info('Batch: "' + item.name + '" did not open yet, clicking again.');
                    TE.conversation.openSidebarItem(item.id);
                    sinceClickMs = 0;
                }
            }
            if (!batch.stopRequested) TE.log.warn('Batch: "' + item.name + '" did not open.');
            return false;
        }

        // Two batch items sharing a message are one conversation read twice:
        // keep the more complete copy. Returns false when `messages` loses.
        keepBatchExtraction(batch, owners, index, messages) {
            const item = batch.items[index];
            let other = -1;
            for (let i = 0; i < messages.length && other === -1; i++) {
                const owner = owners.get(messages[i].id);
                if (owner !== undefined) other = owner;
            }
            if (other !== -1) {
                const previous = batch.items[other];
                if (previous.messages.length >= messages.length) {
                    item.status = 'error';
                    item.error = 'Duplicate of "' + previous.name + '"';
                    TE.log.warn('Batch: "' + item.name + '" is a copy of "' + previous.name + '" (' + messages.length +
                        ' vs ' + previous.messages.length + ' message(s)); skipped.');
                    return false;
                }
                TE.log.warn('Batch: "' + previous.name + '" was an incomplete copy (' + previous.messages.length +
                    ' message(s)); replaced by "' + item.name + '" (' + messages.length + ').');
                for (let i = 0; i < previous.messages.length; i++) owners.set(previous.messages[i].id, index);
                previous.messages = null;
                previous.status = 'error';
                previous.error = 'Duplicate of "' + item.name + '"';
            }
            for (let i = 0; i < messages.length; i++) owners.set(messages[i].id, index);
            return true;
        }

        stop() {
            if (this.session) this.session.stopRequested = true;
            if (this.batch && !this.batch.finished) this.batch.stopRequested = true;
            this.notifyNow();
        }
    }

    Controller.prototype.handlers = {
        ping: () => ({ ok: true }),
        getState: function () { return this.publicState(); },
        start: function (msg) {
            if (this.busy) throw new Error('An extraction is already running');
            this.runSingle(normalizeConfig(msg.config));
            return { ok: true };
        },
        startBatch: function (msg) {
            if (this.busy) throw new Error('An extraction is already running');
            if (!Array.isArray(msg.chats) || !msg.chats.length) throw new Error('No chats selected');
            this.runBatch(msg.chats, normalizeConfig(msg.config));
            return { ok: true };
        },
        stop: function () {
            this.stop();
            return { ok: true };
        },
        clearResults: function () {
            if (!this.busy) this.reset();
            this.notifyNow();
            return { ok: true };
        },
        listChats: () => ({ chats: TE.conversation.listSidebar() }),
        getResult: function () {
            const r = this.requireResult();
            return { conversationName: r.conversationName, messages: r.messages };
        },
        getFiles: function () {
            const r = this.requireResult();
            return { conversationName: r.conversationName, files: TE.attachments.collectFiles(r.messages) };
        },
        getImage: function (msg) {
            return { dataUrl: this.images.get(msg.id) };
        },
        getBatchSummary: function () {
            const b = this.requireBatch();
            const items = [];
            for (let i = 0; i < b.items.length; i++) {
                items.push({ name: b.items[i].name, messageCount: b.items[i].messages ? b.items[i].messages.length : 0 });
            }
            return { items: items };
        },
        getBatchItem: function (msg) {
            const item = this.requireBatch().items[msg.index];
            if (!item) throw new Error('No such batch item');
            return { name: item.name, messages: item.messages || [] };
        },
        markBatchExported: function () {
            if (this.batch) this.batch.exported = true;
            this.notifyNow();
            return { ok: true };
        },
        clearLogs: () => {
            TE.logEntries.length = 0;
            return { ok: true };
        }
    };

    if (TE.controller) {
        try { TE.controller.dispose(); } catch (_) { /* stale copy */ }
    }
    TE.controller = new Controller();
    TE.log.info('Ready on ' + location.hostname);
})(globalThis.TE = globalThis.TE || {});
