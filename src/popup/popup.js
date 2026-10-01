// Popup: a thin remote control. Extraction state and results live in the
// Teams tab's content script (followed live through a port); exports and
// downloads are built and served by the service worker.
(function (TE) {
    'use strict';

    const i18n = TE.i18n;
    const t = i18n.t;
    const setText = i18n.setText;

    const FORMATS = ['csv', 'html', 'md'];
    const LOG_LIMIT = 500;
    // Batch list tree: one node per sidebar item type, with its label key.
    const TREE_NODES = [['chat', 'batchGroupChats'], ['channel', 'batchGroupChannels']];
    const DEFAULT_SETTINGS = {
        extractFiles: true,
        zipDownloads: true,
        currentUserName: '',
        thresholdDate: '',
        formats: ['csv'],
        mode: 'single',
        locale: ''
    };
    // Settings stored one key each by versions before 4.0. The first two are
    // obsolete: only removed.
    const LEGACY_KEYS = ['maxScrollAttempts', 'maxMessages', 'extractFiles', 'zipDownloads', 'currentUserName',
        'thresholdDate', 'exportFormats', 'popupMode', 'selectedLocale'];

    const $ = (id) => document.getElementById(id);
    const ui = {
        langSelect: $('langSelect'),
        notTeams: $('notTeamsAlert'),
        tabSingle: $('tabSingle'),
        tabBatch: $('tabBatch'),
        singlePanel: $('singlePanel'),
        batchPanel: $('batchPanel'),
        singleLock: $('singleLockNote'),
        batchLock: $('batchLockNote'),
        singleStatus: $('singleStatus'),
        singleDot: $('singleDot'),
        singleStatusText: $('singleStatusText'),
        singleStatusDetail: $('singleStatusDetail'),
        startBtn: $('startBtn'),
        haltBtn: $('haltBtn'),
        singleProgress: $('singleProgress'),
        formatButtons: document.querySelectorAll('.format'),
        saveBtn: $('saveBtn'),
        saveFilesBtn: $('saveFilesBtn'),
        resetBtn: $('resetBtn'),
        filesSection: $('filesSection'),
        filesProgress: $('filesProgress'),
        filesText: $('filesText'),
        filesErrors: $('filesErrors'),
        stopFilesBtn: $('stopFilesBtn'),
        batchStatus: $('batchStatus'),
        batchDot: $('batchDot'),
        batchStatusText: $('batchStatusText'),
        batchDownloadBtn: $('batchDownloadBtn'),
        batchSetup: $('batchSetup'),
        batchCount: $('batchCount'),
        refreshChatsBtn: $('refreshChatsBtn'),
        chatList: $('chatList'),
        startBatchBtn: $('startBatchBtn'),
        batchProgressBox: $('batchProgressBox'),
        batchProgressText: $('batchProgressText'),
        batchProgress: $('batchProgress'),
        batchCurrent: $('batchCurrent'),
        stopBatchBtn: $('stopBatchBtn'),
        toggleSettings: $('toggleSettings'),
        toggleLog: $('toggleLog'),
        settingsSection: $('settingsSection'),
        withFiles: $('withFiles'),
        zipBundle: $('zipBundle'),
        thresholdDate: $('thresholdDate'),
        currentUserName: $('currentUserName'),
        logPanel: $('logPanel'),
        logView: $('logView'),
        copyConsoleBtn: $('copyConsoleBtn'),
        clearConsoleBtn: $('clearConsoleBtn'),
        error: $('errorMessage')
    };

    const app = {
        settings: null,
        tabId: null,
        onTeams: false,
        content: null,        // latest state published by the Teams tab
        job: null,            // attachments download run by the service worker
        filesVisible: false,
        filesHideTimer: 0,
        chats: [],
        selected: new Set(),
        collapsed: new Set(), // tree node types folded in the batch list
        chatsLoaded: false,
        exporting: false,     // building a transcript export
        exportingBatch: false,
        logs: [],
        errorTimer: 0
    };

    // --- settings ---

    async function loadSettings() {
        const stored = await chrome.storage.local.get(null);
        const s = Object.assign({}, DEFAULT_SETTINGS, stored.settings);
        if (!stored.settings) {
            for (let i = 2; i < 6; i++) {
                if (stored[LEGACY_KEYS[i]] !== undefined) s[LEGACY_KEYS[i]] = stored[LEGACY_KEYS[i]];
            }
            if (Array.isArray(stored.exportFormats)) s.formats = stored.exportFormats;
            if (stored.popupMode) s.mode = stored.popupMode;
            if (stored.selectedLocale) s.locale = stored.selectedLocale;
            await chrome.storage.local.set({ settings: s });
            await chrome.storage.local.remove(LEGACY_KEYS);
        }
        const formats = [];
        for (let i = 0; i < FORMATS.length; i++) {
            if (Array.isArray(s.formats) && s.formats.indexOf(FORMATS[i]) !== -1) formats.push(FORMATS[i]);
        }
        s.formats = formats;
        return s;
    }

    function saveSettings() {
        chrome.storage.local.set({ settings: app.settings });
    }

    function bindSettings() {
        const s = app.settings;
        ui.withFiles.checked = s.extractFiles;
        ui.zipBundle.checked = s.zipDownloads;
        ui.thresholdDate.value = s.thresholdDate;
        ui.currentUserName.value = s.currentUserName;

        const onChange = () => {
            s.extractFiles = ui.withFiles.checked;
            s.zipDownloads = ui.zipBundle.checked;
            s.thresholdDate = ui.thresholdDate.value;
            s.currentUserName = ui.currentUserName.value.trim();
            saveSettings();
            render();
        };
        const inputs = [ui.withFiles, ui.zipBundle, ui.thresholdDate, ui.currentUserName];
        for (let i = 0; i < inputs.length; i++) inputs[i].addEventListener('change', onChange);
    }

    function extractionConfig() {
        const s = app.settings;
        return {
            extractFiles: s.extractFiles,
            currentUserName: s.currentUserName,
            thresholdDate: s.thresholdDate,
            unknownAuthor: t('unknownAuthor')
        };
    }

    // UI strings the exporters need, in the popup's language.
    function exportLabels() {
        return {
            locale: i18n.locale.replace('_', '-'),
            title: t('htmlTranscriptTitle'),
            quotedFrom: t('htmlQuotedFrom'),
            imageNote: t('htmlImageNote'),
            unknownAuthor: t('unknownAuthor'),
            csv: {
                timestamp: t('csvTimestamp'),
                author: t('csvAuthor'),
                message: t('csvMessage'),
                reactions: t('csvReactions'),
                quotedAuthor: t('csvQuotedAuthor'),
                quotedTime: t('csvQuotedTime'),
                quotedMessage: t('csvQuotedMessage'),
                subject: t('csvSubject'),
                threadId: t('csvThreadId'),
                replyTo: t('csvReplyTo')
            }
        };
    }

    // --- messaging ---

    async function toWorker(message) {
        const res = await chrome.runtime.sendMessage(message);
        if (res && res.error) {
            const e = new Error(res.error);
            e.code = res.code;
            throw e;
        }
        return res || {};
    }

    async function toTab(message) {
        const res = await chrome.tabs.sendMessage(app.tabId, message);
        if (res && res.error) throw new Error(res.error);
        return res || {};
    }

    async function ping() {
        try {
            const res = await chrome.tabs.sendMessage(app.tabId, { type: 'ping' });
            return !!(res && res.ok);
        } catch (_) {
            return false;
        }
    }

    // Reach the tab's content scripts, injecting them if the tab was opened
    // before the extension was installed or reloaded.
    async function connectToTab() {
        if (!(await ping())) {
            try {
                const files = chrome.runtime.getManifest().content_scripts[0].js;
                await chrome.scripting.executeScript({ target: { tabId: app.tabId }, files: files });
            } catch (e) {
                note('warn', 'Could not inject the content scripts: ' + e.message);
            }
            if (!(await ping())) {
                showError(t('errorStartExtraction'));
                return;
            }
        }
        const port = chrome.tabs.connect(app.tabId, { name: 'popup' });
        port.onMessage.addListener(onPortMessage);
        port.onDisconnect.addListener(() => {
            app.content = null;
            render();
        });
        if (app.settings.mode === 'batch') loadChats();
    }

    function onPortMessage(msg) {
        if (msg.type === 'state') {
            onState(msg.state);
        } else if (msg.type === 'logs') {
            app.logs = msg.entries.slice(-LOG_LIMIT);
            if (!ui.logPanel.hidden) renderConsole();
        } else if (msg.type === 'log') {
            addLog(msg.entry);
        }
    }

    function onState(state) {
        const previous = app.content;
        app.content = state;
        // A batch that finishes while the popup is open is exported at once.
        const b = state.batch;
        if (previous && previous.busy === 'batch' && state.busy !== 'batch' && b && b.finished && b.exportable && !b.exported) {
            exportBatch();
        }
        render();
    }

    function onJob(job) {
        const wasRunning = !!app.job && app.job.status === 'running';
        app.job = job;
        clearTimeout(app.filesHideTimer);
        if (job && job.status === 'running') {
            app.filesVisible = true;
        } else if (wasRunning && job && !job.errors.length) {
            // Success: show the result briefly. Errors stay until the next run.
            app.filesHideTimer = setTimeout(() => {
                app.filesVisible = false;
                render();
            }, 10000);
        }
        render();
    }

    // Navigate a hidden iframe to a download served by the service worker;
    // the download manager takes over and the popup may close.
    function startDownload(url) {
        const frame = document.createElement('iframe');
        frame.hidden = true;
        frame.src = url;
        document.body.appendChild(frame);
        setTimeout(() => frame.remove(), 60000);
    }

    // --- actions ---

    async function startExtraction() {
        hideError();
        try {
            await toTab({ type: 'start', config: extractionConfig() });
        } catch (e) {
            note('error', 'Start failed: ' + e.message);
            showError(t('errorStartExtraction'));
        }
    }

    function stopExtraction() {
        toTab({ type: 'stop' }).catch(() => {});
    }

    async function downloadTranscript() {
        app.exporting = true;
        render();
        try {
            const res = await toWorker({
                type: 'export:transcript',
                tabId: app.tabId,
                formats: app.settings.formats,
                includeFiles: app.settings.extractFiles,
                labels: exportLabels()
            });
            startDownload(res.url);
        } catch (e) {
            showError(e.message);
        }
        app.exporting = false;
        render();
    }

    async function downloadFiles() {
        hideError();
        try {
            const res = await toWorker({ type: 'files:start', tabId: app.tabId, zip: app.settings.zipDownloads });
            if (res.url) startDownload(res.url);
        } catch (e) {
            showError(e.code === 'busy' ? t('errorZipBusy') : t('errorFileDownloadStart') + ': ' + e.message);
        }
    }

    async function exportBatch() {
        if (app.exportingBatch) return;
        app.exportingBatch = true;
        render();
        try {
            const res = await toWorker({
                type: 'export:batch',
                tabId: app.tabId,
                formats: app.settings.formats.length ? app.settings.formats : ['csv'],
                includeFiles: app.settings.extractFiles,
                labels: exportLabels()
            });
            if (res.url) startDownload(res.url);
        } catch (e) {
            showError(t('errorZipStartFailed') + ': ' + e.message);
        }
        app.exportingBatch = false;
        render();
    }

    async function startBatch() {
        hideError();
        const chats = [];
        for (let i = 0; i < app.chats.length; i++) {
            if (app.selected.has(app.chats[i].id)) chats.push({ id: app.chats[i].id, name: app.chats[i].name });
        }
        try {
            await toTab({ type: 'startBatch', chats: chats, config: extractionConfig() });
        } catch (e) {
            note('error', 'Batch start failed: ' + e.message);
            showError(t('errorStartExtraction'));
        }
    }

    async function loadChats() {
        app.chatsLoaded = true;
        renderChatPlaceholder('batchLoadingChats');
        try {
            app.chats = (await toTab({ type: 'listChats' })).chats || [];
        } catch (_) {
            app.chats = [];
        }
        const listed = new Set();
        for (let i = 0; i < app.chats.length; i++) listed.add(app.chats[i].id);
        app.selected.forEach((id) => {
            if (!listed.has(id)) app.selected.delete(id);
        });
        renderChatList();
        render();
    }

    function setMode(mode) {
        app.settings.mode = mode === 'batch' ? 'batch' : 'single';
        saveSettings();
        render();
        const busy = app.content && app.content.busy;
        if (app.settings.mode === 'batch' && !app.chatsLoaded && app.content && busy !== 'batch') loadChats();
    }

    function toggleFormat(format) {
        const formats = [];
        const on = app.settings.formats.indexOf(format) === -1;
        for (let i = 0; i < FORMATS.length; i++) {
            const f = FORMATS[i];
            if (f === format ? on : app.settings.formats.indexOf(f) !== -1) formats.push(f);
        }
        app.settings.formats = formats;
        saveSettings();
        render();
    }

    // --- rendering ---

    function setRaw(el, text) {
        el.removeAttribute('data-i18n');
        el.removeAttribute('data-i18n-args');
        el.textContent = text;
    }

    function render() {
        const c = app.content;
        const busy = c ? c.busy : null;
        const batchMode = app.settings.mode === 'batch';

        ui.notTeams.hidden = app.onTeams;
        ui.tabSingle.classList.toggle('active', !batchMode);
        ui.tabBatch.classList.toggle('active', batchMode);
        ui.tabSingle.setAttribute('aria-selected', String(!batchMode));
        ui.tabBatch.setAttribute('aria-selected', String(batchMode));
        ui.tabSingle.classList.toggle('running', busy === 'single');
        ui.tabBatch.classList.toggle('running', busy === 'batch');
        ui.singlePanel.hidden = batchMode;
        ui.batchPanel.hidden = !batchMode;
        // Both modes drive the same tab: only one can run at a time.
        ui.singleLock.hidden = busy !== 'batch';
        ui.batchLock.hidden = busy !== 'single';

        for (let i = 0; i < ui.formatButtons.length; i++) {
            const button = ui.formatButtons[i];
            button.setAttribute('aria-pressed', String(app.settings.formats.indexOf(button.dataset.format) !== -1));
        }

        renderSingle(c, busy);
        renderFiles();
        renderBatch(c, busy);
    }

    function downloadLabel(formats) {
        if (!formats.length) return 'btnSelectFormat';
        if (formats.length > 1) return 'btnDownloadZip';
        return formats[0] === 'html' ? 'btnDownloadHTML' : formats[0] === 'md' ? 'btnDownloadMD' : 'btnDownloadCSV';
    }

    function renderSingle(c, busy) {
        const running = busy === 'single';
        const result = c ? c.result : null;

        let key = null, args, dot = '', detailKey = null, detailArgs, detailText = '';
        if (running) {
            dot = 'running';
            if (c.stopping) key = 'statusFinalizing';
            else if (c.progress && c.progress.messages) {
                key = 'statusExtractingProgress';
                args = [c.progress.messages, c.progress.files];
            } else key = 'statusStarting';
        } else if (c && c.error) {
            dot = 'error';
            key = 'statusFailed';
            detailText = c.error;
        } else if (result) {
            key = 'statusComplete';
            args = [result.messageCount, result.fileCount];
            detailKey = 'statusTotalTime';
            detailArgs = [TE.util.formatDuration(Math.round(result.durationMs / 1000))];
        }
        ui.singleStatus.hidden = !key;
        if (key) {
            ui.singleDot.className = 'dot ' + dot;
            setText(ui.singleStatusText, key, args);
        }
        ui.singleStatusDetail.hidden = !detailKey && !detailText;
        if (detailKey) setText(ui.singleStatusDetail, detailKey, detailArgs);
        else setRaw(ui.singleStatusDetail, detailText);

        ui.startBtn.hidden = running;
        ui.startBtn.disabled = !c || !!busy;
        setText(ui.startBtn, result ? 'btnStartNewExtraction' : 'btnStartExtraction');
        ui.haltBtn.hidden = !running;
        ui.haltBtn.disabled = running && c.stopping;
        setText(ui.haltBtn, running && c.stopping ? 'btnFinalizing' : 'btnStopExtraction');
        ui.singleProgress.hidden = !running;

        const formats = app.settings.formats;
        setText(ui.saveBtn, downloadLabel(formats));
        ui.saveBtn.disabled = !result || !result.messageCount || !formats.length || !!busy || app.exporting;

        const job = app.job;
        const jobRunning = !!job && job.status === 'running';
        ui.saveFilesBtn.hidden = !app.settings.extractFiles;
        if (jobRunning) setText(ui.saveFilesBtn, 'btnDownloadingFiles', [job.done, job.total]);
        else setText(ui.saveFilesBtn, 'btnDownloadFiles', [running && c.progress ? c.progress.files : result ? result.fileCount : 0]);
        ui.saveFilesBtn.disabled = !result || !result.fileCount || jobRunning || !!busy;

        ui.resetBtn.hidden = !result || !!busy;
    }

    function renderFiles() {
        const job = app.job;
        ui.filesSection.hidden = !job || !app.filesVisible;
        if (ui.filesSection.hidden) return;

        const running = job.status === 'running';
        const pct = !running ? 100 : job.total ? Math.round(job.done * 100 / job.total) : 0;
        ui.filesProgress.style.width = pct + '%';

        if (running) {
            if (job.kind === 'zip' && !job.done && !job.current) {
                setText(ui.filesText, 'fileProgressPreparingZip', [job.total]);
            } else {
                setText(ui.filesText, 'fileProgressDownloading', [job.done, job.total]);
                if (job.current) ui.filesText.textContent += ': ' + job.current;
            }
        } else if (job.status === 'done' && !job.errors.length) {
            setText(ui.filesText, 'fileProgressAllDone', [job.done]);
        } else {
            setText(ui.filesText, 'fileProgressDone', [job.done, job.total]);
        }

        if (job.errors.length) {
            let names = '';
            let pathTooLong = false;
            for (let i = 0; i < job.errors.length; i++) {
                names += (i ? ', ' : '') + job.errors[i].name;
                if (/name_too_long/i.test(job.errors[i].error)) pathTooLong = true;
            }
            setText(ui.filesErrors, 'fileErrorsList', [job.errors.length, names]);
            if (pathTooLong) ui.filesErrors.textContent += '\n' + t('filePathTooLongHint');
        } else {
            setRaw(ui.filesErrors, '');
        }
        ui.stopFilesBtn.hidden = !running;
    }

    function renderBatch(c, busy) {
        const b = c ? c.batch : null;
        const running = busy === 'batch';

        ui.batchSetup.hidden = running;
        ui.batchProgressBox.hidden = !running;
        if (running && b) {
            const position = Math.min(b.index + 1, b.total);
            ui.batchProgressText.textContent = position + ' / ' + b.total;
            ui.batchProgress.style.width = (b.total ? position * 100 / b.total : 0) + '%';
            if (b.current) setText(ui.batchCurrent, b.currentStatus === 'error' ? 'batchChatError' : 'batchChatExtracting', [b.current]);
            else setRaw(ui.batchCurrent, '');
            ui.stopBatchBtn.disabled = b.stopRequested;
            setText(ui.stopBatchBtn, b.stopRequested ? 'btnFinalizing' : 'btnStopBatch');
        }

        const finished = !running && !!b && b.finished;
        ui.batchStatus.hidden = !finished && !app.exportingBatch;
        ui.batchDownloadBtn.hidden = true;
        if (app.exportingBatch) {
            ui.batchDot.className = 'dot running';
            setText(ui.batchStatusText, 'batchBuildingZip');
        } else if (finished && !b.exportable) {
            ui.batchDot.className = 'dot error';
            setText(ui.batchStatusText, 'batchStatusNothing');
        } else if (finished) {
            ui.batchDot.className = 'dot';
            setText(ui.batchStatusText, b.exported ? 'batchStatusComplete' : 'batchStatusReady', [b.exportable]);
            ui.batchDownloadBtn.hidden = false;
        }

        const n = app.selected.size;
        ui.batchCount.textContent = app.chats.length ? n + ' / ' + app.chats.length : '';
        setText(ui.startBatchBtn, n === 0 ? 'btnSelectChats' : n === 1 ? 'btnStartBatchOne' : 'btnStartBatch', [n]);
        ui.startBatchBtn.disabled = !n || !app.settings.formats.length || !c || !!busy || app.exportingBatch;
        ui.refreshChatsBtn.disabled = !c || running;
    }

    function renderChatPlaceholder(key) {
        ui.chatList.textContent = '';
        const div = document.createElement('div');
        div.className = 'empty';
        setText(div, key);
        ui.chatList.appendChild(div);
    }

    function checkbox(checked, onChange) {
        const input = document.createElement('input');
        input.type = 'checkbox';
        input.checked = checked;
        input.addEventListener('change', onChange);
        return input;
    }

    // Select or clear every chat of a tree node.
    function selectNode(node, on) {
        for (let i = 0; i < node.chats.length; i++) {
            node.boxes[i].checked = on;
            if (on) app.selected.add(node.chats[i].id);
            else app.selected.delete(node.chats[i].id);
        }
    }

    // Tree node of one sidebar item type: a collapsible header whose checkbox
    // selects all its chats, then one row per chat.
    function renderNode(node, labelKey, sync) {
        const wrap = document.createElement('div');
        wrap.className = 'node';
        const head = document.createElement('div');
        head.className = 'node-head';
        const items = document.createElement('div');
        items.className = 'node-items';
        items.setAttribute('role', 'group');
        items.hidden = app.collapsed.has(node.type);

        const caret = document.createElement('button');
        caret.className = 'caret';
        caret.setAttribute('aria-expanded', String(!items.hidden));
        caret.addEventListener('click', () => {
            items.hidden = !items.hidden;
            caret.setAttribute('aria-expanded', String(!items.hidden));
            if (items.hidden) app.collapsed.add(node.type);
            else app.collapsed.delete(node.type);
        });

        const label = document.createElement('label');
        node.box = checkbox(false, () => {
            selectNode(node, node.box.checked);
            sync();
            render();
        });
        const text = document.createElement('span');
        setText(text, labelKey);
        node.count = document.createElement('span');
        node.count.className = 'muted';
        label.append(node.box, text, node.count);
        head.append(caret, label);

        for (let i = 0; i < node.chats.length; i++) {
            const chat = node.chats[i];
            const row = document.createElement('label');
            const box = checkbox(app.selected.has(chat.id), () => {
                if (box.checked) app.selected.add(chat.id);
                else app.selected.delete(chat.id);
                sync();
                render();
            });
            node.boxes.push(box);
            const name = document.createElement('span');
            name.textContent = chat.name || chat.id.slice(0, 30);
            name.title = chat.id;
            row.append(box, name);
            items.appendChild(row);
        }
        wrap.append(head, items);
        return wrap;
    }

    function renderChatList() {
        if (!app.chats.length) {
            renderChatPlaceholder('batchNoChats');
            return;
        }
        ui.chatList.textContent = '';
        const nodes = [];
        const all = document.createElement('label');
        all.className = 'all';
        const allBox = checkbox(false, () => {
            for (let i = 0; i < nodes.length; i++) selectNode(nodes[i], allBox.checked);
            sync();
            render();
        });
        const allText = document.createElement('span');
        setText(allText, 'batchSelectAll');
        all.append(allBox, allText);
        ui.chatList.appendChild(all);

        // Tri-state checkboxes of the nodes and of "select all".
        const sync = () => {
            let total = 0, selected = 0;
            for (let i = 0; i < nodes.length; i++) {
                const node = nodes[i];
                let n = 0;
                for (let j = 0; j < node.chats.length; j++) {
                    if (app.selected.has(node.chats[j].id)) n++;
                }
                node.box.checked = n === node.chats.length;
                node.box.indeterminate = n > 0 && n < node.chats.length;
                node.count.textContent = n + ' / ' + node.chats.length;
                total += node.chats.length;
                selected += n;
            }
            allBox.checked = selected === total;
            allBox.indeterminate = selected > 0 && selected < total;
        };
        for (let g = 0; g < TREE_NODES.length; g++) {
            const node = { type: TREE_NODES[g][0], chats: [], boxes: [], box: null, count: null };
            for (let i = 0; i < app.chats.length; i++) {
                if (app.chats[i].type === node.type) node.chats.push(app.chats[i]);
            }
            if (!node.chats.length) continue;
            nodes.push(node);
            ui.chatList.appendChild(renderNode(node, TREE_NODES[g][1], sync));
        }
        sync();
    }

    // --- console & errors ---

    function logLine(entry) {
        const line = document.createElement('div');
        line.className = entry.level;
        const time = document.createElement('time');
        time.textContent = new Date(entry.time).toLocaleTimeString();
        line.append(time, entry.text);
        return line;
    }

    function renderConsole() {
        const fragment = document.createDocumentFragment();
        for (let i = 0; i < app.logs.length; i++) fragment.appendChild(logLine(app.logs[i]));
        ui.logView.textContent = '';
        ui.logView.appendChild(fragment);
        ui.logView.scrollTop = ui.logView.scrollHeight;
    }

    function addLog(entry) {
        app.logs.push(entry);
        if (app.logs.length > LOG_LIMIT) app.logs.shift();
        if (ui.logPanel.hidden) return;
        const out = ui.logView;
        out.appendChild(logLine(entry));
        while (out.childElementCount > LOG_LIMIT) out.firstElementChild.remove();
        out.scrollTop = out.scrollHeight;
    }

    // Popup-side entry in the console.
    function note(level, text) {
        addLog({ time: Date.now(), level: level, text: '[popup] ' + text });
    }

    function showError(message) {
        note('error', message);
        ui.error.textContent = message;
        ui.error.hidden = false;
        clearTimeout(app.errorTimer);
        app.errorTimer = setTimeout(hideError, 6000);
    }

    function hideError() {
        ui.error.hidden = true;
    }

    // --- wiring ---

    function wireEvents() {
        ui.tabSingle.addEventListener('click', () => setMode('single'));
        ui.tabBatch.addEventListener('click', () => setMode('batch'));
        ui.startBtn.addEventListener('click', startExtraction);
        ui.haltBtn.addEventListener('click', stopExtraction);
        ui.saveBtn.addEventListener('click', downloadTranscript);
        ui.saveFilesBtn.addEventListener('click', downloadFiles);
        ui.stopFilesBtn.addEventListener('click', () => toWorker({ type: 'files:stop' }).catch(() => {}));
        ui.resetBtn.addEventListener('click', () => toTab({ type: 'clearResults' }).catch(() => {}));
        ui.refreshChatsBtn.addEventListener('click', loadChats);
        ui.startBatchBtn.addEventListener('click', startBatch);
        ui.stopBatchBtn.addEventListener('click', stopExtraction);
        ui.batchDownloadBtn.addEventListener('click', exportBatch);
        for (let i = 0; i < ui.formatButtons.length; i++) {
            const button = ui.formatButtons[i];
            button.addEventListener('click', () => toggleFormat(button.dataset.format));
        }

        ui.toggleSettings.addEventListener('click', () => {
            ui.settingsSection.hidden = !ui.settingsSection.hidden;
            ui.toggleSettings.setAttribute('aria-expanded', String(!ui.settingsSection.hidden));
            ui.toggleSettings.classList.toggle('active', !ui.settingsSection.hidden);
        });
        ui.toggleLog.addEventListener('click', () => {
            ui.logPanel.hidden = !ui.logPanel.hidden;
            setText(ui.toggleLog, ui.logPanel.hidden ? 'toggleShowConsole' : 'toggleHideConsole');
            if (!ui.logPanel.hidden) renderConsole();
        });
        ui.copyConsoleBtn.addEventListener('click', async () => {
            let text = '';
            for (let i = 0; i < app.logs.length; i++) {
                const e = app.logs[i];
                text += '[' + new Date(e.time).toLocaleTimeString() + '] [' + e.level + '] ' + e.text + '\n';
            }
            try {
                await navigator.clipboard.writeText(text);
                setText(ui.copyConsoleBtn, 'btnCopied');
                setTimeout(() => setText(ui.copyConsoleBtn, 'btnCopy'), 1200);
            } catch (e) {
                note('error', 'Copy failed: ' + e.message);
            }
        });
        ui.clearConsoleBtn.addEventListener('click', () => {
            app.logs = [];
            ui.logView.textContent = '';
            if (app.content) toTab({ type: 'clearLogs' }).catch(() => {});
        });

        ui.langSelect.addEventListener('change', async () => {
            app.settings.locale = ui.langSelect.value;
            saveSettings();
            await i18n.load(app.settings.locale);
            i18n.apply(document);
            render();
        });

        // Download progress broadcast by the service worker.
        chrome.runtime.onMessage.addListener((msg) => {
            if (msg && msg.type === 'dl:progress') onJob(msg.job);
        });
    }

    async function init() {
        app.settings = await loadSettings();
        await i18n.load(app.settings.locale || i18n.fromBrowser(chrome.i18n.getUILanguage()));
        for (let i = 0; i < i18n.LOCALES.length; i++) {
            ui.langSelect.add(new Option(i18n.LOCALES[i][1], i18n.LOCALES[i][0]));
        }
        ui.langSelect.value = i18n.locale;
        bindSettings();
        wireEvents();
        i18n.apply(document);

        const tabs = await chrome.tabs.query({ active: true, currentWindow: true });
        app.tabId = tabs.length ? tabs[0].id : null;
        app.onTeams = tabs.length > 0 && TE.util.isTeamsUrl(tabs[0].url || '');
        try {
            const files = await toWorker({ type: 'files:state' });
            app.job = files.job || null;
            app.filesVisible = !!app.job && app.job.status === 'running';
        } catch (_) { /* worker waking up: progress messages will follow */ }

        render();
        document.body.style.visibility = 'visible';
        if (app.onTeams) await connectToTab();
    }

    init();
})(globalThis.TE);
