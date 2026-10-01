// Transcript exporters: CSV, HTML and Markdown.
//
// Pure string builders, run in the service worker. Input:
//   conversation: { title, messages }   messages as produced by the content
//                                       script (see content/richtext.js for
//                                       the body model), already ordered.
//   opts: { includeFiles, labels, images }
//     labels  — UI strings resolved by the popup in the user's language
//     images  — Map<imageId, dataUrl> for captured inline images (HTML only)
//
// Channel conversations (messages carrying a threadId) are grouped by thread:
// the root post under its subject, followed by its replies.
(function (TE) {
    'use strict';

    const util = TE.util;

    // --- shared ---

    function isChannel(messages) {
        for (let i = 0; i < messages.length; i++) {
            if (messages[i].threadId) return true;
        }
        return false;
    }

    function authorCount(messages) {
        const authors = new Set();
        for (let i = 0; i < messages.length; i++) authors.add(messages[i].author);
        return authors.size;
    }

    function formatReactions(reactions) {
        if (!reactions) return '';
        let out = '';
        for (let i = 0; i < reactions.length; i++) {
            const { emoji, count } = reactions[i];
            out += (i > 0 ? '; ' : '') + emoji + (count > 1 ? '×' + count : '');
        }
        return out;
    }

    function dateFormatter(locale, options) {
        try {
            return new Intl.DateTimeFormat(locale, options);
        } catch (_) {
            return new Intl.DateTimeFormat('en', options);
        }
    }

    function formatWith(formatter, iso) {
        if (!iso) return '';
        const d = new Date(iso);
        return isNaN(d.getTime()) ? String(iso) : formatter.format(d);
    }

    // --- CSV ---

    function csvField(value) {
        return '"' + String(value || '').replace(/\r\n?/g, '\n').replace(/"/g, '""') + '"';
    }

    function buildCsv(conversation, opts) {
        const messages = conversation.messages;
        const channel = isChannel(messages);
        const csv = opts.labels.csv;

        const header = [csv.timestamp, csv.author, csv.message, csv.reactions,
            csv.quotedAuthor, csv.quotedTime, csv.quotedMessage];
        if (channel) header.push(csv.subject, csv.threadId, csv.replyTo);
        if (opts.includeFiles) header.push('Attachments', 'Attachment URLs');

        const rows = [];
        let line = '';
        for (let i = 0; i < header.length; i++) line += (i ? ',' : '') + csvField(header[i]);
        rows.push(line);

        for (let i = 0; i < messages.length; i++) {
            const m = messages[i];
            let quotedAuthor = '', quotedTime = '', quotedText = '';
            for (let q = 0; q < m.quotes.length; q++) {
                const sep = q ? '\n' : '';
                quotedAuthor += sep + (m.quotes[q].author || '');
                quotedTime += sep + util.toIso(m.quotes[q].date);
                quotedText += sep + (m.quotes[q].text || '');
            }
            line = csvField(util.toIso(m.timestamp)) + ',' + csvField(m.author || opts.labels.unknownAuthor) + ','
                + csvField(util.blocksToText(m.blocks)) + ',' + csvField(formatReactions(m.reactions)) + ','
                + csvField(quotedAuthor) + ',' + csvField(quotedTime) + ',' + csvField(quotedText);
            if (channel) {
                line += ',' + csvField(m.subject) + ',' + csvField(m.threadId) + ',' + csvField(m.replyToId);
            }
            if (opts.includeFiles) {
                let names = '', urls = '';
                for (let a = 0; a < m.attachments.length; a++) {
                    const att = m.attachments[a];
                    const sep = a ? '; ' : '';
                    names += sep + att.name;
                    urls += sep + (att.url || (att.imageId ? '[embedded image]' : ''));
                }
                line += ',' + csvField(names) + ',' + csvField(urls);
            }
            rows.push(line);
        }
        // The BOM makes Excel read the file as UTF-8 instead of the ANSI code page.
        return '﻿' + rows.join('\r\n');
    }

    // --- HTML ---

    const HTML_ESCAPES = { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' };

    function escapeHtml(s) {
        return String(s == null ? '' : s).replace(/[&<>"']/g, (c) => HTML_ESCAPES[c]);
    }

    function runsToHtml(runs) {
        let out = '';
        for (let i = 0; i < runs.length; i++) {
            const r = runs[i];
            let h = escapeHtml(r.text);
            if (r.mention) h = '<span class="mention">' + h + '</span>';
            if (r.b) h = '<strong>' + h + '</strong>';
            if (r.i) h = '<em>' + h + '</em>';
            if (r.u) h = '<u>' + h + '</u>';
            if (r.s) h = '<s>' + h + '</s>';
            if (r.href) h = '<a href="' + escapeHtml(r.href) + '" target="_blank" rel="noopener noreferrer">' + h + '</a>';
            out += h;
        }
        return out;
    }

    function blocksToHtml(blocks, quotes, ctx) {
        let out = '';
        for (let i = 0; i < blocks.length; i++) {
            const block = blocks[i];
            switch (block.kind) {
                case 'text': {
                    out += '<p>';
                    for (let l = 0; l < block.lines.length; l++) {
                        if (l) out += '<br>';
                        out += runsToHtml(block.lines[l]);
                    }
                    out += '</p>';
                    break;
                }
                case 'code':
                    out += '<pre class="code"><code>' + escapeHtml(block.text) + '</code></pre>';
                    break;
                case 'quote':
                    out += '<blockquote class="inline-quote">' + blocksToHtml(block.blocks, quotes, ctx) + '</blockquote>';
                    break;
                case 'list': {
                    const tag = block.ordered ? 'ol' : 'ul';
                    out += '<' + tag + '>';
                    for (let j = 0; j < block.items.length; j++) out += '<li>' + blocksToHtml(block.items[j], quotes, ctx) + '</li>';
                    out += '</' + tag + '>';
                    break;
                }
                case 'reply': {
                    const q = quotes[block.index];
                    if (q) out += quoteCardHtml(q, ctx);
                    break;
                }
            }
        }
        return out;
    }

    function quoteCardHtml(q, ctx) {
        let head = '';
        if (q.author) {
            head += '<span class="quote-label">' + escapeHtml(ctx.labels.quotedFrom) + '</span> '
                + '<span class="author">' + escapeHtml(q.author) + '</span>';
        }
        const iso = util.toIso(q.date);
        if (iso) head += '<time datetime="' + escapeHtml(iso) + '">' + escapeHtml(formatWith(ctx.dateTime, iso)) + '</time>';
        return '<blockquote class="quote">'
            + (head ? '<header>' + head + '</header>' : '')
            + '<div class="quote-body">' + escapeHtml(q.text).replace(/\n/g, '<br>') + '</div>'
            + '</blockquote>';
    }

    function messageHtml(m, ctx) {
        let cls = 'msg';
        if (m.isReply) cls += ' reply';
        if (m.isMine) cls += ' mine';
        if (m.isSystem) cls += ' system';

        const body = blocksToHtml(m.blocks, m.quotes, ctx) || '<p class="empty">(no text)</p>';
        let html = '<article class="' + cls + '">';
        if (!m.isSystem) {
            const iso = util.toIso(m.timestamp);
            html += '<header><span class="author">' + escapeHtml(m.author || ctx.labels.unknownAuthor) + '</span>'
                + '<time' + (iso ? ' datetime="' + escapeHtml(iso) + '"' : '') + '>'
                + escapeHtml(formatWith(ctx.dateTime, m.timestamp)) + '</time></header>';
        }
        html += '<div class="body" dir="auto">' + body + '</div>';

        if (m.reactions.length) {
            html += '<div class="reactions">';
            for (let i = 0; i < m.reactions.length; i++) {
                const r = m.reactions[i];
                html += '<span class="reaction">' + escapeHtml(r.emoji)
                    + (r.count > 1 ? '<span class="count">' + r.count + '</span>' : '') + '</span>';
            }
            html += '</div>';
        }

        if (ctx.includeFiles && m.attachments.length) {
            let images = '', files = '';
            for (let i = 0; i < m.attachments.length; i++) {
                const a = m.attachments[i];
                const dataUrl = a.imageId && ctx.images ? ctx.images.get(a.imageId) : null;
                // Captured URLs are data:image/…;base64 — the quote check keeps a
                // malformed one from breaking out of the attribute.
                if (dataUrl && dataUrl.startsWith('data:image/') && dataUrl.indexOf('"') === -1) {
                    images += '<figure><img src="' + dataUrl + '" alt="' + escapeHtml(a.name) + '" loading="lazy">'
                        + '<figcaption>' + escapeHtml(a.name) + '</figcaption></figure>';
                    continue;
                }
                if (a.image) ctx.missingImage = true;
                files += a.url
                    ? '<a href="' + escapeHtml(a.url) + '" target="_blank" rel="noopener noreferrer">📎 ' + escapeHtml(a.name) + '</a>'
                    : '<span>📎 ' + escapeHtml(a.name) + '</span>';
            }
            if (images) html += '<div class="images">' + images + '</div>';
            if (files) html += '<div class="files">' + files + '</div>';
        }
        return html + '</article>';
    }

    function buildHtml(conversation, opts) {
        const messages = conversation.messages;
        const labels = opts.labels;
        const ctx = {
            labels: labels,
            includeFiles: opts.includeFiles,
            images: opts.images,
            missingImage: false,
            dateTime: dateFormatter(labels.locale, { dateStyle: 'medium', timeStyle: 'short' }),
            day: dateFormatter(labels.locale, { dateStyle: 'full' })
        };

        let main = '';
        if (isChannel(messages)) {
            let threadId = null;
            for (let i = 0; i < messages.length; i++) {
                const m = messages[i];
                if (m.threadId !== threadId) {
                    if (threadId !== null) main += '</section>';
                    threadId = m.threadId;
                    const subject = !m.isReply && m.subject ? m.subject : '(no subject)';
                    main += '<section class="thread"><h2>' + escapeHtml(subject) + '</h2>';
                }
                main += messageHtml(m, ctx);
            }
            if (threadId !== null) main += '</section>';
        } else {
            let day = '';
            for (let i = 0; i < messages.length; i++) {
                const m = messages[i];
                const d = formatWith(ctx.day, m.timestamp);
                if (d && d !== day) {
                    day = d;
                    main += '<h2 class="day">' + escapeHtml(d) + '</h2>';
                }
                main += messageHtml(m, ctx);
            }
        }

        const title = conversation.title || labels.title;
        const note = ctx.missingImage ? '<p class="note">' + escapeHtml(labels.imageNote) + '</p>' : '';
        return '<!DOCTYPE html>\n<html lang="' + escapeHtml(labels.locale) + '">\n<head>\n<meta charset="UTF-8">\n'
            + '<meta name="viewport" content="width=device-width, initial-scale=1">\n'
            + '<title>' + escapeHtml(title) + '</title>\n<style>' + HTML_CSS + '</style>\n</head>\n<body>\n'
            + '<header class="page"><h1>' + escapeHtml(title) + '</h1>'
            + '<p class="meta">' + messages.length + ' messages · ' + authorCount(messages) + ' authors · Generated '
            + escapeHtml(ctx.dateTime.format(new Date())) + '</p></header>\n'
            + note + '<main>' + main + '</main>\n</body>\n</html>\n';
    }

    const HTML_CSS = `
:root {
  color-scheme: light dark;
  --bg: #f7f7f9; --card: #fff; --text: #1b1b1f; --muted: #6b6b76; --line: #e3e3e8;
  --accent: #5b5fc7; --mine: #f1efff; --quote: #f0f0f6; --code-bg: #1e1e1e; --code-text: #d4d4d4;
}
@media (prefers-color-scheme: dark) {
  :root {
    --bg: #18181b; --card: #242428; --text: #e6e6ea; --muted: #9a9aa6; --line: #36363d;
    --accent: #9ea2ff; --mine: #2b2940; --quote: #2e2e34;
  }
}
* { box-sizing: border-box; }
body { margin: 0 auto; max-width: 860px; padding: 24px 16px; background: var(--bg); color: var(--text);
  font: 14px/1.5 -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; }
h1 { font-size: 22px; margin: 0 0 4px; }
.meta { color: var(--muted); font-size: 13px; margin: 0 0 24px; }
.note { background: #fff6d6; color: #5c4700; border: 1px solid #f0d98a; border-radius: 6px; padding: 10px 12px; }
h2.day { font-size: 14px; color: var(--accent); border-bottom: 1px solid var(--line); padding: 12px 0 6px; margin: 8px 0 12px; }
.thread { margin-bottom: 28px; }
.thread h2 { font-size: 16px; margin: 0 0 8px; }
.msg { background: var(--card); border: 1px solid var(--line); border-left: 4px solid transparent;
  border-radius: 8px; padding: 12px 16px; margin-bottom: 10px; }
.msg.mine { background: var(--mine); border-left-color: var(--accent); }
.msg.reply { margin-left: 32px; }
.msg.system { background: none; border: 0; border-bottom: 1px solid var(--line); border-radius: 0;
  color: var(--muted); font-size: 12px; font-style: italic; padding: 6px 16px; }
.msg > header { display: flex; justify-content: space-between; gap: 12px; align-items: baseline; margin-bottom: 6px; }
.author { font-weight: 600; }
time { color: var(--muted); font-size: 12px; white-space: nowrap; }
.body p { margin: 0 0 6px; overflow-wrap: anywhere; }
.body p:last-child { margin-bottom: 0; }
.body .empty { color: var(--muted); font-style: italic; }
.body ul, .body ol { margin: 4px 0 6px; padding-left: 24px; }
.mention { color: var(--accent); font-weight: 600; }
a { color: var(--accent); }
pre.code { background: var(--code-bg); color: var(--code-text); border-radius: 6px; padding: 10px 14px;
  overflow-x: auto; font: 12px/1.5 Consolas, Monaco, "Courier New", monospace; margin: 6px 0; }
.inline-quote { margin: 6px 0; padding: 4px 12px; border-left: 3px solid var(--muted); color: var(--muted); }
.quote { margin: 4px 0 10px; padding: 8px 12px; border-left: 4px solid var(--accent); background: var(--quote);
  border-radius: 0 8px 8px 0; }
.quote header { display: flex; flex-wrap: wrap; gap: 8px; align-items: baseline; margin-bottom: 4px; }
.quote-label { text-transform: uppercase; letter-spacing: .5px; font-size: 11px; font-weight: 600; color: var(--accent); }
.quote-body { font-style: italic; overflow-wrap: anywhere; }
.reactions { display: flex; flex-wrap: wrap; gap: 6px; margin-top: 8px; }
.reaction { border: 1px solid var(--line); border-radius: 12px; padding: 1px 8px; font-size: 13px; }
.reaction .count { color: var(--muted); font-size: 12px; margin-left: 4px; }
.images { display: flex; flex-wrap: wrap; gap: 10px; margin-top: 8px; }
.images figure { margin: 0; max-width: 320px; }
.images img { display: block; max-width: 100%; height: auto; border: 1px solid var(--line); border-radius: 6px; }
figcaption { font-size: 12px; color: var(--muted); overflow-wrap: anywhere; margin-top: 4px; }
.files { display: flex; flex-direction: column; gap: 4px; margin-top: 8px; }
.files a { text-decoration: none; }
.files a:hover { text-decoration: underline; }
`;

    // --- Markdown ---
    // Portable text that pastes cleanly into notes apps, wikis and AI tools.
    // Captured images are only named: base64 blobs don't belong in Markdown.

    // Wrap non-blank text in a marker, keeping surrounding spaces outside it
    // ("** bold**" would not render).
    function wrapMd(text, mark) {
        const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(text);
        return m[2] ? m[1] + mark + m[2] + mark + m[3] : text;
    }

    function runsToMd(runs) {
        let out = '';
        for (let i = 0; i < runs.length; i++) {
            const r = runs[i];
            let t = r.mention ? wrapMd('@' + r.text, '**') : r.text;
            if (r.b && !r.mention) t = wrapMd(t, '**');
            if (r.i) t = wrapMd(t, '_');
            if (r.s) t = wrapMd(t, '~~');
            if (r.u) t = '<u>' + t + '</u>';
            if (r.href) t = t.trim() === r.href ? '<' + r.href + '>' : '[' + t + '](' + r.href + ')';
            out += t;
        }
        return out;
    }

    // Append Markdown lines for `blocks`, each prefixed (for nesting in "> ").
    function blocksToMd(blocks, quotes, labels, prefix, out) {
        const blank = prefix.trimEnd();
        for (let i = 0; i < blocks.length; i++) {
            const block = blocks[i];
            if (i) out.push(blank);
            switch (block.kind) {
                case 'text':
                    for (let l = 0; l < block.lines.length; l++) out.push(prefix + runsToMd(block.lines[l]));
                    break;
                case 'code': {
                    out.push(prefix + '```');
                    const lines = block.text.split('\n');
                    for (let l = 0; l < lines.length; l++) out.push(prefix + lines[l]);
                    out.push(prefix + '```');
                    break;
                }
                case 'quote':
                    blocksToMd(block.blocks, quotes, labels, prefix + '> ', out);
                    break;
                case 'list':
                    for (let j = 0; j < block.items.length; j++) {
                        const marker = block.ordered ? (j + 1) + '. ' : '- ';
                        const indent = ' '.repeat(marker.length);
                        const item = [];
                        blocksToMd(block.items[j], quotes, labels, '', item);
                        for (let l = 0; l < item.length; l++) {
                            out.push(prefix + (l === 0 ? marker : (item[l] ? indent : '')) + item[l]);
                        }
                    }
                    break;
                case 'reply': {
                    const q = quotes[block.index];
                    if (q) quoteToMd(q, labels, prefix, out);
                    break;
                }
            }
        }
    }

    function quoteToMd(q, labels, prefix, out) {
        let head = q.author || '';
        const iso = util.toIso(q.date);
        if (iso) head += (head ? ' · ' : '') + iso;
        out.push(prefix + '> **' + labels.quotedFrom + (head ? ' — ' + head : '') + '**');
        const lines = (q.text || '').split('\n');
        for (let l = 0; l < lines.length; l++) out.push(prefix + '> ' + lines[l]);
    }

    function messageToMd(m, opts, prefix, out) {
        const blank = prefix.trimEnd();
        out.push(prefix + (m.isReply ? '**↳ ' : '**') + (m.author || opts.labels.unknownAuthor) + '** · `' + util.toIso(m.timestamp) + '`');
        if (m.blocks.length) {
            out.push(blank);
            blocksToMd(m.blocks, m.quotes, opts.labels, prefix, out);
        }
        const reactions = formatReactions(m.reactions);
        const files = opts.includeFiles ? m.attachments : [];
        if (reactions || files.length) out.push(blank);
        if (reactions) out.push(prefix + '*Reactions:* ' + reactions);
        for (let i = 0; i < files.length; i++) {
            const a = files[i];
            if (a.imageId) out.push(prefix + '- 📎 ' + a.name + ' *(embedded image)*');
            else if (a.url) out.push(prefix + '- 📎 [' + a.name + '](' + a.url + ')');
            else out.push(prefix + '- 📎 ' + a.name);
        }
    }

    function buildMarkdown(conversation, opts) {
        const messages = conversation.messages;
        const out = [
            '# ' + (conversation.title || opts.labels.title), '',
            '*Generated ' + new Date().toISOString() + ' · ' + messages.length + ' messages · '
                + authorCount(messages) + ' participants*', '', '---', ''
        ];

        if (!isChannel(messages)) {
            for (let i = 0; i < messages.length; i++) {
                messageToMd(messages[i], opts, '', out);
                out.push('', '---', '');
            }
            return out.join('\n');
        }

        let threadId = null;
        for (let i = 0; i < messages.length; i++) {
            const m = messages[i];
            if (m.threadId !== threadId) {
                if (threadId !== null) out.push('---', '');
                threadId = m.threadId;
                out.push('## ' + (!m.isReply && m.subject ? m.subject : '(no subject)'), '');
            }
            messageToMd(m, opts, m.isReply ? '> ' : '', out);
            out.push('');
        }
        if (threadId !== null) out.push('---', '');
        return out.join('\n');
    }

    // --- entry point ---

    const FORMATS = {
        csv: { ext: 'csv', mime: 'text/csv;charset=utf-8', build: buildCsv },
        html: { ext: 'html', mime: 'text/html;charset=utf-8', build: buildHtml },
        md: { ext: 'md', mime: 'text/markdown;charset=utf-8', build: buildMarkdown }
    };

    // Returns { content, ext, mime }.
    function build(format, conversation, opts) {
        const f = FORMATS[format];
        if (!f) throw new Error('Unknown export format: ' + format);
        return { content: f.build(conversation, opts), ext: f.ext, mime: f.mime };
    }

    TE.exporters = { FORMATS: Object.keys(FORMATS), build: build };
})(globalThis.TE = globalThis.TE || {});
