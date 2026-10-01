// Message body parsing: Teams message HTML → a small structured model the
// exporters render to CSV, HTML and Markdown.
//
//   body  = Block[]
//   Block = { kind: 'text',  lines: Run[][] }          lines of runs; [] is a blank line
//         | { kind: 'code',  text }                    <pre>, indentation kept
//         | { kind: 'quote', blocks: Block[] }         a <blockquote> typed in the message
//         | { kind: 'list',  ordered, items: Block[][] }
//         | { kind: 'reply', index }                   where quoted-reply card quotes[index] sat
//   Run   = { text, b?, i?, u?, s?, mention?, href? }  flags only present when true
//
// Quoted replies are returned beside the body as [{ author, date, text }].
(function (TE) {
    'use strict';

    const toIso = TE.util.toIso;

    const BLOCK_TAGS = new Set([
        'P', 'DIV', 'LI', 'TR', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6',
        'SECTION', 'ARTICLE', 'HEADER', 'FOOTER', 'TABLE', 'FIGURE', 'DL', 'DT', 'DD'
    ]);
    const SKIPPED_TAGS = new Set(['SCRIPT', 'STYLE', 'TEMPLATE', 'svg', 'NOSCRIPT']);
    const REPLY_TAG = 'TE-REPLY';
    const PLAIN = {};

    // --- emoji ---

    function isEmoji(s) {
        if (typeof s !== 'string') return false;
        const t = s.trim();
        if (t.length === 0 || t.length > 16) return false;
        return /^:[\w+-]+:$/.test(t) || /\p{Extended_Pictographic}/u.test(t);
    }

    // The character an emoji <img> stands for, or '' for a content image.
    function emojiText(img) {
        const alt = (img.getAttribute('alt') || '').trim();
        if (!alt) return '';
        if (img.closest('[data-tid="emoticon-renderer"]') ||
            /schema\.skype\.com\/Emoji/i.test(img.getAttribute('itemtype') || '') ||
            isEmoji(alt)) return alt;
        return '';
    }

    // --- inline styles ---

    function withFlag(style, flag) {
        if (style[flag]) return style;
        const s = Object.assign({}, style);
        s[flag] = true;
        return s;
    }

    function inlineStyle(el, tag, style) {
        switch (tag) {
            case 'B': case 'STRONG': return withFlag(style, 'b');
            case 'I': case 'EM': return withFlag(style, 'i');
            case 'U': case 'INS': return withFlag(style, 'u');
            case 'S': case 'STRIKE': case 'DEL': return withFlag(style, 's');
        }
        const css = el.getAttribute('style');
        if (!css) return style;
        let s = style;
        if (/text-decoration[^;]*underline/i.test(css)) s = withFlag(s, 'u');
        if (/text-decoration[^;]*line-through/i.test(css)) s = withFlag(s, 's');
        if (/font-weight\s*:\s*(bold|[6-9]00)/i.test(css)) s = withFlag(s, 'b');
        if (/font-style\s*:\s*italic/i.test(css)) s = withFlag(s, 'i');
        return s;
    }

    function sameStyle(run, style) {
        return !run.b === !style.b && !run.i === !style.i && !run.u === !style.u && !run.s === !style.s &&
            !run.mention === !style.mention && (run.href || '') === (style.href || '');
    }

    function makeRun(text, style) {
        const run = { text: text };
        if (style.b) run.b = true;
        if (style.i) run.i = true;
        if (style.u) run.u = true;
        if (style.s) run.s = true;
        if (style.mention) run.mention = true;
        if (style.href) run.href = style.href;
        return run;
    }

    // --- block builder ---

    // Accumulates lines of runs; non-text blocks flush the pending lines.
    class BlockBuilder {
        constructor() {
            this.blocks = [];
            this.lines = null;
            this.line = null;
            this.lineHasContent = false;
        }

        newLine() {
            if (!this.lines) this.lines = [];
            this.line = [];
            this.lines.push(this.line);
            this.lineHasContent = false;
        }

        text(str, style) {
            if (!str) return;
            if (!this.line) this.newLine();
            const line = this.line;
            const last = line.length ? line[line.length - 1] : null;
            if (last && sameStyle(last, style)) last.text += str;
            else line.push(makeRun(str, style));
            if (!this.lineHasContent && /\S/.test(str)) this.lineHasContent = true;
        }

        // Explicit <br>: always ends the line, so <br><br> keeps a blank line.
        lineBreak() {
            if (!this.line) this.newLine();
            this.newLine();
        }

        // Block element boundary: a new line only if the current one has text.
        // Teams nests a <div> per visual line; this never doubles breaks.
        softBreak() {
            if (!this.line) return;
            if (this.lineHasContent) this.newLine();
            else this.line.length = 0;
        }

        block(block) {
            this.flush();
            this.blocks.push(block);
        }

        flush() {
            if (!this.lines) return;
            pushTextBlocks(normalizeLines(this.lines), this.blocks);
            this.lines = null;
            this.line = null;
            this.lineHasContent = false;
        }

        finish() {
            this.flush();
            return this.blocks;
        }
    }

    const WHITESPACE_RE = /\s+/g;

    // Collapse whitespace like a browser would, trim both line ends and drop
    // empty runs.
    function normalizeRuns(line) {
        const out = [];
        let afterSpace = true; // leading spaces of a line are dropped
        for (let i = 0; i < line.length; i++) {
            const run = line[i];
            let t = run.text.replace(WHITESPACE_RE, ' ');
            if (afterSpace && t.charCodeAt(0) === 32) t = t.slice(1);
            if (!t) continue;
            afterSpace = t.charCodeAt(t.length - 1) === 32;
            run.text = t;
            out.push(run);
        }
        while (out.length) {
            const last = out[out.length - 1];
            const t = last.text.replace(/ +$/, '');
            if (t) {
                last.text = t;
                break;
            }
            out.pop();
        }
        return out;
    }

    // Normalize every line, keep at most one blank line in a row, and drop
    // blank lines at both ends.
    function normalizeLines(lines) {
        const out = [];
        let previousBlank = true;
        for (let i = 0; i < lines.length; i++) {
            const runs = normalizeRuns(lines[i]);
            if (!runs.length && previousBlank) continue;
            out.push(runs);
            previousBlank = !runs.length;
        }
        if (out.length && !out[out.length - 1].length) out.pop();
        return out;
    }

    function isBulletOnly(line) {
        return line.length === 1 && /^[•◦‣▪▸►◆●]$/.test(line[0].text);
    }

    function pushText(lines, from, to, blocks) {
        while (from < to && !lines[from].length) from++;
        while (to > from && !lines[to - 1].length) to--;
        if (to > from) blocks.push({ kind: 'text', lines: lines.slice(from, to) });
    }

    // Teams sometimes renders a bullet as its own line ("•") followed by the
    // item's text. Turn those pairs back into a real list.
    function pushTextBlocks(lines, blocks) {
        let segmentStart = 0;
        let items = null;
        let i = 0;
        while (i < lines.length) {
            if (i + 1 < lines.length && isBulletOnly(lines[i]) && lines[i + 1].length) {
                if (!items) {
                    pushText(lines, segmentStart, i, blocks);
                    items = [];
                }
                items.push([{ kind: 'text', lines: [lines[i + 1]] }]);
                i += 2;
                segmentStart = i;
                continue;
            }
            if (items && lines[i].length) {
                blocks.push({ kind: 'list', ordered: false, items: items });
                items = null;
            }
            i++;
        }
        if (items) blocks.push({ kind: 'list', ordered: false, items: items });
        else pushText(lines, segmentStart, lines.length, blocks);
    }

    // --- DOM walk ---

    function walkChildren(node, b, style) {
        for (let child = node.firstChild; child; child = child.nextSibling) visit(child, b, style);
    }

    function parseChildren(node, style) {
        const b = new BlockBuilder();
        walkChildren(node, b, style);
        return b.finish();
    }

    function visit(node, b, style) {
        if (node.nodeType === 3) {
            b.text(node.nodeValue, style);
            return;
        }
        if (node.nodeType !== 1) return;
        const tag = node.tagName;
        if (SKIPPED_TAGS.has(tag)) return;

        if (tag === REPLY_TAG) {
            b.block({ kind: 'reply', index: Number(node.getAttribute('data-index')) });
            return;
        }
        if (node.hasAttribute('data-mention-type')) {
            const name = node.textContent.replace(/^@/, '').trim();
            if (name) b.text(name, withFlag(style, 'mention'));
            return;
        }

        switch (tag) {
            case 'BR':
                b.lineBreak();
                return;
            case 'IMG': {
                const emoji = emojiText(node);
                if (emoji) b.text(emoji, style);
                return;
            }
            case 'PRE': {
                const text = preformattedText(node);
                if (text) b.block({ kind: 'code', text: text });
                return;
            }
            case 'UL':
            case 'OL':
                list(node, b, style, tag === 'OL');
                return;
            case 'BLOCKQUOTE': {
                const blocks = parseChildren(node, style);
                if (blocks.length) b.block({ kind: 'quote', blocks: blocks });
                return;
            }
            case 'A': {
                const href = (node.getAttribute('href') || '').trim();
                if (/^https?:/i.test(href)) {
                    const linked = Object.assign({}, style);
                    linked.href = href;
                    walkChildren(node, b, linked);
                    return;
                }
                break;
            }
            case 'TD':
            case 'TH':
                walkChildren(node, b, style);
                b.text(' ', style);
                return;
        }

        const inner = inlineStyle(node, tag, style);
        const isBlock = BLOCK_TAGS.has(tag);
        if (isBlock) b.softBreak();
        walkChildren(node, b, inner);
        if (isBlock) b.softBreak();
    }

    function list(node, b, style, ordered) {
        const items = [];
        for (let child = node.firstElementChild; child; child = child.nextElementSibling) {
            let blocks;
            if (child.tagName === 'LI') {
                blocks = parseChildren(child, style);
            } else {
                // A list nested directly in a list (invalid, but editors do it).
                const inner = new BlockBuilder();
                visit(child, inner, style);
                blocks = inner.finish();
            }
            if (blocks.length) items.push(blocks);
        }
        if (items.length) b.block({ kind: 'list', ordered: ordered, items: items });
    }

    // Code block text: <br> and block children become newlines; indentation
    // and blank lines inside the code are preserved.
    function preformattedText(pre) {
        let out = '';
        const walk = (node) => {
            for (let c = node.firstChild; c; c = c.nextSibling) {
                if (c.nodeType === 3) {
                    out += c.nodeValue;
                } else if (c.nodeType === 1) {
                    if (c.tagName === 'BR') {
                        out += '\n';
                        continue;
                    }
                    const isBlock = c.tagName === 'DIV' || c.tagName === 'P';
                    if (isBlock && out && !out.endsWith('\n')) out += '\n';
                    walk(c);
                    if (isBlock && out && !out.endsWith('\n')) out += '\n';
                }
            }
        };
        walk(pre);
        return out.replace(/ /g, ' ').replace(/\r\n?/g, '\n').replace(/[ \t]+$/gm, '').replace(/^\n+|\n+$/g, '');
    }

    // --- quoted replies ---

    // Teams has shipped many quoted-reply markups; any of these may be one.
    const QUOTE_SELECTOR = [
        '[data-track-action-scenario="messageQuotedReply"]',
        '[data-tid="quoted-reply-card"]',
        '[data-tid="quoted-message"]',
        '[data-tid="quoted-message-container"]',
        '[data-tid="reply-context"]',
        '[data-tid*="quote"]',
        '[data-tid*="reply"]',
        '[data-testid*="quote"]',
        '[class*="fui-Quote"]',
        '[class*="quoted-message"]',
        '[class*="quote-container"]',
        '[class*="reply-context"]',
        'blockquote',
        '[role="complementary"]'
    ].join(', ');

    const QUOTE_AUTHOR_SELECTOR = '[data-tid="quoted-author"], [data-tid="message-author-name"], cite, header, [class*="author"]';

    // A candidate is a real quoted reply only if it names an author or a time
    // (or is the classic card): the broad selectors above also match plain
    // wrappers and typed blockquotes, which must stay in the body.
    function isQuoteCard(node) {
        return node.getAttribute('data-track-action-scenario') === 'messageQuotedReply' ||
            node.getAttribute('data-tid') === 'quoted-reply-card' ||
            !!node.querySelector('[data-tid="quoted-reply-preview-content"], [data-tid="quoted-reply-timestamp"]') ||
            !!node.querySelector(QUOTE_AUTHOR_SELECTOR) ||
            !!node.querySelector('time[datetime]');
    }

    // Pull the quoted-reply cards out of `root` (a detached clone), replacing
    // each with a marker so the body keeps the quote at its position.
    function extractQuoteCards(root) {
        const quotes = [];
        const candidates = root.querySelectorAll(QUOTE_SELECTOR);
        const cards = [];
        for (let i = 0; i < candidates.length; i++) {
            const node = candidates[i];
            if (!isQuoteCard(node)) continue;
            // Document order: an enclosing card always comes first.
            let nested = false;
            for (let j = 0; j < cards.length && !nested; j++) nested = cards[j].contains(node);
            if (!nested) cards.push(node);
        }

        for (let i = 0; i < cards.length; i++) {
            const card = cards[i];
            const stamp = card.querySelector('[data-tid="quoted-reply-timestamp"]');
            let authorEl = card.querySelector(QUOTE_AUTHOR_SELECTOR);
            // The classic card has no author attribute: the author is the
            // other non-empty element next to the timestamp.
            if (!authorEl && stamp && stamp.parentElement) {
                for (let c = stamp.parentElement.firstElementChild; c; c = c.nextElementSibling) {
                    if (c !== stamp && c.textContent.trim()) {
                        authorEl = c;
                        break;
                    }
                }
            }
            const author = authorEl ? authorEl.textContent.trim() : '';
            if (authorEl) authorEl.remove();

            const timeEl = card.querySelector('time[datetime]') || stamp;
            const date = timeEl ? toIso(timeEl.getAttribute('datetime') || timeEl.textContent.trim()) : '';
            if (timeEl) timeEl.remove();

            const preview = card.querySelector('[data-tid="quoted-reply-preview-content"]');
            const text = TE.util.blocksToText(parseChildren(preview || card, PLAIN));

            if (!author && !text) {
                card.remove();
                continue;
            }
            const marker = document.createElement(REPLY_TAG);
            marker.setAttribute('data-index', String(quotes.length));
            quotes.push({ author: author, date: date, text: text });
            card.replaceWith(marker);
        }
        return quotes;
    }

    // Fallback when no card markup matched: the flattened text itself reads
    //   Author M/D/YYYY h:mm AM      (author and date on one or two lines)
    //   quoted text
    //   <blank line>
    //   reply
    // Only looked for in the first two lines, so a date mentioned later in an
    // ordinary message is never mistaken for a quote header.
    const QUOTE_DATE_RE = /(\d{1,2}\/\d{1,2}\/\d{2,4}),?\s+\d{1,2}:\d{2}(?:\s*[AaPp][Mm])?/;

    function applyTextualQuote(blocks, quotes) {
        const first = blocks[0];
        if (!first || first.kind !== 'text' || first.lines.length < 2) return;
        const texts = new Array(first.lines.length);
        for (let i = 0; i < texts.length; i++) texts[i] = TE.util.runsToText(first.lines[i]);

        let dateLine = -1;
        let match = null;
        for (let i = 0; i < 2 && dateLine === -1; i++) {
            match = QUOTE_DATE_RE.exec(texts[i]);
            if (match) dateLine = i;
        }
        if (dateLine === -1) return;
        const author = texts[dateLine].slice(0, match.index).trim() || (dateLine > 0 ? texts[0].trim() : '');
        if (!author) return;

        let blank = -1;
        for (let i = dateLine + 1; i < texts.length && blank === -1; i++) {
            if (!texts[i]) blank = i;
        }
        let quoteText, restStart;
        if (blank !== -1 && blank < texts.length - 1) {
            quoteText = texts.slice(dateLine + 1, blank).join('\n').trim();
            restStart = blank + 1;
        } else if (texts.length >= dateLine + 3) {
            quoteText = texts[dateLine + 1];
            restStart = dateLine + 2;
        } else {
            quoteText = '';
            restStart = dateLine + 1;
        }
        const rest = first.lines.slice(restStart);
        if (!rest.length && blocks.length === 1) return; // nothing left: not a reply

        quotes.push({ author: author, date: toIso(match[0]), text: quoteText });
        if (rest.length) first.lines = rest;
        else blocks.shift();
        blocks.unshift({ kind: 'reply', index: 0 });
    }

    // Parse a message content element (left untouched: we work on a clone).
    function parseContent(contentEl) {
        const root = contentEl.cloneNode(true);
        const quotes = extractQuoteCards(root);
        const blocks = parseChildren(root, PLAIN);
        if (!quotes.length) applyTextualQuote(blocks, quotes);
        return { blocks: blocks, quotes: quotes };
    }

    // Quote cards carry a one-line snippet. When the quoted message is part of
    // the same extraction, replace the snippet with its full text. Matching is
    // by author + minute: the quote header only shows minutes.
    function authorMinuteKey(iso, author) {
        if (!iso || !author) return null;
        const stamp = String(iso);
        return stamp.length >= 16 && stamp.charAt(10) === 'T' ? author + '|' + stamp.substring(0, 16) : null;
    }

    function resolveQuotes(messages) {
        const byMinute = new Map();
        for (let i = 0; i < messages.length; i++) {
            const key = authorMinuteKey(messages[i].timestamp, messages[i].author);
            if (key && !byMinute.has(key)) byMinute.set(key, messages[i]);
        }
        let resolved = 0;
        for (let i = 0; i < messages.length; i++) {
            const quotes = messages[i].quotes;
            for (let q = 0; q < quotes.length; q++) {
                const original = byMinute.get(authorMinuteKey(quotes[q].date, quotes[q].author));
                if (!original || original === messages[i]) continue;
                const text = TE.util.blocksToText(original.blocks);
                if (text.length > (quotes[q].text || '').length) {
                    quotes[q].text = text;
                    quotes[q].resolved = true;
                    resolved++;
                }
            }
        }
        if (resolved) TE.log.info('Completed ' + resolved + ' truncated quote(s) from the full messages.');
    }

    TE.richtext = {
        isEmoji: isEmoji,
        parseContent: parseContent,
        resolveQuotes: resolveQuotes
    };
})(globalThis.TE = globalThis.TE || {});
