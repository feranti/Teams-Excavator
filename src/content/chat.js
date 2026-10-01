// Chat extraction: one linear, virtualized message list scrolled from the
// newest message up to the oldest.
(function (TE) {
    'use strict';

    const dom = TE.dom;
    const SCROLL_STEP = -500;
    // Consecutive scrolls without new messages before the history is
    // considered exhausted.
    const MAX_EMPTY_SCROLLS = 5;

    function containerOf(body) {
        return body.closest('[data-tid="chat-pane-item"]') || body.closest('.fui-ChatMessage, .fui-ChatMyMessage');
    }

    // System messages (meeting started, member added…) may have no <time>;
    // their text then starts with a local "DD/MM/YYYY HH:MM".
    function systemTimestamp(body) {
        const m = /^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})/.exec((body.textContent || '').trim());
        return m ? new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]), Number(m[4]), Number(m[5])).toISOString() : '';
    }

    function scanAttachments(session, container, id) {
        try {
            return TE.attachments.scan(session, container, id, null);
        } catch (e) {
            TE.log.warn('Attachment scan failed for message ' + id + ':', e);
            return session.attachmentState(id).list;
        }
    }

    function parseMessage(session, body) {
        const id = dom.bodyId(body);
        if (!id || session.skipped.has(id)) return;
        const container = containerOf(body);

        if (session.messages.has(id)) {
            // Seen before. Teams loads images lazily and recycles nodes, so
            // rescan: new finds merge into the stored attachment list.
            if (session.config.extractFiles && container) scanAttachments(session, container, id);
            return;
        }

        try {
            const authorEl = container && container.querySelector('[data-tid="message-author-name"]');
            const author = session.authorName(authorEl);
            const isSystem = !!container && container.classList.contains('fui-ChatControlMessageItem');
            const timeEl = container && container.querySelector('time');
            let timestamp = TE.util.toIso(timeEl ? timeEl.getAttribute('datetime') : '');
            if (!timestamp && isSystem) timestamp = systemTimestamp(body);

            const content = body.querySelector('[id^="content-"]');
            const parsed = content ? TE.richtext.parseContent(content) : { blocks: [], quotes: [] };

            session.add({
                id: id,
                timestamp: timestamp,
                author: author,
                isMine: !!body.closest('.fui-ChatMyMessage, [class*="mine"], [class*="myMessage"]') ||
                    (!!session.userName && author === session.userName),
                isSystem: isSystem,
                blocks: parsed.blocks,
                quotes: parsed.quotes,
                reactions: TE.reactions.extract(container, null),
                // The session's list itself: later rescans append to it.
                attachments: session.config.extractFiles ? scanAttachments(session, container, id) : []
            });
        } catch (e) {
            session.skipped.add(id);
            TE.log.error('Could not read message ' + id + ':', e);
        }
    }

    function extractVisible(session) {
        const bodies = document.querySelectorAll(dom.BODY_SELECTOR);
        for (let i = 0; i < bodies.length; i++) parseMessage(session, bodies[i]);
    }

    function scrollUp(scroller) {
        if (scroller) {
            const before = scroller.scrollTop;
            scroller.scrollTop = before + SCROLL_STEP;
            if (scroller.scrollTop === before) scroller.scrollBy(0, SCROLL_STEP);
        } else {
            window.scrollBy(0, SCROLL_STEP);
        }
        // Bringing the oldest rendered message into view is what makes Teams
        // page in older history.
        const first = document.querySelector(dom.BODY_SELECTOR);
        if (first) first.scrollIntoView({ behavior: 'smooth', block: 'start' });
    }

    function byTime(a, b) {
        const ta = Date.parse(a.timestamp), tb = Date.parse(b.timestamp);
        if (isNaN(ta)) return isNaN(tb) ? 0 : 1; // undated messages go last
        if (isNaN(tb)) return -1;
        return ta - tb;
    }

    async function crawl(session) {
        let scroller = dom.locateChatScroller();
        if (!scroller) TE.log.warn('No message list scroller found; scrolling the window.');

        // Start from the newest message wherever the user was scrolled.
        if (scroller) scroller.scrollTop = scroller.scrollHeight;
        else window.scrollTo(0, document.body.scrollHeight);
        await TE.sleep(500);
        await dom.awaitFreshMessages(session, 3000, scroller);
        scroller = scroller || dom.locateChatScroller();
        extractVisible(session);

        let empty = 0;
        let lastSize = session.messages.size;
        while (empty < MAX_EMPTY_SCROLLS && !session.halted) {
            if (!scroller) scroller = dom.locateChatScroller();
            // Pinned at the top with nothing new: the history is almost surely
            // exhausted, so wait less and give up sooner.
            const pinned = empty > 0 && !!scroller && scroller.scrollTop <= 0;
            scrollUp(scroller);
            await dom.awaitFreshMessages(session, pinned ? 1500 : 3000, scroller);
            extractVisible(session);

            if (session.reachedThreshold()) {
                TE.log.info('Reached the "stop at date" threshold.');
                break;
            }
            if (session.messages.size > lastSize) {
                lastSize = session.messages.size;
                empty = 0;
                continue;
            }
            empty++;
            TE.log.info('No new messages (' + empty + '/' + MAX_EMPTY_SCROLLS + ')');
            if (pinned && empty >= 3) break;
            if (empty >= 2) {
                const first = document.querySelector(dom.BODY_SELECTOR);
                if (first) {
                    first.scrollIntoView({ behavior: 'auto', block: 'start' });
                    await dom.awaitFreshMessages(session, 3000, scroller);
                    extractVisible(session);
                }
            }
        }

        extractVisible(session);
        await session.settleCaptures();
        TE.log.info('Chat extraction finished: ' + session.messages.size + ' message(s).');
        return Array.from(session.messages.values()).sort(byTime);
    }

    TE.chat = { crawl: crawl };
})(globalThis.TE = globalThis.TE || {});
