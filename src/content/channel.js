// Channel ("Posts") extraction.
//
// Channels differ from chats in ways a linear scroller can't handle:
//   - threaded: a root post with a subject line, followed by its replies;
//   - collapsed: the feed shows only the last 2-3 replies of a thread, the
//     rest sit behind an "Open N replies" button that opens a separate thread
//     view (Teams' "L2" view), which we drill into and back out of;
//   - virtualized: off-screen posts leave the DOM.
//
// Feed DOM, keyed on stable ids and data-tids (never hashed Fluent classes):
//   <div data-tid="channel-pane-message" id="reply-chain-summary-{rootId}">
//     <span id="author-{rootId}"> <time id="timestamp-{rootId}">
//     <h2 id="subject-line-{rootId}"> <div id="content-{rootId}">   root post (no body wrapper)
//     <div data-tid="response-surface" id="response-surface-{rootId}">
//       <div id="message-body-{replyId}">…                          inline replies
//       <button id="response-summary-{rootId}">                     "Open N replies"
// Message ids are the send time in epoch ms, which gives exact timestamps
// (the thread view's <time> shows a bare wall-clock time).
(function (TE) {
    'use strict';

    const dom = TE.dom;
    const SCROLL_STEP = -500;
    const MAX_DRILL_ATTEMPTS = 4;

    const SEL = {
        feedScroll: ['[data-tid="channel-pane-viewport"]', '[data-tid="message-pane-body"]'],
        feedRoot: ['[data-tid="channel-pane-message"]', '[id^="reply-chain-summary-"]'],
        threadRunway: ['[data-tid="channel-replies-runway"]', '#channel-pane-l2'],
        threadMessage: ['[data-tid="channel-replies-pane-message"]'],
        surface: ['[data-tid="response-surface"]', '[id^="response-surface-"]'],
        subject: ['[data-tid="subject-line"]', 'h2[id^="subject-line-"]'],
        author: ['[id^="author-"]', '[data-tid="message-author-name"]'],
        timestamp: ['time[data-tid="timestamp"]', 'time[id^="timestamp-"]', 'time'],
        content: ['[data-tid="message-body"][id^="content-"]', '[id^="content-"]'],
        body: [dom.BODY_SELECTOR],
        flexWrapper: ['[data-testid="message-body-flex-wrapper"][data-mid][data-reply-chain-id]'],
        // The thread view's header has a direct "Go to channel" button; the
        // "More channel options" menu holds the same action as a fallback.
        backButton: ['[data-tid="close-l2-view-button"]', 'button[aria-label="Go to channel" i]',
            '[data-tid="thread-header-back-button"]', 'button[aria-label*="Back" i]', 'button[title*="Back" i]'],
        closeMenuItem: ['[data-tid="close-l2-view-menu-item"]'],
        menuTrigger: ['[data-tid="more-apps-button-tid"]', 'button[aria-label*="More channel options" i]',
            'button[aria-haspopup="menu"][aria-label*="more" i]']
    };

    // --- where are we ---

    // On a channel surface: the feed or a thread view.
    function onChannelSurface() {
        if (document.querySelector(SEL.feedScroll[0]) || dom.pickFirst(document, SEL.feedRoot) || dom.pickFirst(document, SEL.threadRunway)) {
            return true;
        }
        // Structural fallback for renamed panes: a message whose reply-chain
        // id differs from its own id only exists on a threaded surface.
        const probe = document.querySelectorAll(SEL.flexWrapper[0]);
        const limit = Math.min(probe.length, 50);
        for (let i = 0; i < limit; i++) {
            const own = probe[i].dataset.mid;
            if (own !== probe[i].dataset.replyChainId) return true;
        }
        return false;
    }

    function isFeedPresent() {
        return !!dom.pickFirst(document, SEL.feedRoot) && !dom.pickFirst(document, SEL.threadRunway);
    }

    function findRunway() {
        const runway = dom.pickFirst(document, SEL.threadRunway);
        if (runway) return runway;
        const message = dom.pickFirst(document, SEL.threadMessage);
        return message ? message.closest('[id^="channel-pane-"]') : null;
    }

    // Prefer a scrollable feed viewport, else the viewport anyway: right after
    // leaving a thread the feed can be briefly non-scrollable, and the generic
    // fallback would grab the left navigation tree instead.
    function findFeedScroller() {
        for (let i = 0; i < SEL.feedScroll.length; i++) {
            const el = document.querySelector(SEL.feedScroll[i]);
            if (el && dom.isOverflowing(el)) return el;
        }
        return dom.pickFirst(document, SEL.feedScroll) || dom.locateChatScroller();
    }

    // --- parsing ---

    function timestampOf(id, scope) {
        if (/^\d{12,14}$/.test(id)) return TE.util.toIso(Number(id));
        const t = dom.pickFirst(scope, SEL.timestamp);
        return t ? TE.util.toIso(t.getAttribute('datetime') || t.getAttribute('aria-label') || '') : '';
    }

    // "8 replies", "8 réponses", "8 Antworten"…: the number is all we need.
    function replyCount(surface) {
        const m = /\d+/.exec(surface ? surface.getAttribute('aria-label') || '' : '');
        return m ? Number(m[0]) : 0;
    }

    // The root post of a feed container: { rootId, hasBody }. Feed roots have
    // no #message-body wrapper; the thread view's root does.
    function rootInfo(container) {
        if (container.id && container.id.startsWith('reply-chain-summary-')) {
            const rootId = container.id.slice('reply-chain-summary-'.length);
            return { rootId: rootId, hasBody: !!document.getElementById('message-body-' + rootId) };
        }
        const wrappers = dom.pickAll(container, SEL.flexWrapper);
        for (let i = 0; i < wrappers.length; i++) {
            const mid = wrappers[i].getAttribute('data-mid');
            if (mid === wrappers[i].getAttribute('data-reply-chain-id') && wrappers[i].closest(dom.BODY_SELECTOR)) {
                return { rootId: mid, hasBody: true };
            }
        }
        const body = dom.pickFirst(container, SEL.body);
        if (body && !body.closest('[data-tid="response-surface"], [id^="response-surface-"]')) {
            return { rootId: dom.bodyId(body), hasBody: true };
        }
        return null;
    }

    // Parse one post or reply. `scope` is its #message-body, or the whole feed
    // container for a feed root, whose reply chain is then `excluded`.
    function parse(session, scope, id, threadId, isReply, excluded) {
        if (!id || session.skipped.has(id)) return;
        if (session.messages.has(id)) {
            if (session.config.extractFiles) {
                try { TE.attachments.scan(session, scope, id, excluded); } catch (_) { /* rescan is best effort */ }
            }
            return;
        }
        try {
            const authorEl = document.getElementById('author-' + id) || dom.pickFirst(scope, SEL.author);
            const subjectEl = isReply ? null : document.getElementById('subject-line-' + id) || dom.pickFirst(scope, SEL.subject);
            const contentEl = document.getElementById('content-' + id) || dom.pickFirst(scope, SEL.content);
            const parsed = contentEl ? TE.richtext.parseContent(contentEl) : { blocks: [], quotes: [] };
            const author = session.authorName(authorEl);

            let attachments = [];
            if (session.config.extractFiles) {
                try {
                    attachments = TE.attachments.scan(session, scope, id, excluded);
                } catch (e) {
                    TE.log.warn('Attachment scan failed for post ' + id + ':', e);
                }
            }

            session.add({
                id: id,
                timestamp: timestampOf(id, scope),
                author: author,
                isMine: !!scope.closest('.fui-ChatMyMessage') || (!!session.userName && author === session.userName),
                isSystem: !!scope.closest('.fui-ChatControlMessageItem'),
                blocks: parsed.blocks,
                quotes: parsed.quotes,
                reactions: TE.reactions.extract(scope, excluded),
                attachments,
                subject: subjectEl ? subjectEl.textContent.trim() : '',
                threadId: threadId,
                replyToId: isReply ? threadId : '',
                isReply: isReply
            });
        } catch (e) {
            session.skipped.add(id);
            TE.log.error('Could not read post ' + id + ':', e);
        }
    }

    // --- feed ---

    // Capture every root (and its inline replies) in the feed, then drill into
    // at most one thread with hidden replies. Drilling navigates away and
    // back, which recycles the rest of the feed, so the caller simply comes
    // back for the next one. `processed` holds fully handled thread ids.
    async function extractFeed(session, processed) {
        const roots = dom.pickAll(document, SEL.feedRoot);
        const toDrill = [];
        for (let i = 0; i < roots.length && !session.halted; i++) {
            const info = rootInfo(roots[i]);
            if (!info || processed.has(info.rootId)) continue;
            const rootId = info.rootId;
            const surface = dom.pickFirst(roots[i], SEL.surface);

            if (info.hasBody) {
                const body = document.getElementById('message-body-' + rootId);
                if (body) parse(session, body, rootId, rootId, false, null);
            } else {
                parse(session, roots[i], rootId, rootId, false, surface);
            }

            const inline = surface ? dom.pickAll(surface, SEL.body) : [];
            for (let j = 0; j < inline.length; j++) parse(session, inline[j], dom.bodyId(inline[j]), rootId, true, null);

            // The count comes from a localized label; the button's presence is
            // the language-independent sign of hidden replies. Drilling twice
            // is harmless (messages are deduplicated by id).
            const total = replyCount(surface);
            if (surface && (total > inline.length || document.getElementById('response-summary-' + rootId))) {
                toDrill.push({ rootId: rootId, total: total, inline: inline.length });
            } else {
                processed.add(rootId);
            }
        }

        // Newest first (the feed lists oldest → newest): the date threshold
        // check trusts that newer threads were drilled before older ones.
        for (let i = toDrill.length - 1; i >= 0 && !session.halted; i--) {
            const job = toDrill[i];
            if (processed.has(job.rootId)) continue;
            // Only ever click this thread's own button. If the feed recycled
            // it, retry on a later pass, and give up after a few.
            if (document.getElementById('response-summary-' + job.rootId) === null) {
                const attempts = (session.drillAttempts.get(job.rootId) || 0) + 1;
                session.drillAttempts.set(job.rootId, attempts);
                if (attempts >= MAX_DRILL_ATTEMPTS) {
                    TE.log.warn('Thread ' + job.rootId + ': replies button never came back; kept the ' + job.inline + ' inline replies.');
                    processed.add(job.rootId);
                }
                continue;
            }
            try {
                await openThread(session, job);
            } catch (e) {
                TE.log.warn('Thread ' + job.rootId + ': drill-in failed:', e);
                await returnToFeed(session);
            }
            processed.add(job.rootId);
            break;
        }
    }

    // --- thread view ---

    async function waitForThread(session, maxWait) {
        const deadline = Date.now() + maxWait;
        while (Date.now() < deadline && !session.stopRequested) {
            const runway = findRunway();
            if (runway && dom.pickFirst(runway, SEL.body)) return runway;
            await TE.sleep(200);
        }
        return null;
    }

    // Every message of the open thread view; the view holds only this thread.
    function extractThread(session, rootId) {
        const runway = findRunway();
        if (!runway) return;
        const bodies = dom.pickAll(runway, SEL.body);
        for (let i = 0; i < bodies.length; i++) {
            const id = dom.bodyId(bodies[i]);
            parse(session, bodies[i], id, rootId, id !== rootId, null);
        }
    }

    // Open a thread, scroll it up to its first reply, and come back.
    async function openThread(session, job) {
        const button = document.getElementById('response-summary-' + job.rootId);
        if (!button) return;

        // Teams resets the feed to the bottom on return; restore the position
        // so the crawl resumes where it was.
        const feed = findFeedScroller();
        const savedTop = feed && feed.isConnected ? feed.scrollTop : null;
        const post = button.closest('[data-tid="channel-pane-message"]') || document.getElementById('reply-chain-summary-' + job.rootId) || button;
        post.scrollIntoView({ behavior: 'smooth', block: 'center' });
        await TE.sleep(400);

        const before = session.messages.size;
        button.click();
        const runway = await waitForThread(session, 10000);
        if (!runway) {
            TE.log.warn('Thread ' + job.rootId + ': the replies view did not open; kept ' + job.inline + ' inline replies.');
        } else {
            const scroller = dom.scrollableAncestor(runway);
            extractThread(session, job.rootId);
            // Stop when the advertised reply count is reached, at the top of
            // the thread, or after three scrolls without anything new.
            let empty = 0;
            let size = session.messages.size;
            for (let guard = 0; empty < 3 && !session.halted && guard < 2000; guard++) {
                await scrollUp(session, scroller, 3000, false);
                extractThread(session, job.rootId);
                empty = session.messages.size > size ? 0 : empty + 1;
                size = session.messages.size;
                const thread = session.threads.get(job.rootId);
                if (job.total > 0 && thread && thread.replies >= job.total) break;
                if (scroller && scroller.isConnected && scroller.scrollTop <= 0) break;
            }
            TE.log.info('Thread ' + job.rootId + ': +' + (session.messages.size - before) + ' message(s) (≈' + job.total + ' replies).');
        }
        await returnToFeed(session);
        if (savedTop !== null) {
            const scroller = findFeedScroller();
            if (scroller) scroller.scrollTop = savedTop;
        }
    }

    async function feedReady(ms) {
        const until = performance.now() + ms;
        do {
            if (isFeedPresent()) {
                await TE.sleep(300);
                return true;
            }
            await TE.sleep(200);
        } while (performance.now() < until);
        return false;
    }

    // "…" menu triggers of the thread view header, never those of messages.
    function menuTriggers() {
        const out = [];
        const add = (el) => {
            if (out.indexOf(el) !== -1 || el.closest('[data-tid="channel-replies-runway"], #channel-pane-l2')) return;
            out.push(el);
        };
        for (let i = 0; i < SEL.menuTrigger.length; i++) document.querySelectorAll(SEL.menuTrigger[i]).forEach(add);
        const popups = document.querySelectorAll('button[aria-haspopup="menu"], [role="button"][aria-haspopup="menu"]');
        for (let i = 0; i < popups.length; i++) {
            const btn = popups[i];
            const hint = (btn.getAttribute('aria-label') || '') + ' ' + (btn.dataset.tid || '');
            if (/more|option|overflow/i.test(hint)) add(btn);
        }
        return out.slice(0, 8);
    }

    // "Go to channel" lives in a menu that only exists while open: open each
    // candidate, look for the item, and close the menu again if absent.
    async function closeViaMenu() {
        let item = dom.pickFirst(document, SEL.closeMenuItem);
        if (item) {
            item.click();
            return true;
        }
        const triggers = menuTriggers();
        for (let i = 0; i < triggers.length; i++) {
            triggers[i].click();
            const deadline = Date.now() + 1200;
            while (!item && Date.now() < deadline) {
                await TE.sleep(100);
                item = dom.pickFirst(document, SEL.closeMenuItem);
            }
            if (item) {
                item.click();
                return true;
            }
            triggers[i].click();
            await TE.sleep(150);
        }
        return false;
    }

    // Leave the thread view. Opening a thread swaps the view without reliably
    // pushing history, and synthetic Escape is ignored, so click the genuine
    // "Go to channel" control first; history.back() is the last resort.
    async function returnToFeed(session) {
        if (isFeedPresent()) return true;
        const back = dom.pickFirst(document, SEL.backButton);
        if (back) {
            back.click();
            if (await feedReady(4000)) return true;
        }
        if (await closeViaMenu() && await feedReady(4000)) return true;
        history.back();
        if (await feedReady(5000)) return true;
        TE.log.warn('Could not get back to the channel feed after opening a thread.');
        return false;
    }

    // --- scrolling ---

    // Scroll a feed/thread viewport one step up and wait for new posts.
    // At the top, Teams fetches older posts only when the viewport enters the
    // near-top zone, so wiggling in place does nothing: leave, then come back
    // in small wheel steps like a user. `forceTop` jumps through the whole
    // list when Teams keeps snapping the position back.
    async function scrollUp(session, container, maxWait, forceTop) {
        if (container !== null && container !== undefined && container.isConnected) {
            const before = container.scrollTop;
            if (forceTop && before > 0) {
                container.scrollTop = container.scrollHeight;
                await TE.sleep(500);
                container.scrollTop = 0;
                dom.wheel(container, -500);
                try {
                    container.dispatchEvent(new KeyboardEvent('keydown', { key: 'Home', code: 'Home', bubbles: true, cancelable: true }));
                } catch (_) { /* ignore */ }
            } else if (before <= 0) {
                const away = Math.min(container.scrollHeight - container.clientHeight, Math.round(container.clientHeight * 1.5));
                if (away > 0) {
                    container.scrollTop = away;
                    await TE.sleep(400);
                }
                for (let y = container.scrollTop; y > 0; y -= 150) {
                    container.scrollTop = Math.max(0, y - 150);
                    dom.wheel(container, -150);
                    await TE.sleep(120);
                }
                container.scrollTop = 0;
                dom.wheel(container, -500);
            } else {
                container.scrollTop = Math.max(0, before + SCROLL_STEP);
                if (container.scrollTop === before) container.scrollBy(0, SCROLL_STEP);
            }
        } else {
            window.scrollBy(0, SCROLL_STEP);
        }
        await dom.awaitFreshMessages(session, maxWait, container);
    }

    // Threads oldest first; in a thread, the root then replies in send order.
    function threadOrder(a, b) {
        const rootA = Number(a.threadId || a.id);
        const rootB = Number(b.threadId || b.id);
        if (rootA !== rootA || rootB !== rootB) return 0;
        if (rootA !== rootB) return rootA - rootB;
        if (a.isReply !== b.isReply) return Number(!!a.isReply) - Number(!!b.isReply);
        return Number(a.id) - Number(b.id);
    }

    // Extraction started inside a thread view: extract that thread.
    async function crawlOpenThread(session) {
        TE.log.info('Thread view open: extracting this thread only.');
        const runway = dom.pickFirst(document, SEL.threadRunway);
        let rootId = null;
        const wrappers = dom.pickAll(runway || document, SEL.flexWrapper);
        for (let i = 0; i < wrappers.length && !rootId; i++) {
            const mid = wrappers[i].getAttribute('data-mid');
            if (mid === wrappers[i].getAttribute('data-reply-chain-id')) rootId = mid;
        }
        if (!rootId) {
            const first = dom.pickFirst(runway || document, SEL.body);
            rootId = first ? dom.bodyId(first) : null;
        }
        if (!rootId) {
            TE.log.warn('Could not identify the thread root: nothing extracted.');
            return [];
        }

        extractThread(session, rootId);
        let stalled = 0;
        let size = session.messages.size;
        while (stalled < 3 && !session.halted) {
            const scroller = dom.scrollableAncestor(dom.pickFirst(document, SEL.threadRunway) || document.body);
            const top = scroller ? scroller.scrollTop : 0;
            await scrollUp(session, scroller, 3000, false);
            extractThread(session, rootId);
            const after = scroller ? scroller.scrollTop : 0;
            const grew = session.messages.size > size;
            if (after <= 0 && !grew) break; // top reached, nothing new
            stalled = !grew && after >= top - 1 ? stalled + 1 : 0;
            size = session.messages.size;
        }
        await session.settleCaptures();
        return Array.from(session.messages.values()).sort(threadOrder);
    }

    async function crawl(session) {
        if (dom.pickFirst(document, SEL.threadRunway) && !dom.pickFirst(document, SEL.feedRoot)) return crawlOpenThread(session);

        TE.log.info('Channel view: extracting threaded posts.');
        const processed = new Set();
        let scroller = findFeedScroller();
        if (scroller) scroller.scrollTop = scroller.scrollHeight;
        else window.scrollTo(0, document.body.scrollHeight);
        await TE.sleep(500);
        await dom.awaitFreshMessages(session, 3000, scroller);
        await extractFeed(session, processed);

        // The feed pages in older threads in batches and a scroll often lands
        // between two: keep going for 8 empty passes.
        const maxEmpty = 8;
        let empty = 0, atTop = 0, stuck = 0, passes = 0;
        let previousRoots = processed.size;

        while (empty < maxEmpty && !session.halted) {
            passes++;
            // Re-acquire every pass: returning from a thread re-renders the feed.
            scroller = findFeedScroller();
            // Teams re-focuses the last thread's replies button after a drill,
            // scrolling it back into view; drop focus so our scroll sticks.
            const active = document.activeElement;
            if (active && active !== document.body && typeof active.blur === 'function') active.blur();

            const beforeTop = scroller ? scroller.scrollTop : 0;
            // At the top, fetching older posts is a network round trip: give
            // each confirmation pass longer before calling the feed exhausted.
            await scrollUp(session, scroller, atTop > 0 ? 3000 + atTop * 2000 : 3000, stuck >= 2);
            const rootsBefore = processed.size;
            await extractFeed(session, processed);

            // Threads are ordered by last activity. Once a fully processed
            // thread's latest message is at/before the threshold, every thread
            // above it is too: stop.
            if (!isNaN(session.threshold) && processed.size > rootsBefore) {
                let index = 0, reached = false;
                for (const id of processed) {
                    if (index++ < rootsBefore) continue;
                    const thread = session.threads.get(id);
                    if (thread && thread.latest > -Infinity && thread.latest <= session.threshold) {
                        reached = true;
                        break;
                    }
                }
                if (reached) {
                    TE.log.info('Reached the "stop at date" threshold.');
                    break;
                }
            }

            const live = scroller && scroller.isConnected ? scroller : findFeedScroller();
            const afterTop = live ? live.scrollTop : 0;
            const grew = processed.size > previousRoots;
            const movedUp = beforeTop - afterTop > 1;
            // Not at the top, nothing new, no movement: Teams snapped us back.
            stuck = !grew && !movedUp && afterTop > 0 ? stuck + 1 : 0;
            if (grew || movedUp) {
                // New roots, or the view really moved up: lazy loading may lag
                // a pass or two behind, so this is not an empty pass.
                empty = 0;
                atTop = 0;
            } else if (stuck === 0 || stuck > 10) {
                empty++;
                atTop = afterTop <= 0 ? atTop + 1 : 0;
                if (atTop === 1) TE.log.info('Reached the top of the feed; confirming there is nothing older.');
            }
            if (atTop >= 4) break;
            previousRoots = processed.size;
        }
        TE.log.info('Feed crawl done after ' + passes + ' pass(es): ' + processed.size + ' thread(s).');
        await session.settleCaptures();
        return Array.from(session.messages.values()).sort(threadOrder);
    }

    TE.channel = { onChannelSurface: onChannelSurface, crawl: crawl };
})(globalThis.TE = globalThis.TE || {});
