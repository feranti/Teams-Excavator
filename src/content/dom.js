// DOM helpers shared by the chat and channel crawlers.
(function (TE) {
    'use strict';

    const BODY_SELECTOR = '[id^="message-body-"]';
    const BODY_PREFIX_LENGTH = 'message-body-'.length;
    const EMPTY = [];

    // First element matching one selector of a fallback chain (ordered most
    // specific first), or null.
    function pickFirst(scope, selectors) {
        for (let i = 0; i < selectors.length; i++) {
            const el = scope.querySelector(selectors[i]);
            if (el) return el;
        }
        return null;
    }

    // All elements matching the first selector of the chain that matches any.
    function pickAll(scope, selectors) {
        for (let i = 0; i < selectors.length; i++) {
            const list = scope.querySelectorAll(selectors[i]);
            if (list.length) return list;
        }
        return EMPTY;
    }

    // True when `el` is not inside `excluded` (a subtree to ignore, e.g. the
    // replies under a channel root post).
    function outside(el, excluded) {
        return !excluded || !excluded.contains(el);
    }

    function bodyId(el) {
        return el.id.slice(BODY_PREFIX_LENGTH);
    }

    function isOverflowing(el) {
        return el.scrollHeight > el.clientHeight || el.scrollTop > 0;
    }

    // The chat's message list scroller.
    const SCROLL_SELECTORS = [
        '[data-tid="message-pane-list"]',
        '[data-tid="chat-pane-list"]',
        '.fui-Tree',
        '[role="tree"]',
        '[class*="virtual-tree"]',
        '[class*="message-list"]',
        '[class*="scrollable"]',
        '[role="presentation"]'
    ];

    function locateChatScroller() {
        for (let i = 0; i < SCROLL_SELECTORS.length; i++) {
            const list = document.querySelectorAll(SCROLL_SELECTORS[i]);
            for (let j = 0; j < list.length; j++) {
                if (isOverflowing(list[j])) return list[j];
            }
        }
        // Last resort: any mid-sized element with its own vertical scrolling.
        const all = document.body.getElementsByTagName('*');
        for (let i = 0; i < all.length; i++) {
            const el = all[i];
            if (el.clientHeight <= 100 || el.clientHeight >= window.innerHeight || el.scrollHeight <= el.clientHeight) continue;
            const overflow = getComputedStyle(el).overflowY;
            if (overflow === 'auto' || overflow === 'scroll') return el;
        }
        return null;
    }

    // Nearest scrolling ancestor of `start` (the channel thread view has no
    // stable selector for its viewport).
    function scrollableAncestor(start) {
        for (let el = start; el && el !== document.body && el !== document.documentElement; el = el.parentElement) {
            const overflow = getComputedStyle(el).overflowY;
            if ((overflow === 'auto' || overflow === 'scroll') && el.scrollHeight > el.clientHeight + 4) return el;
        }
        return document.querySelector('[data-tid="channel-pane-viewport"]') || locateChatScroller();
    }

    // Is there a message in the DOM the session has not processed yet?
    // Cheap enough to run on every mutation batch.
    function hasUnseenMessage(session) {
        const bodies = document.querySelectorAll(BODY_SELECTOR);
        for (let i = 0; i < bodies.length; i++) {
            if (!session.seen(bodyId(bodies[i]))) return true;
        }
        const roots = document.querySelectorAll('[id^="reply-chain-summary-"]');
        for (let i = 0; i < roots.length; i++) {
            if (!session.seen(roots[i].id.slice('reply-chain-summary-'.length))) return true;
        }
        return false;
    }

    // Resolve true as soon as check() holds (on DOM mutations under `observed`
    // and on a polling fallback), false after `maxWait` ms or on stop.
    function waitFor(session, check, maxWait, observed) {
        return new Promise((resolve) => {
            let settled = false;
            let observer = null;
            const finish = (value) => {
                if (settled) return;
                settled = true;
                if (observer) observer.disconnect();
                resolve(value);
            };
            if (check()) {
                finish(true);
                return;
            }
            if (observed && observed.isConnected) {
                observer = new MutationObserver(() => { if (check()) finish(true); });
                observer.observe(observed, { childList: true, subtree: true });
            }
            const deadline = Date.now() + maxWait;
            (async () => {
                while (!settled) {
                    await TE.sleep(250);
                    if (session.stopRequested) finish(false);
                    else if (check()) finish(true);
                    else if (Date.now() >= deadline) finish(false);
                }
            })();
        });
    }

    function awaitFreshMessages(session, maxWait, observed) {
        return waitFor(session, () => hasUnseenMessage(session), maxWait, observed);
    }

    function wheel(el, deltaY) {
        try {
            el.dispatchEvent(new WheelEvent('wheel', { deltaY: deltaY, bubbles: true, cancelable: true }));
        } catch (_) { /* ignore */ }
    }

    TE.dom = {
        BODY_SELECTOR: BODY_SELECTOR,
        pickFirst: pickFirst,
        pickAll: pickAll,
        outside: outside,
        bodyId: bodyId,
        isOverflowing: isOverflowing,
        locateChatScroller: locateChatScroller,
        scrollableAncestor: scrollableAncestor,
        waitFor: waitFor,
        awaitFreshMessages: awaitFreshMessages,
        wheel: wheel
    };
})(globalThis.TE = globalThis.TE || {});
