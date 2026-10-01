// Emoji reactions → [{ emoji, count }].
//
// Current Teams (Fluent UI v9) renders a toolbar of pills:
//   [data-tid="diverse-reaction-summary"]           the toolbar
//     [data-tid="diverse-reaction-pill-button"]     one per emoji
//       [data-tid="emoticon-renderer"] > img[alt]   the emoji character
//       aria-labelledby → "3 Heart reactions."      the count
//     [data-tid="add-reaction-picker-entry-point-button"]  "+" (skipped)
// Older layouts are covered by the fallback selectors.
(function (TE) {
    'use strict';

    const MAX_REACTIONS = 50;
    const MAX_COUNT = 1000000;

    const SCOPE_SELECTORS = [
        '[data-tid="diverse-reaction-summary"]',
        '[aria-label*="Reaction summary" i]',
        '[class*="fui-ChatMessage__reactions"]'
    ];
    // The last one is only safe because it runs inside the toolbar.
    const PILL_SELECTORS = [
        '[data-tid="diverse-reaction-pill-button"]',
        '[data-tid^="reaction-pill"]',
        '[class*="reaction-pill"]',
        '[class*="ReactionPill"]',
        'button[aria-pressed]'
    ];

    function firstInteger(text) {
        const m = /\d{1,9}/.exec(text || '');
        if (!m) return 0;
        const n = Number(m[0]);
        return n > MAX_COUNT ? 0 : n;
    }

    function emojiOf(pill) {
        const img = pill.querySelector('[data-tid="emoticon-renderer"] img[alt]') || pill.querySelector('img[alt]');
        if (img && TE.richtext.isEmoji(img.alt)) return img.alt.trim();
        const glyph = pill.querySelector('svg[aria-label]');
        const glyphLabel = glyph ? glyph.getAttribute('aria-label') : '';
        if (TE.richtext.isEmoji(glyphLabel)) return glyphLabel.trim();
        const m = /\p{Extended_Pictographic}️?/u.exec(pill.textContent || '');
        return m ? m[0] : '';
    }

    // The aria-labelledby text carries the real count even when the pill shows
    // "99+"; older layouts print the digit inline; no number means one.
    function countOf(pill) {
        const ids = (pill.getAttribute('aria-labelledby') || '').split(/\s+/);
        for (let i = 0; i < ids.length; i++) {
            const el = ids[i] && document.getElementById(ids[i]);
            const n = el ? firstInteger(el.textContent) : 0;
            if (n > 0) return n;
        }
        return firstInteger(pill.textContent) || 1;
    }

    // `excluded`: subtree whose reactions belong to another message.
    function extract(container, excluded) {
        if (!container) return [];
        let scope = null;
        for (let i = 0; i < SCOPE_SELECTORS.length && !scope; i++) {
            const found = container.querySelectorAll(SCOPE_SELECTORS[i]);
            for (let j = 0; j < found.length; j++) {
                if (TE.dom.outside(found[j], excluded)) {
                    scope = found[j];
                    break;
                }
            }
        }
        if (!scope) return [];

        const pills = TE.dom.pickAll(scope, PILL_SELECTORS);
        const tally = new Map();
        for (let i = 0; i < pills.length && tally.size < MAX_REACTIONS; i++) {
            const pill = pills[i];
            try {
                if (/add-reaction|picker-entry/i.test(pill.getAttribute('data-tid') || '')) continue;
                if (/^more reactions?$/i.test((pill.getAttribute('aria-label') || '').trim())) continue;
                const emoji = emojiOf(pill);
                if (!emoji) continue;
                const count = countOf(pill);
                const prior = tally.get(emoji);
                if (prior === undefined) tally.set(emoji, { emoji, count });
                else if (count > prior.count) prior.count = count;
            } catch (e) {
                // One malformed pill must not lose the message's other reactions.
                TE.log.warn('Reaction pill skipped:', e);
            }
        }
        return [...tally.values()];
    }

    TE.reactions = { extract: extract };
})(globalThis.TE = globalThis.TE || {});
