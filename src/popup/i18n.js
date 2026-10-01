// Popup translations. chrome.i18n always follows the browser language; this
// loader lets the user pick another one, reading the same _locales files.
(function (TE) {
    'use strict';

    const LOCALES = [
        ['en', '🇬🇧 English'], ['es', '🇪🇸 Español'], ['fr', '🇫🇷 Français'], ['de', '🇩🇪 Deutsch'],
        ['it', '🇮🇹 Italiano'], ['nl', '🇳🇱 Nederlands'], ['pl', '🇵🇱 Polski'], ['pt_BR', '🇧🇷 Português'],
        ['ru', '🇷🇺 Русский'], ['uk', '🇺🇦 Українська'], ['cs', '🇨🇿 Čeština'], ['da', '🇩🇰 Dansk'],
        ['zh_CN', '🇨🇳 简体中文'], ['zh_TW', '🇹🇼 繁體中文'], ['ja', '🇯🇵 日本語'], ['ko', '🇰🇷 한국어'],
        ['hi', '🇮🇳 हिन्दी'], ['bn', '🇧🇩 বাংলা'], ['ur', '🇵🇰 اردو'], ['ar', '🇸🇦 العربية']
    ];
    const RTL = new Set(['ar', 'ur']);

    let current = 'en';
    let messages = {};
    let fallback = {};

    function isSupported(locale) {
        for (let i = 0; i < LOCALES.length; i++) {
            if (LOCALES[i][0] === locale) return true;
        }
        return false;
    }

    // Browser UI language → one of ours. Taiwan/Hong Kong/Macau and "Hant"
    // get Traditional Chinese, every Portuguese gets pt_BR.
    function fromBrowser(language) {
        const lang = String(language || 'en').toLowerCase().replace('_', '-');
        if (lang.startsWith('zh')) return /-(tw|hk|mo|hant)/.test(lang) ? 'zh_TW' : 'zh_CN';
        if (lang.startsWith('pt')) return 'pt_BR';
        const base = lang.split('-')[0];
        return isSupported(base) ? base : 'en';
    }

    async function fetchMessages(locale) {
        const res = await fetch(chrome.runtime.getURL('_locales/' + locale + '/messages.json'));
        return res.json();
    }

    async function load(locale) {
        current = isSupported(locale) ? locale : 'en';
        const loaded = await Promise.all([fetchMessages(current), current === 'en' ? null : fetchMessages('en')]);
        messages = loaded[0];
        fallback = loaded[1] || loaded[0];
    }

    // chrome.i18n.getMessage() semantics: named $PLACEHOLDER$s map to $1, $2…
    function t(key, substitutions) {
        const entry = messages[key] || fallback[key];
        if (!entry) return '';
        let text = entry.message;
        if (substitutions && entry.placeholders) {
            for (const name in entry.placeholders) {
                const index = parseInt(entry.placeholders[name].content.slice(1), 10) - 1;
                if (index < 0 || index >= substitutions.length) continue;
                const value = String(substitutions[index]);
                // Function replacement: names may contain "$&" and friends.
                text = text.replace(new RegExp('\\$' + name + '\\$', 'gi'), () => value);
            }
        }
        return text;
    }

    // Translate an element and remember the key (and arguments) so apply()
    // can re-translate it after a language switch.
    function setText(el, key, substitutions) {
        el.setAttribute('data-i18n', key);
        if (substitutions) el.setAttribute('data-i18n-args', JSON.stringify(substitutions));
        else el.removeAttribute('data-i18n-args');
        el.textContent = t(key, substitutions);
    }

    function apply(root) {
        document.documentElement.lang = current.replace('_', '-');
        document.documentElement.dir = RTL.has(current) ? 'rtl' : 'ltr';
        const texts = root.querySelectorAll('[data-i18n]');
        for (let i = 0; i < texts.length; i++) {
            const args = texts[i].getAttribute('data-i18n-args');
            const text = t(texts[i].getAttribute('data-i18n'), args ? JSON.parse(args) : undefined);
            if (text) texts[i].textContent = text;
        }
        const titles = root.querySelectorAll('[data-i18n-title]');
        for (let i = 0; i < titles.length; i++) {
            const text = t(titles[i].getAttribute('data-i18n-title'));
            if (text) {
                titles[i].title = text;
                titles[i].setAttribute('aria-label', text);
            }
        }
        const placeholders = root.querySelectorAll('[data-i18n-placeholder]');
        for (let i = 0; i < placeholders.length; i++) {
            const text = t(placeholders[i].getAttribute('data-i18n-placeholder'));
            if (text) placeholders[i].placeholder = text;
        }
    }

    TE.i18n = {
        LOCALES: LOCALES,
        get locale() { return current; },
        isSupported: isSupported,
        fromBrowser: fromBrowser,
        load: load,
        t: t,
        setText: setText,
        apply: apply
    };
})(globalThis.TE = globalThis.TE || {});
