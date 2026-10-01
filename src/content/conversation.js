// Who is who: the signed-in user, the conversation's display name, and the
// chats/channels listed in the Teams sidebar.
(function (TE) {
    'use strict';

    // Best effort: Teams exposes the signed-in user's name in different places
    // depending on the build.
    function currentUserName() {
        try {
            const w = window;
            const candidates = [
                w.__TEAMS_STATE__ && (w.__TEAMS_STATE__.userName || (w.__TEAMS_STATE__.user && w.__TEAMS_STATE__.user.displayName)),
                w._currentUser && (w._currentUser.displayName || w._currentUser.name),
                w.me && (w.me.displayName || w.me.name),
                w.currentUser && (w.currentUser.displayName || w.currentUser.name)
            ];
            for (let i = 0; i < candidates.length; i++) {
                if (typeof candidates[i] === 'string' && candidates[i].length > 1) return candidates[i];
            }
        } catch (_) { /* page globals are not always reachable */ }

        const selectors = [
            '[data-tid="me"]', '[data-tid*="profile"]', '[aria-label*="My profile"]',
            '[data-tid="header-profile-button"]', '[class*="user-name"]', '[class*="displayName"]',
            '[data-tid="persona-name"]', '[aria-label*="account"]'
        ];
        for (let i = 0; i < selectors.length; i++) {
            const el = document.querySelector(selectors[i]);
            const text = el ? (el.textContent || el.getAttribute('aria-label') || '').trim() : '';
            if (text.length > 1) return text;
        }
        const header = document.querySelector('.fui-HubHeader, [data-tid="header"]');
        if (header) {
            const names = header.querySelectorAll('[class*="name"], [class*="display"], [class*="user"]');
            for (let i = 0; i < names.length; i++) {
                const t = (names[i].textContent || '').trim();
                if (t.length > 2 && t.indexOf(' ') !== -1) return t;
            }
        }
        return '';
    }

    // Channel tab titles read "… | <Channel> | <Team> | Microsoft Teams": the
    // app name is last, team and channel come right before it. Nothing on the
    // channel page itself shows the team name.
    function channelNameFromTitle() {
        const parts = [];
        const raw = document.title.split('|');
        for (let i = 0; i < raw.length; i++) {
            const p = raw[i].trim();
            if (p) parts.push(p);
        }
        if (parts.length < 3) return '';
        return parts[parts.length - 2] + ' - ' + parts[parts.length - 3];
    }

    function readChatHeader() {
        let name = '';
        let isThread = false;
        const participants = [];
        const title = document.querySelector('[data-tid="chat-title"]');
        if (title) {
            // Group and meeting chats are "…@thread.v2"; 1:1 chats are
            // "…@unq.gbl.spaces" and only ever show the other person's name.
            const container = title.closest('[id^="chat-header-"]');
            isThread = !!container && container.id.indexOf('@thread.') !== -1;
            const span = title.querySelector('span[title], span[dir="auto"]');
            if (span) name = (span.getAttribute('title') || span.textContent || '').trim();
            // Unnamed group chats list every participant in the header menu,
            // even when the visible title shows only the first one.
            const menu = title.querySelector('[data-tid="chat-topic-menu"]');
            if (menu) {
                const people = menu.querySelectorAll('[id^="chat-topic-person-"]');
                for (let i = 0; i < people.length; i++) {
                    const t = people[i].textContent.trim();
                    if (t && participants.indexOf(t) === -1) participants.push(t);
                }
                if (!name) {
                    const items = menu.querySelectorAll(':scope > [data-tid^="participant-"]');
                    const only = items.length === 1 ? items[0].querySelector('[id^="chat-topic-person-"]') : null;
                    if (only) name = only.textContent.trim();
                }
            }
        }
        if (!name) {
            const selectors = ['[data-tid="channelTitle-text"]', '[data-tid="chat-header-topic"]',
                '[data-tid="channel-header-topic"]', '[data-tid="topic"]'];
            for (let i = 0; i < selectors.length && !name; i++) {
                const el = document.querySelector(selectors[i]);
                if (el) name = el.textContent.trim();
            }
        }
        return { name: name, participants: participants, isThread: isThread };
    }

    function looselyMatches(a, b) {
        const x = a.toLowerCase(), y = b.toLowerCase();
        return x === y || x.indexOf(y) !== -1 || y.indexOf(x) !== -1;
    }

    // Display name of the open conversation. Channels: "Team - Channel".
    // Chats: the header, cross-checked against the message authors — the header
    // may be a localized "Unknown user", a first name only, or one participant
    // of an unnamed group chat.
    function name(session, messages) {
        if (TE.channel.onChannelSurface()) {
            const fromTitle = channelNameFromTitle();
            if (fromTitle) return fromTitle;
        }
        const header = readChatHeader();
        const unknown = session.config.unknownAuthor;
        if (unknown && header.name.toLowerCase() === unknown) header.name = '';
        if (!messages.length) return header.name;

        const me = session.userName;
        const others = [];
        for (let i = 0; i < messages.length; i++) {
            const a = messages[i].author;
            if (a && a !== me && others.indexOf(a) === -1) others.push(a);
        }

        // Unnamed group chat: title it with everyone but the current user,
        // preferring full author names over the menu's first names.
        if (header.participants.length > 1 || (header.participants.length === 1 && others.length > 1)) {
            const names = [];
            for (let i = 0; i < header.participants.length; i++) {
                const p = header.participants[i];
                if (p === me) continue;
                let full = p;
                for (let j = 0; j < others.length; j++) {
                    if (others[j].toLowerCase().indexOf(p.toLowerCase()) !== -1) {
                        full = others[j];
                        break;
                    }
                }
                if (names.indexOf(full) === -1) names.push(full);
            }
            for (let i = 0; i < others.length; i++) {
                if (names.indexOf(others[i]) === -1) names.push(others[i]);
            }
            if (names.length > 1) return names.join(', ');
        }

        let headerIsParticipant = false;
        for (let i = 0; header.name && i < others.length && !headerIsParticipant; i++) {
            headerIsParticipant = looselyMatches(others[i], header.name);
        }
        if (others.length > 1 && (!header.name || headerIsParticipant)) return others.join(', ');
        // A named group or meeting chat where only one other person wrote.
        if (header.isThread && header.name && !headerIsParticipant) return header.name;
        // One other author: a 1:1 chat. Trust the author over the header.
        if (others.length === 1) return others[0];
        return header.name || (others.length ? others[0] : '');
    }

    // Chats and channels of the sidebar: [{ id, name, type }].
    function listSidebar() {
        const out = [];
        const chats = document.querySelectorAll('[data-item-type="chat"]');
        for (let i = 0; i < chats.length; i++) {
            const id = chats[i].getAttribute('data-fui-tree-item-value');
            if (!id) continue;
            const title = chats[i].querySelector('[id^="title-chat-list-item_"]');
            out.push({ id: id, name: title ? title.textContent.trim() : '', type: 'chat' });
        }
        const channels = document.querySelectorAll('[data-item-type="channel"]');
        for (let i = 0; i < channels.length; i++) {
            const id = channels[i].getAttribute('data-fui-tree-item-value');
            if (!id) continue;
            const title = channels[i].querySelector('[id^="title-channel-list-item-"]');
            let label = title ? title.textContent.trim() : '';
            // The channel's group is a child of the team's own tree item.
            const group = channels[i].closest('[role="group"]');
            const team = group ? group.parentElement : null;
            if (team && team.getAttribute('data-item-type') === 'team') {
                const teamTitle = team.querySelector('[id^="title-team-list-item-"]');
                if (teamTitle) label = teamTitle.textContent.trim() + ' > ' + label;
            }
            out.push({ id: id, name: label, type: 'channel' });
        }
        return out;
    }

    function sidebarItem(id) {
        return document.querySelector('[data-fui-tree-item-value="' + CSS.escape(id) + '"]');
    }

    // Open a sidebar chat/channel. Returns false if it's not in the sidebar.
    function openSidebarItem(id) {
        const item = sidebarItem(id);
        if (!item) return false;
        const target = item.querySelector('[data-inp="simple-collab-channel-switch"]') ||
            item.querySelector('[data-inp="simple-collab-chat-switch"]') || item;
        target.click();
        return true;
    }

    function isSidebarItemSelected(id) {
        const item = sidebarItem(id);
        if (!item) return false;
        const current = item.getAttribute('aria-current');
        return item.getAttribute('aria-selected') === 'true' || (!!current && current !== 'false');
    }

    // What identifies the open conversation: its header, the tab title and the
    // first rendered message. Compared by hasSwitched() after a sidebar click.
    function openState() {
        const header = document.querySelector('[id^="chat-header-"]');
        const body = document.querySelector(TE.dom.BODY_SELECTOR);
        return { key: (header ? header.id : '') + '|' + document.title, firstBody: body ? body.id : '' };
    }

    // Has Teams replaced the conversation captured in `before`? The header
    // switches first; the old messages must be gone too.
    function hasSwitched(before) {
        const header = document.querySelector('[id^="chat-header-"]');
        if ((header ? header.id : '') + '|' + document.title === before.key) return false;
        const body = document.querySelector(TE.dom.BODY_SELECTOR);
        return !body || !before.firstBody || body.id !== before.firstBody;
    }

    TE.conversation = {
        currentUserName: currentUserName,
        name: name,
        listSidebar: listSidebar,
        openSidebarItem: openSidebarItem,
        isSidebarItemSelected: isSidebarItemSelected,
        openState: openState,
        hasSwitched: hasSwitched
    };
})(globalThis.TE = globalThis.TE || {});
