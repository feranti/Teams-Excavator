// Attachment discovery and inline-image capture.
//
// An attachment is { name, type, url?, imageId?, image? }:
//   url      remote location (SharePoint/OneDrive, asm.skype.com…), if known
//   imageId  captured bytes in the tab's ImageStore (embeddable, auth-free)
//   image    true for pictures (the HTML export notes pictures it could not embed)
// A card exposing only a file name is kept name-only so it still appears in
// the transcripts; it can't be downloaded.
(function (TE) {
    'use strict';

    const util = TE.util;
    const dom = TE.dom;

    const ASM_OBJECT_RE = /asm\.skype\.com((\.[a-z0-9-]+)*\.mcas\.ms)?\/v1\/objects\//i;
    const IMAGE_MIME_EXTENSIONS = {
        'image/png': 'png', 'image/jpeg': 'jpg', 'image/jpg': 'jpg', 'image/gif': 'gif',
        'image/webp': 'webp', 'image/bmp': 'bmp', 'image/svg+xml': 'svg', 'image/heic': 'heic',
        'image/heif': 'heif', 'image/tiff': 'tiff', 'image/avif': 'avif'
    };
    const EXCLUDED_IMAGE_HOSTS = ['giphy.com', 'static.skypeassets.com', 'statics.teams.cdn.office.net'];
    const EXCLUDED_IMAGE_SEGMENTS = new Set(['emoji', 'emoticon', 'sticker', 'avatar', 'presence', 'icon', 'icons', 'memojis']);

    // Dedup key: origin + path (query strings carry volatile tokens).
    function urlKey(url) {
        try {
            const u = new URL(url);
            return u.origin + u.pathname;
        } catch (_) {
            return url;
        }
    }

    // SharePoint/OneDrive sharing links open a viewer; download=1 returns the file.
    function directDownloadUrl(href) {
        try {
            const u = new URL(href);
            if (!u.searchParams.has('download') && isSharedFileUrl(u.hostname)) {
                u.searchParams.set('download', '1');
                return u.toString();
            }
        } catch (_) { /* not a URL */ }
        return href;
    }

    function isSharedFileUrl(href) {
        return /\.sharepoint\.com|sharepoint-df\.com|1drv\.ms/i.test(href || '');
    }

    // Add to the message's list unless already there. Returns the attachment,
    // or null for a duplicate.
    function add(session, state, key, attachment) {
        if (state.keys.has(key)) return null;
        state.keys.add(key);
        state.list.push(attachment);
        session.fileKeys.add(key);
        return attachment;
    }

    function mimeToExt(dataUrl) {
        const m = /^data:([^;,]+)/.exec(dataUrl);
        return (m && IMAGE_MIME_EXTENSIONS[m[1].toLowerCase()]) || 'png';
    }

    function withExtension(name, ext) {
        return util.fileExtension(name) ? name : name + '.' + ext;
    }

    function replaceExtension(name, ext) {
        const at = name.lastIndexOf('.');
        return (at > 0 ? name.substring(0, at) : name) + '.' + ext;
    }

    // Cheap content signature to recognise the same image across sightings.
    function signature(dataUrl) {
        let h = 17;
        for (let i = 0; i < dataUrl.length; i += 89) h = (Math.imul(h, 33) ^ dataUrl.charCodeAt(i)) | 0;
        return dataUrl.length + '/' + (h >>> 0).toString(36);
    }

    // Returns false if the message already holds these bytes.
    function claimSignature(session, messageId, sig) {
        let sigs = session.imageSignatures.get(messageId);
        if (!sigs) {
            sigs = new Set();
            session.imageSignatures.set(messageId, sigs);
        }
        if (sigs.has(sig)) return false;
        sigs.add(sig);
        return true;
    }

    // Ask the service worker to fetch an image with the extension's host
    // permissions (no page CORS) and return its exact bytes as a data: URL.
    async function fetchViaWorker(url) {
        const res = await TE.withTimeout(TE.send({ type: 'fetchImage', url: url }).catch(() => null), 20000);
        return res && res.ok ? res.dataUrl : null;
    }

    // Force an image type so the data URL embeds; browsers sniff the real
    // format when rendering and the file name gets the right extension.
    function readBlobAsUrl(blob) {
        const isImage = /^image\/[a-z0-9.+-]+$/i.test(blob.type);
        const source = isImage ? blob : new Blob([blob], { type: 'image/png' });
        return new Promise((resolve, reject) => {
            const fr = new FileReader();
            fr.addEventListener('load', () => resolve(fr.result));
            fr.addEventListener('error', () => reject(fr.error));
            fr.readAsDataURL(source);
        });
    }

    // Repaint a loaded <img> to PNG; null when the canvas is tainted.
    function rasterize(img) {
        try {
            const c = document.createElement('canvas');
            c.width = img.naturalWidth;
            c.height = img.naturalHeight;
            c.getContext('2d').drawImage(img, 0, 0);
            return c.toDataURL('image/png');
        } catch (_) {
            return null;
        }
    }

    // Upgrade a URL-based image attachment (grid screenshot, channel image) to
    // captured bytes. The original stays downloadable by URL if this fails.
    function captureUrlImage(session, messageId, attachment) {
        session.capture((async () => {
            if (session.images.full || !attachment.url.startsWith('https://')) return;
            const dataUrl = await fetchViaWorker(attachment.url);
            if (!dataUrl || !claimSignature(session, messageId, signature(dataUrl))) return;
            const ext = mimeToExt(dataUrl);
            attachment.imageId = session.images.add(dataUrl);
            // Teams may store a pasted ".png" as JPEG: trust the bytes.
            attachment.name = replaceExtension(attachment.name, ext);
            attachment.type = ext;
        })());
    }

    // Capture an inline <img>, preferring paths that copy the stored file
    // verbatim over re-encoding:
    //   1. in-page fetch — blob:/same-origin sources, exact original bytes
    //   2. service-worker fetch — cross-origin hosts we have permission for
    //   3. canvas — rendered pixels re-encoded as PNG, last resort
    // Never awaits img.decode(): an image stuck loading in Teams' virtualized
    // list would hang forever. Images not ready yet are retried next sighting.
    async function grabImage(session, messageId, img, state) {
        const src = img.currentSrc || img.src || '';
        let dataUrl = null;

        if (!session.images.full && src && src !== location.href && !src.startsWith('data:')) {
            let sameOrigin = src.startsWith('blob:');
            if (!sameOrigin) {
                try { sameOrigin = new URL(src).origin === location.origin; } catch (_) { /* keep false */ }
            }
            if (sameOrigin) {
                try {
                    const res = await fetch(src, { credentials: 'include' });
                    if (res.ok && !util.isErrorPageMime(util.responseMime(res))) {
                        const blob = await res.blob();
                        if (blob.size) dataUrl = await readBlobAsUrl(blob);
                    }
                } catch (_) { /* next path */ }
            }
            if (!dataUrl && src.startsWith('https://')) dataUrl = await fetchViaWorker(src);
        }
        if (!dataUrl && !session.images.full && img.naturalWidth && img.naturalHeight) {
            dataUrl = rasterize(img);
        }

        const label = (img.getAttribute('alt') || img.getAttribute('title') || '').trim();
        const baseName = label && !/^\p{Extended_Pictographic}/u.test(label) ? label : 'image-' + messageId + '-' + state.list.length;

        if (dataUrl) {
            const sig = signature(dataUrl);
            if (!claimSignature(session, messageId, sig)) return;
            const ext = mimeToExt(dataUrl);
            add(session, state, 'img:' + messageId + ':' + sig, {
                name: withExtension(baseName, ext),
                type: ext,
                imageId: session.images.add(dataUrl),
                image: true
            });
            return;
        }
        // No bytes: keep an https source, which may still download with the
        // user's cookies. blob: URLs mean nothing outside the page.
        if (src.startsWith('https://')) {
            const ext = util.fileExtension(baseName) || 'png';
            add(session, state, urlKey(src), { name: withExtension(baseName, ext), type: ext, url: src, image: true });
        }
    }

    // Real shared picture, or a UI glyph (emoji, avatar, icon…)?
    function isSharedPicture(img) {
        const alt = (img.alt || '').trim();
        if (alt && alt.length <= 4 && /\p{Extended_Pictographic}/u.test(alt)) return false;
        const raw = img.getAttribute('src') || img.currentSrc || '';
        if (!raw || img.src === location.href) return false;

        let url;
        try {
            url = new URL(img.src, location.href);
        } catch (_) {
            return false;
        }
        if (url.protocol !== 'https:' && url.protocol !== 'blob:' && url.protocol !== 'data:') return false;
        const host = url.hostname.toLowerCase();
        for (let i = 0; i < EXCLUDED_IMAGE_HOSTS.length; i++) {
            const h = EXCLUDED_IMAGE_HOSTS[i];
            if (host === h || host.endsWith('.' + h)) return false;
        }
        const segments = url.pathname.toLowerCase().split('/');
        for (let i = 0; i < segments.length; i++) {
            if (EXCLUDED_IMAGE_SEGMENTS.has(segments[i])) return false;
        }
        // Skip only clearly icon-sized pictures; unknown size (not loaded) passes.
        const w = img.naturalWidth, h = img.naturalHeight;
        return !(w && h && w < 32 && h < 32);
    }

    function asmUrlOf(el) {
        const direct = el.getAttribute('amspreviewurl') || el.getAttribute('amsimageurl');
        if (direct) return direct.trim();
        const attrs = el.attributes;
        for (let i = 0; i < attrs.length; i++) {
            if (ASM_OBJECT_RE.test(attrs[i].value)) return attrs[i].value.trim();
        }
        return '';
    }

    // A readable file name from the nearest aria-label, else a generated one.
    function imageName(el, messageId, state) {
        const labelled = el.getAttribute('aria-label') ? el : (el.closest('[aria-label]') || el.querySelector('[aria-label]'));
        const label = labelled ? (labelled.getAttribute('aria-label') || '').trim() : '';
        const ext = util.fileExtension(label);
        return { name: ext ? label : 'image-' + messageId + '-' + state.list.length + '.png', ext: ext || 'png' };
    }

    // Scan a message for attachments; merges into (and returns) the message's
    // persistent list. `excluded`: subtree to ignore (a channel root post's
    // reply chain, whose attachments belong to the replies).
    function scan(session, container, messageId, excluded) {
        const state = session.attachmentState(messageId);
        if (!container) return state.list;
        const outside = dom.outside;

        let grid = null;
        const grids = container.querySelectorAll('[data-tid="file-attachment-grid"]');
        for (let i = 0; i < grids.length && !grid; i++) {
            if (outside(grids[i], excluded)) grid = grids[i];
        }

        // 1. File preview cards: <button title="name\nurl">.
        const cards = container.querySelectorAll('[data-tid="file-preview-root"] button[title]');
        for (let i = 0; i < cards.length; i++) {
            if (!outside(cards[i], excluded)) continue;
            const title = cards[i].getAttribute('title') || '';
            const nl = title.indexOf('\n');
            if (nl === -1) continue;
            const name = title.slice(0, nl).trim();
            const url = title.slice(nl + 1).trim();
            if (name && url) add(session, state, urlKey(url), { name: name, type: util.fileExtension(name), url: directDownloadUrl(url) });
        }

        // 2. Pasted screenshots in the grid (teams.cloud.microsoft): the
        //    original lives in an asm.skype.com URL attribute; the only <img>
        //    is a low-res blob: thumbnail.
        const asmNodes = new Set(container.querySelectorAll('[amspreviewurl], [amsimageurl]'));
        if (grid) {
            const all = grid.getElementsByTagName('*');
            for (let i = 0; i < all.length; i++) {
                const attrs = all[i].attributes;
                for (let a = 0; a < attrs.length; a++) {
                    if (ASM_OBJECT_RE.test(attrs[a].value)) {
                        asmNodes.add(all[i]);
                        break;
                    }
                }
            }
        }
        asmNodes.forEach((node) => {
            if (!outside(node, excluded)) return;
            const url = asmUrlOf(node);
            if (!/^https?:/i.test(url)) return;
            const n = imageName(node, messageId, state);
            const added = add(session, state, urlKey(url), { name: n.name, type: n.ext, url: url, image: true });
            if (added) captureUrlImage(session, messageId, added);
        });

        if (grid) {
            // 3. Content cards (2026 layout): "name\nurl" in the aria-label of
            //    the card title; picture cards were handled above.
            const titles = grid.querySelectorAll('[data-testid="content-card-custom-title"][aria-label]');
            for (let i = 0; i < titles.length; i++) {
                const label = titles[i].getAttribute('aria-label');
                const nl = label.indexOf('\n');
                if (nl === -1) continue;
                const name = label.slice(0, nl).trim();
                const url = label.slice(nl + 1).trim();
                if (name && /^https?:/i.test(url)) add(session, state, urlKey(url), { name: name, type: util.fileExtension(name), url: directDownloadUrl(url) });
            }
            // Cards that expose a file name but no URL: name only.
            const items = grid.querySelectorAll('[role="group"][aria-label]');
            for (let i = 0; i < items.length; i++) {
                const name = items[i].getAttribute('aria-label').trim();
                if (!util.fileExtension(name) || asmUrlOf(items[i])) continue;
                if (items[i].querySelector('[data-testid="content-card-custom-title"][aria-label]')) continue;
                let known = false;
                for (let a = 0; a < state.list.length && !known; a++) known = state.list[a].name === name;
                if (!known) add(session, state, 'name:' + name, { name: name, type: util.fileExtension(name) });
            }
        }

        const body = document.getElementById('message-body-' + messageId) || container;

        // 4. Channel images (schema.skype.com/AMSImage): the visible src is a
        //    blob: thumbnail, the original sits in data-orig-src.
        const lazy = body.querySelectorAll('img[data-orig-src], img[data-gallery-src]');
        for (let i = 0; i < lazy.length; i++) {
            const img = lazy[i];
            if (!outside(img, excluded) || (grid && grid.contains(img))) continue;
            const url = (img.getAttribute('data-orig-src') || img.getAttribute('data-gallery-src') || '').trim();
            if (!/^https?:/i.test(url)) continue;
            const label = (img.getAttribute('aria-label') || img.getAttribute('alt') || '').trim();
            const ext = util.fileExtension(label) || 'png';
            const name = util.fileExtension(label) ? label : 'image-' + messageId + '-' + state.list.length + '.' + ext;
            const added = add(session, state, urlKey(url), { name: name, type: ext, url: url, image: true });
            if (added) captureUrlImage(session, messageId, added);
        }

        // 5. SharePoint/OneDrive links in the text.
        const links = container.querySelectorAll('a[href]');
        for (let i = 0; i < links.length; i++) {
            const link = links[i];
            if (!outside(link, excluded) || (grid && grid.contains(link)) || !isSharedFileUrl(link.href)) continue;
            const name = link.getAttribute('download') || link.getAttribute('title') || link.textContent.trim() ||
                'file-' + messageId + '-' + state.list.length;
            add(session, state, urlKey(link.href), { name: name, type: util.fileExtension(name), url: directDownloadUrl(link.href) });
        }

        // 6. Inline pictures in the message body. Avatars and reactions live
        //    outside the body; emoji and card thumbnails are filtered here.
        const images = body.getElementsByTagName('img');
        for (let i = 0; i < images.length; i++) {
            const img = images[i];
            if (!outside(img, excluded) || (grid && grid.contains(img))) continue;
            if (img.getAttribute('data-tid') === 'rich-file-preview-image') continue;
            if (img.hasAttribute('data-orig-src') || img.hasAttribute('data-gallery-src')) continue;
            if (img.closest('[data-tid="emoticon-renderer"]') || !isSharedPicture(img)) continue;

            // A stable https source needs one attempt; a blob: source may not
            // be ready at first sight, so it is retried on later sightings.
            const src = img.currentSrc || img.src || '';
            if (src && !src.startsWith('blob:')) {
                let tried = session.imageSourcesTried.get(messageId);
                if (!tried) {
                    tried = new Set();
                    session.imageSourcesTried.set(messageId, tried);
                }
                if (tried.has(src)) continue;
                tried.add(src);
            }
            session.capture(grabImage(session, messageId, img, state));
        }

        return state.list;
    }

    // Unique downloadable attachments of a result: [{ name, url?, imageId? }].
    function collectFiles(messages) {
        const files = [];
        const keys = new Set();
        for (let i = 0; i < messages.length; i++) {
            const attachments = messages[i].attachments;
            for (let a = 0; a < attachments.length; a++) {
                const att = attachments[a];
                const key = att.imageId ? 'img:' + att.imageId : (att.url ? urlKey(att.url) : null);
                if (!key || keys.has(key)) continue;
                keys.add(key);
                files.push(att.imageId ? { name: att.name, imageId: att.imageId } : { name: att.name, url: att.url });
            }
        }
        return files;
    }

    TE.attachments = { scan: scan, collectFiles: collectFiles };
})(globalThis.TE = globalThis.TE || {});
