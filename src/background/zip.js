// ZIP writer (PKWARE APPNOTE 6.3), used by the service worker.
//
// Two kinds of entries:
//   - addFile(): bytes known up front. Stored with sizes in the local header,
//     no data descriptor, no ZIP64 — the most portable layout. Used for the
//     generated transcripts.
//   - openStreamEntry(): size unknown while streaming. General-purpose bit 3 (data
//     descriptor) plus a ZIP64 extra field in the local header, so the trailing
//     descriptor carries 8-byte sizes and entries can exceed 4 GB. Used for
//     attachments fetched straight into a streamed download.
// The central directory only gets ZIP64 fields for values that overflow, and
// the ZIP64 end records are only written when the archive needs them.
(function (TE) {
    'use strict';

    const CRC_TABLE = new Uint32Array(256);
    for (let i = 0; i < 256; i++) {
        let c = i;
        for (let k = 0; k < 8; k++) c = (c & 1) ? (0xedb88320 ^ (c >>> 1)) : (c >>> 1);
        CRC_TABLE[i] = c >>> 0;
    }

    // Incremental CRC-32: crc32(crc32(0, a), b) === crc32(0, a ++ b).
    function crc32(crc, bytes) {
        let c = ~crc;
        for (let i = 0; i < bytes.length; i++) c = CRC_TABLE[(c ^ bytes[i]) & 0xff] ^ (c >>> 8);
        return ~c >>> 0;
    }

    const SIG_LOCAL = 0x04034b50;
    const SIG_DESCRIPTOR = 0x08074b50;
    const SIG_CENTRAL = 0x02014b50;
    const SIG_END64 = 0x06064b50;
    const SIG_LOCATOR64 = 0x07064b50;
    const SIG_END = 0x06054b50;

    const FLAG_UTF8 = 0x0800;
    const FLAG_DESCRIPTOR = 0x0008;
    const VERSION_DEFAULT = 20;
    const VERSION_ZIP64 = 45;
    const U16_MAX = 0xffff;
    const U32_MAX = 0xffffffff;
    const TWO_32 = 0x100000000;

    const encoder = new TextEncoder();

    // Little-endian writer over a fixed-size buffer.
    class Out {
        constructor(size) {
            this.bytes = new Uint8Array(size);
            this.view = new DataView(this.bytes.buffer);
            this.pos = 0;
        }
        u16(n) { this.view.setUint16(this.pos, n, true); this.pos += 2; return this; }
        u32(n) { this.view.setUint32(this.pos, n >>> 0, true); this.pos += 4; return this; }
        // 64-bit from a Number (exact up to 2^53) without BigInt allocations.
        u64(n) {
            this.view.setUint32(this.pos, n % TWO_32, true);
            this.view.setUint32(this.pos + 4, Math.floor(n / TWO_32), true);
            this.pos += 8;
            return this;
        }
        raw(b) { this.bytes.set(b, this.pos); this.pos += b.length; return this; }
    }

    function toDosStamp(d) {
        return {
            time: (d.getHours() << 11) | (d.getMinutes() << 5) | (d.getSeconds() >> 1),
            date: ((Math.max(d.getFullYear(), 1980) - 1980) << 9) | ((d.getMonth() + 1) << 5) | d.getDate()
        };
    }

    // An entry being streamed. Store: call consume() and produce() with the
    // same chunk. Deflate: consume() each uncompressed chunk (CRC + size) and
    // produce() each compressed chunk written downstream.
    class StreamedEntry {
        constructor(writer, name, method, date) {
            this.writer = writer;
            this.record = {
                nameBytes: encoder.encode(name),
                method: method,
                flags: FLAG_UTF8 | FLAG_DESCRIPTOR,
                version: VERSION_ZIP64,
                dt: toDosStamp(date),
                offset: writer.offset,
                crc: 0,
                size: 0,
                compSize: 0
            };
            const r = this.record;
            this.header = new Out(30 + r.nameBytes.length + 20)
                .u32(SIG_LOCAL).u16(VERSION_ZIP64).u16(r.flags).u16(method)
                .u16(r.dt.time).u16(r.dt.date)
                .u32(0).u32(U32_MAX).u32(U32_MAX)            // CRC and sizes live in the descriptor
                .u16(r.nameBytes.length).u16(20)
                .raw(r.nameBytes)
                .u16(0x0001).u16(16).u64(0).u64(0)           // ZIP64 extra: sizes follow in descriptor
                .bytes;
            writer.offset += this.header.length;
            this.closed = false;
        }

        consume(bytes) {
            this.record.crc = crc32(this.record.crc, bytes);
            this.record.size += bytes.length;
        }

        produce(length) {
            this.record.compSize += length;
            this.writer.offset += length;
        }

        // Drop an entry whose header was never written (its source failed
        // before producing any data).
        abort() {
            if (this.closed) throw new Error('ZIP entry already closed');
            if (this.record.compSize) throw new Error('Cannot abort a ZIP entry after writing data');
            this.closed = true;
            this.writer.offset = this.record.offset;
            this.writer.entryOpen = false;
        }

        // Returns the data descriptor to write after the entry's data.
        close() {
            if (this.closed) throw new Error('ZIP entry closed twice');
            this.closed = true;
            const r = this.record;
            const descriptor = new Out(24)
                .u32(SIG_DESCRIPTOR).u32(r.crc).u64(r.compSize).u64(r.size)
                .bytes;
            this.writer.offset += descriptor.length;
            this.writer.entries.push(r);
            this.writer.entryOpen = false;
            return descriptor;
        }
    }

    class ZipWriter {
        constructor() {
            this.entries = [];
            this.offset = 0;
            this.entryOpen = false;
        }

        assertIdle() {
            if (this.entryOpen) throw new Error('A ZIP entry is still open');
        }

        // Returns [localHeader, data]: the chunks to write, in order.
        addFile(name, data, date) {
            this.assertIdle();
            const nameBytes = encoder.encode(name);
            const r = {
                nameBytes: nameBytes,
                method: 0,
                flags: FLAG_UTF8,
                version: VERSION_DEFAULT,
                dt: toDosStamp(date || new Date()),
                offset: this.offset,
                crc: crc32(0, data),
                size: data.length,
                compSize: data.length
            };
            const header = new Out(30 + nameBytes.length)
                .u32(SIG_LOCAL).u16(VERSION_DEFAULT).u16(r.flags).u16(0)
                .u16(r.dt.time).u16(r.dt.date)
                .u32(r.crc).u32(r.size).u32(r.size)
                .u16(nameBytes.length).u16(0)
                .raw(nameBytes)
                .bytes;
            this.entries.push(r);
            this.offset += header.length + data.length;
            return [header, data];
        }

        // deflate: true for method 8 (raw DEFLATE), false to store.
        openStreamEntry(name, deflate, date) {
            this.assertIdle();
            this.entryOpen = true;
            return new StreamedEntry(this, name, deflate ? 8 : 0, date || new Date());
        }

        // Central directory + end records.
        finish() {
            this.assertIdle();
            const entries = this.entries;
            const cdOffset = this.offset;

            let cdSize = 0;
            for (let i = 0; i < entries.length; i++) cdSize += 46 + entries[i].nameBytes.length + zip64ExtraLength(entries[i]);

            const needZip64 = entries.length >= U16_MAX || cdOffset >= U32_MAX || cdSize >= U32_MAX;
            const out = new Out(cdSize + (needZip64 ? 56 + 20 : 0) + 22);

            for (let i = 0; i < entries.length; i++) {
                const e = entries[i];
                const extra = zip64ExtraLength(e);
                out.u32(SIG_CENTRAL).u16(e.version).u16(e.version).u16(e.flags).u16(e.method)
                    .u16(e.dt.time).u16(e.dt.date).u32(e.crc)
                    .u32(e.compSize >= U32_MAX ? U32_MAX : e.compSize)
                    .u32(e.size >= U32_MAX ? U32_MAX : e.size)
                    .u16(e.nameBytes.length).u16(extra)
                    .u16(0).u16(0).u16(0).u32(0)             // comment, disk, internal/external attrs
                    .u32(e.offset >= U32_MAX ? U32_MAX : e.offset)
                    .raw(e.nameBytes);
                if (extra) {
                    // APPNOTE 4.5.3: only overflowing fields, in this order.
                    out.u16(0x0001).u16(extra - 4);
                    if (e.size >= U32_MAX) out.u64(e.size);
                    if (e.compSize >= U32_MAX) out.u64(e.compSize);
                    if (e.offset >= U32_MAX) out.u64(e.offset);
                }
            }

            const count = entries.length;
            if (needZip64) {
                const end64Offset = cdOffset + cdSize;
                out.u32(SIG_END64).u64(44).u16(VERSION_ZIP64).u16(VERSION_ZIP64).u32(0).u32(0)
                    .u64(count).u64(count).u64(cdSize).u64(cdOffset);
                out.u32(SIG_LOCATOR64).u32(0).u64(end64Offset).u32(1);
            }
            out.u32(SIG_END).u16(0).u16(0)
                .u16(count >= U16_MAX ? U16_MAX : count)
                .u16(count >= U16_MAX ? U16_MAX : count)
                .u32(cdSize >= U32_MAX ? U32_MAX : cdSize)
                .u32(cdOffset >= U32_MAX ? U32_MAX : cdOffset)
                .u16(0);

            this.offset += out.bytes.length;
            return out.bytes;
        }
    }

    function zip64ExtraLength(e) {
        let n = 0;
        if (e.size >= U32_MAX) n += 8;
        if (e.compSize >= U32_MAX) n += 8;
        if (e.offset >= U32_MAX) n += 8;
        return n ? n + 4 : 0;
    }

    // In-memory archive from [{ path, data: string | Uint8Array }]. Returns a
    // Blob so the chunks are never copied into one contiguous buffer.
    function zipToBlob(files) {
        const writer = new ZipWriter();
        const parts = [];
        const now = new Date();
        for (let i = 0; i < files.length; i++) {
            const f = files[i];
            const data = typeof f.data === 'string' ? encoder.encode(f.data) : f.data;
            const chunks = writer.addFile(f.path, data, now);
            parts.push(chunks[0], chunks[1]);
        }
        parts.push(writer.finish());
        return new Blob(parts, { type: 'application/zip' });
    }

    TE.zip = { ZipWriter: ZipWriter, zipToBlob: zipToBlob };
})(globalThis.TE = globalThis.TE || {});
