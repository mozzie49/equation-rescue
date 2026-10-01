import { Inflate, deflateSync } from "fflate";
export const LIMITS = {
  file: 20 * 1024 * 1024,
  entries: 2048,
  total: 64 * 1024 * 1024,
  entry: 16 * 1024 * 1024,
  xml: 8 * 1024 * 1024,
  ratio: 200,
};
export type ZipEntry = {
  name: string;
  flags: number;
  method: number;
  crc: number;
  size: number;
  packed: number;
  local: number;
  data: number;
  end: number;
  central: Uint8Array;
  raw: Uint8Array;
};
const decoder = new TextDecoder("utf-8", { fatal: true }),
  encoder = new TextEncoder();
const table = Uint32Array.from({ length: 256 }, (_, n) => {
  let c = n;
  for (let k = 0; k < 8; k++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
  return c >>> 0;
});
export function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (const b of data) c = table[(c ^ b) & 255] ^ (c >>> 8);
  return (c ^ 0xffffffff) >>> 0;
}
const bad = (s: string): never => {
  throw new Error(s);
};
export class SafeZip {
  entries: ZipEntry[] = [];
  private view: DataView;
  constructor(readonly bytes: Uint8Array) {
    if (bytes.length > LIMITS.file) bad("File exceeds the 20 MB limit");
    this.view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    const u16 = (p: number) => this.view.getUint16(p, true),
      u32 = (p: number) => this.view.getUint32(p, true);
    let end = -1;
    for (let p = bytes.length - 22; p >= Math.max(0, bytes.length - 65557); p--)
      if (u32(p) === 0x06054b50 && p + 22 + u16(p + 20) === bytes.length) {
        end = p;
        break;
      }
    if (end < 0)
      bad("Not a standard ZIP-based DOCX (encrypted files are unsupported)");
    if (u16(end + 4) || u16(end + 6) || u16(end + 8) !== u16(end + 10))
      bad("Multi-disk ZIP files are unsupported");
    const count = u16(end + 10),
      centralSize = u32(end + 12),
      offset = u32(end + 16);
    if (count > LIMITS.entries || count === 0xffff || offset === 0xffffffff)
      bad("ZIP entry limit or ZIP64 format is unsupported");
    if (offset + centralSize !== end) bad("Invalid ZIP directory bounds");
    let p = offset,
      total = 0;
    const names = new Set<string>();
    for (let i = 0; i < count; i++) {
      if (p + 46 > end || u32(p) !== 0x02014b50) bad("Invalid ZIP directory");
      const flags = u16(p + 8),
        method = u16(p + 10),
        crc = u32(p + 16),
        packed = u32(p + 20),
        size = u32(p + 24),
        nl = u16(p + 28),
        el = u16(p + 30),
        cl = u16(p + 32),
        local = u32(p + 42);
      const stop = p + 46 + nl + el + cl;
      if (stop > end) bad("Truncated ZIP entry");
      let name: string;
      try {
        name = decoder.decode(bytes.subarray(p + 46, p + 46 + nl));
      } catch {
        bad("ZIP filenames must be valid UTF-8");
      }
      if (
        !name! ||
        name!.startsWith("/") ||
        name!.includes("\\") ||
        name!.includes(":") ||
        name!.split("/").some((x) => x === ".." || x === ".") ||
        /[\x00-\x1f]/.test(name!)
      )
        bad("Unsafe ZIP path");
      if (names.has(name!.toLowerCase()))
        bad("Duplicate or ambiguous ZIP path");
      names.add(name!.toLowerCase());
      if (flags & ~0x080e || flags & 1 || (method !== 0 && method !== 8))
        bad("Encrypted or unsupported ZIP entry");
      if (u16(p + 34) !== 0) bad("Multi-disk ZIP entry");
      total += size;
      if (
        size > LIMITS.entry ||
        total > LIMITS.total ||
        size > Math.max(1024 * 1024, packed * LIMITS.ratio)
      )
        bad("Expanded ZIP size or compression ratio exceeds safety limits");
      if (local + 30 > offset || u32(local) !== 0x04034b50)
        bad("Invalid ZIP local header");
      if (u16(local + 6) !== flags || u16(local + 8) !== method)
        bad("Conflicting ZIP headers");
      const ln = u16(local + 26),
        le = u16(local + 28),
        data = local + 30 + ln + le;
      if (
        data + packed > offset ||
        decoder.decode(bytes.subarray(local + 30, local + 30 + ln)) !== name!
      )
        bad("Invalid ZIP entry data");
      if (
        !(flags & 8) &&
        (u32(local + 14) !== crc ||
          u32(local + 18) !== packed ||
          u32(local + 22) !== size)
      )
        bad("Conflicting ZIP sizes");
      let localEnd = data + packed;
      if (flags & 8) {
        let dp = localEnd;
        if (dp + 4 <= offset && u32(dp) === 0x08074b50) dp += 4;
        if (
          dp + 12 > offset ||
          u32(dp) !== crc ||
          u32(dp + 4) !== packed ||
          u32(dp + 8) !== size
        )
          bad("Invalid ZIP data descriptor");
        localEnd = dp + 12;
      }
      this.entries.push({
        name: name!,
        flags,
        method,
        crc,
        packed,
        size,
        local,
        data,
        end: localEnd,
        central: bytes.slice(p, stop),
        raw: bytes.slice(local, localEnd),
      });
      p = stop;
    }
    if (p !== end) bad("ZIP directory size mismatch");
    const ordered = [...this.entries].sort((a, b) => a.local - b.local);
    let last = 0;
    for (const e of ordered) {
      if (e.local !== last) bad("Overlapping or unaccounted ZIP data");
      last = e.end;
    }
    if (last !== offset) bad("Unaccounted ZIP data");
  }
  read(name: string): Uint8Array {
    const e = this.entries.find((e) => e.name === name);
    if (!e) bad(`Missing DOCX part: ${name}`);
    const packed = this.bytes.subarray(e!.data, e!.data + e!.packed);
    // Small compressed chunks bound allocation before a lying size can expand a bomb.
    let out: Uint8Array;
    if (e!.method === 0) out = packed.slice();
    else {
      out = new Uint8Array(e!.size);
      let written = 0;
      const inflater = new Inflate((chunk) => {
        if (written + chunk.length > e!.size)
          bad("DEFLATE data exceeds declared size");
        out.set(chunk, written);
        written += chunk.length;
      });
      if (!packed.length) bad("Empty DEFLATE stream");
      for (let p = 0; p < packed.length; p += 1024)
        inflater.push(packed.subarray(p, p + 1024), p + 1024 >= packed.length);
      if (written !== e!.size) bad("DEFLATE size mismatch");
    }
    if (out.length !== e!.size || crc32(out) !== e!.crc)
      bad("ZIP size or checksum mismatch");
    return out;
  }
  replace(name: string, data: Uint8Array): Uint8Array {
    if (data.length > LIMITS.entry) bad("Output part exceeds safety limit");
    const localParts: Uint8Array[] = [],
      centralParts: Uint8Array[] = [];
    let offset = 0;
    for (const e of this.entries) {
      let raw = e.raw,
        central = e.central.slice();
      if (e.name === name) {
        const filename = encoder.encode(e.name),
          compressed = deflateSync(data, { level: 6 }),
          crc = crc32(data);
        raw = new Uint8Array(30 + filename.length + compressed.length);
        const v = new DataView(raw.buffer);
        v.setUint32(0, 0x04034b50, true);
        v.setUint16(4, 20, true);
        v.setUint16(6, 0x800, true);
        v.setUint16(8, 8, true);
        v.setUint32(14, crc, true);
        v.setUint32(18, compressed.length, true);
        v.setUint32(22, data.length, true);
        v.setUint16(26, filename.length, true);
        raw.set(filename, 30);
        raw.set(compressed, 30 + filename.length);
        central = new Uint8Array(46 + filename.length);
        const c = new DataView(central.buffer);
        c.setUint32(0, 0x02014b50, true);
        c.setUint16(4, 20, true);
        c.setUint16(6, 20, true);
        c.setUint16(8, 0x800, true);
        c.setUint16(10, 8, true);
        c.setUint32(16, crc, true);
        c.setUint32(20, compressed.length, true);
        c.setUint32(24, data.length, true);
        c.setUint16(28, filename.length, true);
        central.set(filename, 46);
      }
      new DataView(
        central.buffer,
        central.byteOffset,
        central.byteLength,
      ).setUint32(42, offset, true);
      localParts.push(raw);
      centralParts.push(central);
      offset += raw.length;
    }
    const centralSize = centralParts.reduce((n, b) => n + b.length, 0),
      end = new Uint8Array(22),
      v = new DataView(end.buffer);
    v.setUint32(0, 0x06054b50, true);
    v.setUint16(8, this.entries.length, true);
    v.setUint16(10, this.entries.length, true);
    v.setUint32(12, centralSize, true);
    v.setUint32(16, offset, true);
    const all = [...localParts, ...centralParts, end],
      out = new Uint8Array(offset + centralSize + 22);
    let p = 0;
    for (const b of all) {
      out.set(b, p);
      p += b.length;
    }
    return out;
  }
}
