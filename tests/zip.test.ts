import { describe, it, expect } from "vitest";
import { zipSync, strToU8 } from "fflate";
import { SafeZip, crc32, LIMITS } from "../src/core/zip";
import { parseXml } from "../src/core/xml";
describe("hostile ZIP/XML boundaries", () => {
  it("reads stored and deflated members with CRC checks", () => {
    for (const level of [0, 6] as const) {
      const data = strToU8("abc".repeat(500)),
        zip = new SafeZip(zipSync({ "a.txt": data }, { level }));
      expect(zip.read("a.txt")).toEqual(data);
    }
  });
  it.each(["../evil", "/absolute", "a/../../bad", "a\\bad", "C:bad", "a/./b"])(
    "rejects unsafe %s",
    (name) =>
      expect(() => new SafeZip(zipSync({ [name]: strToU8("x") }))).toThrow(
        /Unsafe/,
      ),
  );
  it("rejects duplicate case-folded names", () =>
    expect(
      () =>
        new SafeZip(
          zipSync({ "Word/A": strToU8("x"), "word/a": strToU8("y") }),
        ),
    ).toThrow(/ambiguous/));
  it("rejects excessive archive size", () =>
    expect(() => new SafeZip(new Uint8Array(LIMITS.file + 1))).toThrow(
      /20 MB/,
    ));
  it("rejects too many entries", () => {
    const entries = Object.fromEntries(
      Array.from({ length: 2049 }, (_, i) => [`${i}`, new Uint8Array(0)]),
    );
    expect(() => new SafeZip(zipSync(entries))).toThrow(/entry limit/);
  });
  it("rejects metadata-based expansion bombs", () => {
    const bytes = zipSync({ a: new Uint8Array(2 * 1024 * 1024) });
    expect(() => new SafeZip(bytes)).toThrow(/compression ratio/);
  });
  it("rejects deflate output larger than lying header even with forged prefix CRC", () => {
    const bytes = zipSync({ a: strToU8("A".repeat(10000)) }),
      v = new DataView(bytes.buffer),
      central = bytes.findIndex(
        (_, i) => i + 4 < bytes.length && v.getUint32(i, true) === 0x02014b50,
      ),
      crc = crc32(strToU8("A"));
    v.setUint32(14, crc, true);
    v.setUint32(22, 1, true);
    v.setUint32(central + 16, crc, true);
    v.setUint32(central + 24, 1, true);
    expect(() => new SafeZip(bytes).read("a")).toThrow(/exceeds declared/);
  });
  it("rejects corrupted data and conflicting local headers", () => {
    const bytes = zipSync({ a: strToU8("abcdef") }, { level: 0 }),
      zip = new SafeZip(bytes);
    bytes[zip.entries[0].data] ^= 1;
    expect(() => new SafeZip(bytes).read("a")).toThrow(/checksum/);
    const another = zipSync({ a: strToU8("abc") });
    new DataView(another.buffer).setUint32(22, 2, true);
    expect(() => new SafeZip(another)).toThrow(/Conflicting/);
  });
  it("rejects XML entities, excessive depth, invalid namespace and malformed XML", () => {
    expect(() => parseXml('<!DOCTYPE r [<!ENTITY x "x">]><r>&x;</r>')).toThrow(
      /DTD/,
    );
    expect(() => parseXml("<a>".repeat(130) + "</a>".repeat(130))).toThrow(
      /complexity/,
    );
    expect(() => parseXml("<a:b/>")).toThrow();
    expect(() => parseXml("<a></b>")).toThrow();
  });
});
