import CFB from "cfb";

const TAG_PARA_TEXT = 0x0043;
const MAX_DECOMPRESS = 100 * 1024 * 1024;

async function tryDecompress(data, format) {
  const ds = new DecompressionStream(format);
  const writer = ds.writable.getWriter();
  const reader = ds.readable.getReader();

  // Enqueue data without awaiting — read loop runs concurrently
  writer.write(data).catch(() => {});
  writer.close().catch(() => {});

  const chunks = [];
  let total = 0;
  while (total < MAX_DECOMPRESS) {
    try {
      const { done, value } = await reader.read();
      if (done) break;
      chunks.push(value);
      total += value.length;
    } catch {
      // "Junk found after end of compressed data" — browser is strict,
      // Node.js zlib is lenient. Accept what was successfully decompressed.
      break;
    }
  }
  if (chunks.length === 0) return null;
  const out = new Uint8Array(total);
  let off = 0;
  for (const c of chunks) { out.set(c, off); off += c.length; }
  return out;
}

async function decompress(data) {
  // HWP5 uses zlib (0x78 header) or raw deflate
  const isZlib = data.length >= 2 && data[0] === 0x78;
  const result = await tryDecompress(data, isZlib ? "deflate" : "deflate-raw");
  if (result) return result;
  // Fallback to the other format
  const fallback = await tryDecompress(data, isZlib ? "deflate-raw" : "deflate");
  return fallback ?? data;
}

function parseRecords(data) {
  const view = new DataView(data.buffer, data.byteOffset, data.byteLength);
  const records = [];
  let offset = 0;
  while (offset + 4 <= data.length) {
    const hdr = view.getUint32(offset, true);
    offset += 4;
    const tagId = hdr & 0x3ff;
    let size = (hdr >> 20) & 0xfff;
    if (size === 0xfff) {
      if (offset + 4 > data.length) break;
      size = view.getUint32(offset, true);
      offset += 4;
    }
    if (offset + size > data.length) break;
    records.push({ tagId, data: data.subarray(offset, offset + size) });
    offset += size;
  }
  return records;
}

function utf16LeToString(data) {
  const bytes = data.length % 2 === 0 ? data : data.slice(0, -1);
  const decoded = new TextDecoder("utf-16le").decode(bytes);
  let result = "";
  for (const ch of decoded) {
    const cp = ch.codePointAt(0);
    // Skip HWP inline control codes (0x00-0x1F)
    if (cp < 0x20) continue;
    // Skip HWP private-use object markers (0xE000-0xF8FF)
    if (cp >= 0xe000 && cp <= 0xf8ff) continue;
    result += ch;
  }
  return result;
}

function findEntryByPath(cfb, pattern) {
  for (let i = 0; i < cfb.FullPaths.length; i++) {
    if (pattern.test(cfb.FullPaths[i])) return cfb.FileIndex[i];
  }
  return null;
}

function findSections(cfb) {
  const found = [];
  for (let i = 0; i < cfb.FullPaths.length; i++) {
    const m = cfb.FullPaths[i].match(/Section(\d+)$/i);
    if (m && /BodyText/i.test(cfb.FullPaths[i])) {
      found.push({ idx: parseInt(m[1]), entry: cfb.FileIndex[i] });
    }
  }
  found.sort((a, b) => a.idx - b.idx);
  return found.map((s) => s.entry);
}

function isCompressed(cfb) {
  const entry = findEntryByPath(cfb, /FileHeader$/i);
  if (!entry?.content) return true;
  const h = new Uint8Array(entry.content);
  if (h.length < 40) return true;
  const flags = new DataView(h.buffer, h.byteOffset).getUint32(36, true);
  return (flags & 0x01) !== 0;
}

export async function parseHwp(file) {
  const buf = await file.arrayBuffer();
  let cfb;
  try {
    cfb = CFB.parse(new Uint8Array(buf));
  } catch (e) {
    throw new Error(`HWP CFB 파싱 실패: ${e.message}`);
  }

  console.log("[hwpParser] FullPaths:", cfb.FullPaths);

  const compressed = isCompressed(cfb);
  const sections = findSections(cfb);

  console.log("[hwpParser] compressed:", compressed, "/ sections:", sections.length);

  const chunks = [];
  let paraNum = 0;

  for (const entry of sections) {
    let data = new Uint8Array(entry.content);
    if (compressed) data = await decompress(data);
    const records = parseRecords(data);
    const paraTexts = records.filter((r) => r.tagId === TAG_PARA_TEXT);
    console.log(`[hwpParser] records: ${records.length}, PARA_TEXT: ${paraTexts.length}`);
    for (const { data: recData } of paraTexts) {
      const text = utf16LeToString(recData).trim();
      if (text) {
        paraNum++;
        chunks.push({ text, location: `${paraNum}번째 단락` });
      }
    }
  }

  console.log("[hwpParser] total chunks:", chunks.length);
  if (chunks.length > 0) {
    console.log("[hwpParser] first 3:", chunks.slice(0, 3).map((c) => c.text));
  }

  return chunks;
}