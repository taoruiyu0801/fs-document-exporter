const textEncoder = new TextEncoder();

const CRC_TABLE = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i += 1) {
    let c = i;
    for (let k = 0; k < 8; k += 1) {
      c = (c & 1) ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

function crc32(bytes) {
  let crc = 0xffffffff;
  for (let i = 0; i < bytes.length; i += 1) {
    crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

function u16(value) {
  const bytes = new Uint8Array(2);
  bytes[0] = value & 0xff;
  bytes[1] = (value >>> 8) & 0xff;
  return bytes;
}

function u32(value) {
  const bytes = new Uint8Array(4);
  bytes[0] = value & 0xff;
  bytes[1] = (value >>> 8) & 0xff;
  bytes[2] = (value >>> 16) & 0xff;
  bytes[3] = (value >>> 24) & 0xff;
  return bytes;
}

function concatUint8Arrays(parts) {
  const total = parts.reduce((sum, arr) => sum + arr.length, 0);
  const result = new Uint8Array(total);
  let offset = 0;
  for (const part of parts) {
    result.set(part, offset);
    offset += part.length;
  }
  return result;
}

function toUint8Array(data) {
  if (data instanceof Uint8Array) return data;
  if (data instanceof ArrayBuffer) return new Uint8Array(data);
  if (typeof data === "string") return textEncoder.encode(data);
  throw new Error("Unsupported zip file data type");
}

async function normalizeFile(file) {
  const nameBytes = textEncoder.encode(file.name);
  let content = file.data;
  if (content instanceof Blob) {
    content = new Uint8Array(await content.arrayBuffer());
  }
  const dataBytes = toUint8Array(content);
  const checksum = crc32(dataBytes);
  return { nameBytes, dataBytes, checksum };
}

export async function createZipBlob(files) {
  const normalized = [];
  for (const file of files) {
    normalized.push(await normalizeFile(file));
  }

  const localParts = [];
  const centralParts = [];
  let offset = 0;

  for (const file of normalized) {
    const { nameBytes, dataBytes, checksum } = file;
    const localHeader = concatUint8Arrays([
      u32(0x04034b50),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(checksum),
      u32(dataBytes.length),
      u32(dataBytes.length),
      u16(nameBytes.length),
      u16(0),
      nameBytes
    ]);

    localParts.push(localHeader, dataBytes);

    const centralHeader = concatUint8Arrays([
      u32(0x02014b50),
      u16(20),
      u16(20),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(checksum),
      u32(dataBytes.length),
      u32(dataBytes.length),
      u16(nameBytes.length),
      u16(0),
      u16(0),
      u16(0),
      u16(0),
      u32(0),
      u32(offset),
      nameBytes
    ]);
    centralParts.push(centralHeader);

    offset += localHeader.length + dataBytes.length;
  }

  const centralSize = centralParts.reduce((sum, arr) => sum + arr.length, 0);
  const centralOffset = offset;
  const eocd = concatUint8Arrays([
    u32(0x06054b50),
    u16(0),
    u16(0),
    u16(normalized.length),
    u16(normalized.length),
    u32(centralSize),
    u32(centralOffset),
    u16(0)
  ]);

  return new Blob([...localParts, ...centralParts, eocd], {
    type: "application/zip"
  });
}
