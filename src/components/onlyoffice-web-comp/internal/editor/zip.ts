const crcTable = (() => {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let k = 0; k < 8; k++) {
      c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    }
    table[i] = c >>> 0;
  }
  return table;
})();

export function crc32(data: Uint8Array) {
  let crc = 0xffffffff;
  for (const byte of data) {
    crc = crcTable[(crc ^ byte) & 0xff] ^ (crc >>> 8);
  }
  return (crc ^ 0xffffffff) >>> 0;
}

export function readU16(data: Uint8Array, offset: number) {
  return data[offset] | (data[offset + 1] << 8);
}

export function readU32(data: Uint8Array, offset: number) {
  return (
    data[offset] |
    (data[offset + 1] << 8) |
    (data[offset + 2] << 16) |
    (data[offset + 3] << 24)
  ) >>> 0;
}

export function writeU16(data: Uint8Array, offset: number, value: number) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
}

export function writeU32(data: Uint8Array, offset: number, value: number) {
  data[offset] = value & 0xff;
  data[offset + 1] = (value >>> 8) & 0xff;
  data[offset + 2] = (value >>> 16) & 0xff;
  data[offset + 3] = (value >>> 24) & 0xff;
}

export function concatBytes(parts: Uint8Array[]) {
  const size = parts.reduce((total, part) => total + part.length, 0);
  const output = new Uint8Array(size);
  let offset = 0;
  for (const part of parts) {
    output.set(part, offset);
    offset += part.length;
  }
  return output;
}

async function decompressDeflateRaw(data: Uint8Array) {
  if (!("DecompressionStream" in globalThis)) {
    return null;
  }

  const stream = new Blob([data as Uint8Array<ArrayBuffer>])
    .stream()
    .pipeThrough(new DecompressionStream("deflate-raw"));
  return new Uint8Array(await new Response(stream).arrayBuffer());
}

export type ZipEntry = {
  name: string;
  nameBytes: Uint8Array;
  method: number;
  modTime: number;
  modDate: number;
  crc: number;
  compressedSize: number;
  uncompressedSize: number;
  compressedData: Uint8Array;
  internalAttrs: number;
  externalAttrs: number;
};

export function readZipEntries(data: ArrayBuffer) {
  const input = new Uint8Array(data);
  let eocd = -1;
  for (let i = input.length - 22; i >= Math.max(0, input.length - 65558); i--) {
    if (readU32(input, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) return null;

  const entryCount = readU16(input, eocd + 10);
  const centralOffset = readU32(input, eocd + 16);
  const decoder = new TextDecoder();
  const entries: ZipEntry[] = [];
  let readOffset = centralOffset;

  for (let i = 0; i < entryCount; i++) {
    if (readU32(input, readOffset) !== 0x02014b50) return null;

    const method = readU16(input, readOffset + 10);
    const modTime = readU16(input, readOffset + 12);
    const modDate = readU16(input, readOffset + 14);
    const crc = readU32(input, readOffset + 16);
    const compressedSize = readU32(input, readOffset + 20);
    const uncompressedSize = readU32(input, readOffset + 24);
    const nameLength = readU16(input, readOffset + 28);
    const extraLength = readU16(input, readOffset + 30);
    const commentLength = readU16(input, readOffset + 32);
    const internalAttrs = readU16(input, readOffset + 36);
    const externalAttrs = readU32(input, readOffset + 38);
    const localOffset = readU32(input, readOffset + 42);
    const nameBytes = input.slice(readOffset + 46, readOffset + 46 + nameLength);
    const name = decoder.decode(nameBytes);

    if (readU32(input, localOffset) !== 0x04034b50) return null;
    const localNameLength = readU16(input, localOffset + 26);
    const localExtraLength = readU16(input, localOffset + 28);
    const dataStart = localOffset + 30 + localNameLength + localExtraLength;
    const compressedData = input.slice(dataStart, dataStart + compressedSize);

    entries.push({
      name,
      nameBytes,
      method,
      modTime,
      modDate,
      crc,
      compressedSize,
      uncompressedSize,
      compressedData,
      internalAttrs,
      externalAttrs,
    });

    readOffset += 46 + nameLength + extraLength + commentLength;
  }

  return entries;
}

export async function readZipEntryData(entry: ZipEntry) {
  if (entry.method === 0) {
    return entry.compressedData;
  }
  if (entry.method === 8) {
    return decompressDeflateRaw(entry.compressedData);
  }
  return null;
}

function buildLocalHeader(
  entry: Pick<ZipEntry, "nameBytes" | "method" | "modTime" | "modDate">,
  crc: number,
  compressedSize: number,
  uncompressedSize: number,
) {
  const localHeader = new Uint8Array(30 + entry.nameBytes.length);
  writeU32(localHeader, 0, 0x04034b50);
  writeU16(localHeader, 4, 20);
  writeU16(localHeader, 6, 0);
  writeU16(localHeader, 8, entry.method);
  writeU16(localHeader, 10, entry.modTime);
  writeU16(localHeader, 12, entry.modDate);
  writeU32(localHeader, 14, crc);
  writeU32(localHeader, 18, compressedSize);
  writeU32(localHeader, 22, uncompressedSize);
  writeU16(localHeader, 26, entry.nameBytes.length);
  writeU16(localHeader, 28, 0);
  localHeader.set(entry.nameBytes, 30);
  return localHeader;
}

function buildCentralHeader(
  entry: Pick<
    ZipEntry,
    | "nameBytes"
    | "method"
    | "modTime"
    | "modDate"
    | "internalAttrs"
    | "externalAttrs"
  >,
  crc: number,
  compressedSize: number,
  uncompressedSize: number,
  localOffset: number,
) {
  const centralHeader = new Uint8Array(46 + entry.nameBytes.length);
  writeU32(centralHeader, 0, 0x02014b50);
  writeU16(centralHeader, 4, 20);
  writeU16(centralHeader, 6, 20);
  writeU16(centralHeader, 8, 0);
  writeU16(centralHeader, 10, entry.method);
  writeU16(centralHeader, 12, entry.modTime);
  writeU16(centralHeader, 14, entry.modDate);
  writeU32(centralHeader, 16, crc);
  writeU32(centralHeader, 20, compressedSize);
  writeU32(centralHeader, 24, uncompressedSize);
  writeU16(centralHeader, 28, entry.nameBytes.length);
  writeU16(centralHeader, 30, 0);
  writeU16(centralHeader, 32, 0);
  writeU16(centralHeader, 34, 0);
  writeU16(centralHeader, 36, entry.internalAttrs);
  writeU32(centralHeader, 38, entry.externalAttrs);
  writeU32(centralHeader, 42, localOffset);
  centralHeader.set(entry.nameBytes, 46);
  return centralHeader;
}

export type ZipReplacement = {
  data: Uint8Array;
  modTime?: number;
  modDate?: number;
};

export type ZipXmlSizeStats = {
  totalSize: number;
  entryCount: number;
};

export function isOfficeXmlEntryName(name: string) {
  return /\.(xml|rels)$/i.test(name);
}

export function getZipXmlUncompressedSize(
  entries: readonly ZipEntry[],
): ZipXmlSizeStats {
  let totalSize = 0;
  let entryCount = 0;

  for (const entry of entries) {
    if (entry.name.endsWith("/") || !isOfficeXmlEntryName(entry.name)) {
      continue;
    }

    totalSize += entry.uncompressedSize;
    entryCount += 1;
  }

  return { totalSize, entryCount };
}

export function writeZipEntries(
  entries: ZipEntry[],
  replacements: Map<string, ZipReplacement>,
) {
  const encoder = new TextEncoder();
  const localParts: Uint8Array[] = [];
  const centralParts: Uint8Array[] = [];
  let outputOffset = 0;
  const emitted = new Set<string>();
  let emittedCount = 0;

  const emit = (
    entry: Pick<
      ZipEntry,
      | "name"
      | "nameBytes"
      | "method"
      | "modTime"
      | "modDate"
      | "crc"
      | "compressedData"
      | "uncompressedSize"
      | "internalAttrs"
      | "externalAttrs"
    >,
    replacement?: ZipReplacement,
  ) => {
    const outputData = replacement?.data ?? entry.compressedData;
    const method = replacement ? 0 : entry.method;
    const crc = replacement ? crc32(outputData) : entry.crc;
    const uncompressedSize = replacement
      ? outputData.length
      : entry.uncompressedSize;
    const modTime = replacement?.modTime ?? entry.modTime;
    const modDate = replacement?.modDate ?? entry.modDate;
    const nextEntry = { ...entry, method, modTime, modDate };

    const localHeader = buildLocalHeader(
      nextEntry,
      crc,
      outputData.length,
      uncompressedSize,
    );
    localParts.push(localHeader, outputData);
    centralParts.push(
      buildCentralHeader(
        nextEntry,
        crc,
        outputData.length,
        uncompressedSize,
        outputOffset,
      ),
    );
    outputOffset += localHeader.length + outputData.length;
    emitted.add(entry.name);
    emittedCount += 1;
  };

  for (const entry of entries) {
    emit(entry, replacements.get(entry.name));
  }

  for (const [name, replacement] of replacements) {
    if (emitted.has(name)) continue;
    const nameBytes = encoder.encode(name);
    emit({
      name,
      nameBytes,
      method: 0,
      modTime: replacement.modTime ?? 0,
      modDate: replacement.modDate ?? 0,
      crc: 0,
      compressedData: replacement.data,
      uncompressedSize: replacement.data.length,
      internalAttrs: 0,
      externalAttrs: 0,
    });
  }

  const centralDirectory = concatBytes(centralParts);
  const end = new Uint8Array(22);
  writeU32(end, 0, 0x06054b50);
  writeU16(end, 8, emittedCount);
  writeU16(end, 10, emittedCount);
  writeU32(end, 12, centralDirectory.length);
  writeU32(end, 16, outputOffset);

  return concatBytes([...localParts, centralDirectory, end]).buffer;
}
