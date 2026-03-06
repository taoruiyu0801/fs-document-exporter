const ILLEGAL_RE = /[<>:"/\\|?*\x00-\x1F]/g;
const RESERVED_RE = /^(con|prn|aux|nul|com[1-9]|lpt[1-9])$/i;
const MAX_SEGMENT_LENGTH = 80;
const MAX_FILENAME_LENGTH = 140;
const MAX_PATH_LENGTH = 220;
const INVISIBLE_RE = /[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g;
const CTRL_RE = /[\u0000-\u001F\u007F-\u009F]/g;
const TITLE_SUFFIX_RE =
  /\s*[-_–—|]\s*(飞书云文档|飞书文档|Feishu\s*Docs?|Lark\s*Docs?)\s*$/i;

function truncateByCodePoints(value, maxLen) {
  const chars = Array.from(String(value || ""));
  if (chars.length <= maxLen) {
    return chars.join("");
  }
  return chars.slice(0, maxLen).join("");
}

function normalizeRawText(value) {
  let text = String(value || "");
  try {
    text = text.normalize("NFKC");
  } catch {
    // ignore normalize failures
  }
  return text
    .replace(INVISIBLE_RE, "")
    .replace(CTRL_RE, "")
    .replace(/\u00A0/g, " ");
}

export function normalizeDocumentTitle(input, fallback = "feishu-document") {
  const raw = normalizeRawText(input)
    .replace(TITLE_SUFFIX_RE, "")
    .replace(/[，,]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
  return raw || fallback;
}

export function buildExportBaseName(input, fallback = "feishu-document") {
  const normalized = normalizeDocumentTitle(input, fallback);
  return sanitizeFilename(normalized, fallback, 80);
}

export function sanitizeFilename(input, fallback = "untitled", maxLength = MAX_FILENAME_LENGTH) {
  const raw = normalizeRawText(input).trim();
  const base = raw
    .replace(ILLEGAL_RE, "_")
    .replace(/\s+/g, " ")
    .replace(/^\.+/, "")
    .replace(/\.+$/, "")
    .replace(/[. ]+$/g, "");
  const safe = truncateByCodePoints(base || fallback, maxLength) || fallback;
  return RESERVED_RE.test(safe) ? `${safe}_` : safe;
}

export function sanitizePathParts(pathParts = []) {
  const rawParts = pathParts
    .flatMap((part) =>
      normalizeRawText(part)
        .split(/[\\/]+/)
        .map((segment) => segment.trim())
        .filter(Boolean)
    )
    .filter(Boolean);

  const looksAbsolutePath =
    rawParts.length > 0 &&
    (/^[A-Za-z]:$/.test(rawParts[0]) ||
      rawParts[0].startsWith("\\") ||
      rawParts[0].startsWith("/") ||
      rawParts[0] === "~");

  if (looksAbsolutePath) {
    return [];
  }

  const safeParts = rawParts
    .map((part) => sanitizeFilename(part, "untitled", MAX_SEGMENT_LENGTH))
    .filter((part) => part !== "." && part !== "..")
    .filter(Boolean);
  if (safeParts.length > 1 && /^[A-Za-z]_$/i.test(safeParts[0])) {
    return safeParts.slice(1);
  }
  return safeParts;
}

export function ensureExtension(filename, extWithoutDot) {
  const ext = String(extWithoutDot || "").replace(/^\./, "");
  const safeName = sanitizeFilename(filename || "untitled", "untitled", MAX_FILENAME_LENGTH);
  if (!ext) {
    return safeName;
  }
  const suffix = `.${ext}`;
  if (safeName.toLowerCase().endsWith(suffix.toLowerCase())) {
    return truncateByCodePoints(safeName, MAX_FILENAME_LENGTH);
  }
  const headMax = Math.max(1, MAX_FILENAME_LENGTH - suffix.length);
  const head = truncateByCodePoints(safeName, headMax).replace(/[. ]+$/g, "");
  return `${head || "untitled"}${suffix}`;
}

export function buildDownloadPath({ filename, pathParts = [] }) {
  const safeParts = sanitizePathParts(pathParts);
  let filenameInput = filename;
  if ((!filenameInput || !String(filenameInput).trim()) && safeParts.length > 0) {
    filenameInput = safeParts.pop();
  }

  let safeFile = sanitizeFilename(filenameInput, "untitled", MAX_FILENAME_LENGTH);
  if (!safeParts.length) {
    return safeFile || "untitled";
  }
  const dir = safeParts.join("/");
  let full = `${dir}/${safeFile}`;
  if (full.length > MAX_PATH_LENGTH) {
    const spare = Math.max(20, MAX_PATH_LENGTH - dir.length - 1);
    safeFile = sanitizeFilename(safeFile, "untitled", spare);
    full = `${dir}/${safeFile}`;
  }
  return full;
}

export function arrayBufferToBase64(buffer) {
  const bytes = new Uint8Array(buffer);
  const chunkSize = 0x8000;
  let binary = "";
  for (let i = 0; i < bytes.length; i += chunkSize) {
    const chunk = bytes.subarray(i, i + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

export function dataUrlToUint8Array(dataUrl) {
  const match = /^data:([^;,]+)?(?:;charset=[^;,]+)?;base64,(.*)$/i.exec(
    dataUrl || ""
  );
  if (!match) {
    throw new Error("Invalid data URL");
  }
  const binary = atob(match[2]);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) {
    bytes[i] = binary.charCodeAt(i);
  }
  return bytes;
}

export function guessExtensionFromDataUrl(dataUrl, fallback = "bin") {
  const match = /^data:([^;,]+)[;,]/i.exec(dataUrl || "");
  if (!match) {
    return fallback;
  }
  const mime = match[1].toLowerCase();
  if (mime.includes("png")) return "png";
  if (mime.includes("jpeg") || mime.includes("jpg")) return "jpg";
  if (mime.includes("webp")) return "webp";
  if (mime.includes("gif")) return "gif";
  if (mime.includes("svg")) return "svg";
  if (mime.includes("bmp")) return "bmp";
  return fallback;
}

export async function blobToDataUrl(blob) {
  const mime = blob.type || "application/octet-stream";
  const buffer = await blob.arrayBuffer();
  return `data:${mime};base64,${arrayBufferToBase64(buffer)}`;
}
