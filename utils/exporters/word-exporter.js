import {
  buildExportBaseName,
  dataUrlToUint8Array,
  ensureExtension,
  guessExtensionFromDataUrl,
  sanitizeFilename
} from "../filename.js";
import { createZipBlob } from "../zip.js";

const OOXML_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types";
const PACKAGE_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const PX_TO_EMU = 9525;
const DOCX_MAX_IMAGE_WIDTH_PX = 580;
const DOCX_MAX_IMAGE_HEIGHT_PX = 760;

const BLOCK_TAGS = new Set([
  "p",
  "div",
  "section",
  "article",
  "header",
  "footer",
  "main",
  "aside",
  "blockquote",
  "pre",
  "figure",
  "figcaption",
  "h1",
  "h2",
  "h3",
  "h4",
  "h5",
  "h6",
  "ul",
  "ol",
  "li",
  "table",
  "thead",
  "tbody",
  "tfoot",
  "tr",
  "td",
  "th"
]);

const HEADING_STYLE_MAP = {
  h1: "Heading1",
  h2: "Heading2",
  h3: "Heading3",
  h4: "Heading3",
  h5: "Heading3",
  h6: "Heading3"
};

const SELF_CLOSING_TAGS = new Set(["br", "img", "hr", "meta", "link", "input"]);

const MIME_BY_EXT = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  webp: "image/webp",
  svg: "image/svg+xml"
};
const DOCX_MIME =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const INVISIBLE_CHARS_RE = /[\u200B-\u200F\u202A-\u202E\u2060-\u206F\uFEFF]/g;
const CTRL_CHARS_RE = /[\u0000-\u001F\u007F-\u009F]/g;
const TITLE_SUFFIX_RE =
  /\s*[-_–—|]\s*(飞书云文档|飞书文档|Feishu\s*Docs?|Lark\s*Docs?)\s*$/i;

function escapeXml(text) {
  return String(text || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function escapeXmlAttr(text) {
  return escapeXml(text).replace(/"/g, "&quot;").replace(/'/g, "&apos;");
}

function decodeHtmlEntities(text) {
  const named = {
    amp: "&",
    lt: "<",
    gt: ">",
    quot: "\"",
    apos: "'",
    nbsp: " "
  };
  return String(text || "").replace(/&(#x?[0-9a-fA-F]+|[a-zA-Z]+);/g, (_full, token) => {
    if (!token) return "";
    if (token[0] === "#") {
      const hex = token[1] === "x" || token[1] === "X";
      const raw = hex ? token.slice(2) : token.slice(1);
      const num = Number.parseInt(raw, hex ? 16 : 10);
      if (!Number.isFinite(num) || num <= 0) return "";
      try {
        return String.fromCodePoint(num);
      } catch {
        return "";
      }
    }
    return named[token] || `&${token};`;
  });
}

function sanitizeTextNode(text) {
  return decodeHtmlEntities(text)
    .replace(INVISIBLE_CHARS_RE, "")
    .replace(CTRL_CHARS_RE, "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ");
}

function sanitizeUrl(href, baseUrl) {
  const raw = String(href || "").trim();
  if (!raw || raw.startsWith("#")) return "";
  if (/^(javascript|vbscript|data):/i.test(raw)) return "";
  try {
    return new URL(raw, baseUrl || "https://example.com").href;
  } catch {
    return "";
  }
}

function normalizeExportTitle(text) {
  const value = sanitizeTextNode(text).replace(TITLE_SUFFIX_RE, "").trim();
  return value || "未命名文档";
}

function parseAttributes(rawTag) {
  const attrs = {};
  const source = String(rawTag || "");
  const body = source
    .replace(/^<\s*\/?\s*[^\s/>]+/i, "")
    .replace(/\/?\s*>$/i, "");
  const attrRegex = /([^\s=/>]+)(?:\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s"'=<>`]+)))?/g;
  let match = attrRegex.exec(body);
  while (match) {
    const key = String(match[1] || "").toLowerCase();
    const value = match[2] ?? match[3] ?? match[4] ?? "";
    if (key) {
      attrs[key] = decodeHtmlEntities(value);
    }
    match = attrRegex.exec(body);
  }
  return attrs;
}

function parseTag(token) {
  const raw = String(token || "");
  if (!raw.startsWith("<")) return null;
  const endTagMatch = /^<\s*\/\s*([a-zA-Z0-9:_-]+)/.exec(raw);
  if (endTagMatch) {
    return {
      type: "end",
      tagName: endTagMatch[1].toLowerCase(),
      attrs: {}
    };
  }

  const startTagMatch = /^<\s*([a-zA-Z0-9:_-]+)/.exec(raw);
  if (!startTagMatch) return null;
  const tagName = startTagMatch[1].toLowerCase();
  const explicitSelfClosing = /\/\s*>$/.test(raw);
  return {
    type: explicitSelfClosing || SELF_CLOSING_TAGS.has(tagName) ? "self" : "start",
    tagName,
    attrs: parseAttributes(raw)
  };
}

function parseCssSizePx(rawValue) {
  const value = String(rawValue || "").trim().toLowerCase();
  if (!value) return null;
  if (value.endsWith("%")) return null;
  const match = /^([0-9]+(?:\.[0-9]+)?)(px)?$/.exec(value);
  if (!match) return null;
  const number = Number.parseFloat(match[1]);
  if (!Number.isFinite(number) || number <= 0) return null;
  return number;
}

function parseCssSizeFromStyle(styleText, key) {
  const style = String(styleText || "");
  const regex = new RegExp(`${key}\\s*:\\s*([^;]+)`, "i");
  const match = regex.exec(style);
  if (!match?.[1]) return null;
  return parseCssSizePx(match[1]);
}

function normalizeImageSize(widthPx, heightPx, fallbackWidthPx, fallbackHeightPx) {
  let width = Number.isFinite(widthPx) && widthPx > 0 ? widthPx : null;
  let height = Number.isFinite(heightPx) && heightPx > 0 ? heightPx : null;
  const fallbackWidth =
    Number.isFinite(fallbackWidthPx) && fallbackWidthPx > 0 ? fallbackWidthPx : 1200;
  const fallbackHeight =
    Number.isFinite(fallbackHeightPx) && fallbackHeightPx > 0 ? fallbackHeightPx : 800;

  if (!width && !height) {
    width = fallbackWidth;
    height = fallbackHeight;
  } else if (!width && height) {
    width = (height * fallbackWidth) / Math.max(1, fallbackHeight);
  } else if (width && !height) {
    height = (width * fallbackHeight) / Math.max(1, fallbackWidth);
  }

  const ratio = Math.min(
    1,
    DOCX_MAX_IMAGE_WIDTH_PX / Math.max(1, width),
    DOCX_MAX_IMAGE_HEIGHT_PX / Math.max(1, height)
  );
  return {
    widthPx: Math.max(1, Math.round(width * ratio)),
    heightPx: Math.max(1, Math.round(height * ratio))
  };
}

function tryReadPngSize(bytes) {
  if (bytes.length < 24) return null;
  const isPng =
    bytes[0] === 0x89 &&
    bytes[1] === 0x50 &&
    bytes[2] === 0x4e &&
    bytes[3] === 0x47 &&
    bytes[4] === 0x0d &&
    bytes[5] === 0x0a &&
    bytes[6] === 0x1a &&
    bytes[7] === 0x0a;
  if (!isPng) return null;
  const width =
    (bytes[16] << 24) | (bytes[17] << 16) | (bytes[18] << 8) | (bytes[19] & 0xff);
  const height =
    (bytes[20] << 24) | (bytes[21] << 16) | (bytes[22] << 8) | (bytes[23] & 0xff);
  if (!width || !height) return null;
  return { width: width >>> 0, height: height >>> 0 };
}

function tryReadGifSize(bytes) {
  if (bytes.length < 10) return null;
  const signature =
    String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3], bytes[4], bytes[5]) || "";
  if (signature !== "GIF87a" && signature !== "GIF89a") return null;
  const width = bytes[6] | (bytes[7] << 8);
  const height = bytes[8] | (bytes[9] << 8);
  if (!width || !height) return null;
  return { width, height };
}

function tryReadJpegSize(bytes) {
  if (bytes.length < 4 || bytes[0] !== 0xff || bytes[1] !== 0xd8) return null;
  let offset = 2;
  while (offset + 8 < bytes.length) {
    if (bytes[offset] !== 0xff) {
      offset += 1;
      continue;
    }
    const marker = bytes[offset + 1];
    if (marker === 0xd9 || marker === 0xda) break;
    const segmentLength = (bytes[offset + 2] << 8) | bytes[offset + 3];
    if (segmentLength < 2) break;
    const isSof =
      marker === 0xc0 ||
      marker === 0xc1 ||
      marker === 0xc2 ||
      marker === 0xc3 ||
      marker === 0xc5 ||
      marker === 0xc6 ||
      marker === 0xc7 ||
      marker === 0xc9 ||
      marker === 0xca ||
      marker === 0xcb ||
      marker === 0xcd ||
      marker === 0xce ||
      marker === 0xcf;
    if (isSof && offset + 8 < bytes.length) {
      const height = (bytes[offset + 5] << 8) | bytes[offset + 6];
      const width = (bytes[offset + 7] << 8) | bytes[offset + 8];
      if (width && height) {
        return { width, height };
      }
    }
    offset += 2 + segmentLength;
  }
  return null;
}

function readWebpChunk(bytes, offset) {
  if (offset + 8 > bytes.length) return null;
  const type = String.fromCharCode(
    bytes[offset],
    bytes[offset + 1],
    bytes[offset + 2],
    bytes[offset + 3]
  );
  const size =
    bytes[offset + 4] |
    (bytes[offset + 5] << 8) |
    (bytes[offset + 6] << 16) |
    (bytes[offset + 7] << 24);
  return { type, size: size >>> 0, dataOffset: offset + 8 };
}

function tryReadWebpSize(bytes) {
  if (bytes.length < 30) return null;
  const riff =
    String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) === "RIFF" &&
    String.fromCharCode(bytes[8], bytes[9], bytes[10], bytes[11]) === "WEBP";
  if (!riff) return null;

  let offset = 12;
  while (offset + 8 <= bytes.length) {
    const chunk = readWebpChunk(bytes, offset);
    if (!chunk) break;
    const { type, size, dataOffset } = chunk;

    if (type === "VP8X" && dataOffset + 10 <= bytes.length) {
      const width =
        1 + bytes[dataOffset + 4] + (bytes[dataOffset + 5] << 8) + (bytes[dataOffset + 6] << 16);
      const height =
        1 + bytes[dataOffset + 7] + (bytes[dataOffset + 8] << 8) + (bytes[dataOffset + 9] << 16);
      if (width && height) {
        return { width, height };
      }
    }

    if (type === "VP8 " && dataOffset + 10 <= bytes.length) {
      const startCode =
        bytes[dataOffset + 3] === 0x9d &&
        bytes[dataOffset + 4] === 0x01 &&
        bytes[dataOffset + 5] === 0x2a;
      if (startCode) {
        const width = bytes[dataOffset + 6] | ((bytes[dataOffset + 7] & 0x3f) << 8);
        const height = bytes[dataOffset + 8] | ((bytes[dataOffset + 9] & 0x3f) << 8);
        if (width && height) {
          return { width, height };
        }
      }
    }

    if (type === "VP8L" && dataOffset + 5 <= bytes.length) {
      const signature = bytes[dataOffset] === 0x2f;
      if (signature) {
        const b1 = bytes[dataOffset + 1];
        const b2 = bytes[dataOffset + 2];
        const b3 = bytes[dataOffset + 3];
        const b4 = bytes[dataOffset + 4];
        const width = 1 + (b1 | ((b2 & 0x3f) << 8));
        const height = 1 + ((b2 >> 6) | (b3 << 2) | ((b4 & 0x0f) << 10));
        if (width && height) {
          return { width, height };
        }
      }
    }

    const step = 8 + size + (size % 2);
    offset += step;
  }
  return null;
}

function detectImageSize(bytes) {
  return (
    tryReadPngSize(bytes) ||
    tryReadJpegSize(bytes) ||
    tryReadGifSize(bytes) ||
    tryReadWebpSize(bytes) ||
    null
  );
}

function hashText(text) {
  let hash = 2166136261;
  const value = String(text || "");
  for (let i = 0; i < value.length; i += 1) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 16777619);
  }
  return (hash >>> 0).toString(16);
}

function parseHtmlToBlocks(html, baseUrl) {
  // Pre-process: strip <head>...</head> entirely and remove inline <style> blocks
  // Also strip the injected-html.js document wrapper header to avoid duplicate title
  let cleanedHtml = String(html || "")
    .replace(/<head[\s>][\s\S]*?<\/head\s*>/gi, "")
    .replace(/<style[\s>][\s\S]*?<\/style\s*>/gi, "")
    .replace(/<script[\s>][\s\S]*?<\/script\s*>/gi, "")
    .replace(/<noscript[\s>][\s\S]*?<\/noscript\s*>/gi, "")
    .replace(/<title[\s>][\s\S]*?<\/title\s*>/gi, "")
    .replace(/<meta[^>]*>/gi, "")
    .replace(/<div[^>]*class="[^"]*document-header[^"]*"[^>]*>[\s\S]*?<\/div\s*>/gi, "");

  const blocks = [];
  let currentParagraph = null;
  let boldDepth = 0;
  let italicDepth = 0;
  let underlineDepth = 0;
  let strikeDepth = 0;
  const linkStack = [];
  const listStack = [];
  let skipUntilClose = null; // tag name to skip until its closing tag

  function currentTextState() {
    const href = linkStack.length ? linkStack[linkStack.length - 1] : "";
    return {
      bold: boldDepth > 0,
      italic: italicDepth > 0,
      underline: underlineDepth > 0,
      strike: strikeDepth > 0,
      href
    };
  }

  function flushParagraph(force = false) {
    if (!currentParagraph) return;
    const hasContent = currentParagraph.runs.some((run) => {
      if (run.type === "break") return true;
      return String(run.text || "").length > 0;
    });
    if (hasContent || force) {
      blocks.push(currentParagraph);
    }
    currentParagraph = null;
  }

  function ensureParagraph(styleId = "Normal") {
    if (!currentParagraph) {
      currentParagraph = {
        type: "paragraph",
        styleId,
        runs: []
      };
      return;
    }
    if (!currentParagraph.styleId) {
      currentParagraph.styleId = styleId;
    }
  }

  function pushText(text) {
    const normalized = sanitizeTextNode(text);
    if (!normalized) return;
    ensureParagraph("Normal");
    const state = currentTextState();
    currentParagraph.runs.push({
      type: "text",
      text: normalized,
      bold: state.bold,
      italic: state.italic,
      underline: state.underline,
      strike: state.strike,
      href: state.href
    });
  }

  function pushListPrefix() {
    if (!listStack.length) return;
    const depth = Math.max(0, listStack.length - 1);
    const top = listStack[listStack.length - 1];
    const prefix = top.type === "ol" ? `${top.index}. ` : "\u2022 ";
    if (top.type === "ol") {
      top.index += 1;
    }
    const indent = "  ".repeat(depth);
    pushText(`${indent}${prefix}`);
  }

  const tokenRegex = /<!--[\s\S]*?-->|<\/?[^>]+>|[^<]+/g;
  let tokenMatch = tokenRegex.exec(cleanedHtml);
  while (tokenMatch) {
    const token = tokenMatch[0] || "";

    if (token.startsWith("<!--")) {
      tokenMatch = tokenRegex.exec(cleanedHtml);
      continue;
    }

    // If we're skipping content inside a tag (e.g. leftover style/script), check for close
    if (skipUntilClose) {
      if (token.startsWith("<")) {
        const tag = parseTag(token);
        if (tag && tag.type === "end" && tag.tagName === skipUntilClose) {
          skipUntilClose = null;
        }
      }
      tokenMatch = tokenRegex.exec(cleanedHtml);
      continue;
    }

    if (!token.startsWith("<")) {
      pushText(token);
      tokenMatch = tokenRegex.exec(cleanedHtml);
      continue;
    }

    const tag = parseTag(token);
    if (!tag) {
      tokenMatch = tokenRegex.exec(cleanedHtml);
      continue;
    }
    const tagName = tag.tagName;
    const isStart = tag.type === "start" || tag.type === "self";
    const isEnd = tag.type === "end";

    if (isStart) {
      if (tagName === "script" || tagName === "style" || tagName === "noscript" || tagName === "head") {
        skipUntilClose = tagName;
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "br") {
        ensureParagraph("Normal");
        currentParagraph.runs.push({ type: "break" });
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "hr") {
        flushParagraph();
        blocks.push({
          type: "paragraph",
          styleId: "Normal",
          runs: [{ type: "text", text: "----------" }]
        });
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "ul") {
        flushParagraph();
        listStack.push({ type: "ul", index: 1 });
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "ol") {
        flushParagraph();
        listStack.push({ type: "ol", index: 1 });
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "li") {
        flushParagraph();
        ensureParagraph("Normal");
        pushListPrefix();
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "a") {
        const href = sanitizeUrl(tag.attrs.href || "", baseUrl);
        linkStack.push(href);
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "strong" || tagName === "b") boldDepth += 1;
      if (tagName === "em" || tagName === "i") italicDepth += 1;
      if (tagName === "u") underlineDepth += 1;
      if (tagName === "s" || tagName === "strike" || tagName === "del") strikeDepth += 1;

      if (tagName === "img") {
        const tokenValue = String(tag.attrs["data-token"] || "").trim();
        const altText =
          String(tag.attrs.alt || tag.attrs["data-name"] || tokenValue || "image").trim();
        const src = String(tag.attrs.src || "").trim();
        const widthFromAttr = parseCssSizePx(tag.attrs.width || "");
        const heightFromAttr = parseCssSizePx(tag.attrs.height || "");
        const widthFromStyle = parseCssSizeFromStyle(tag.attrs.style || "", "width");
        const heightFromStyle = parseCssSizeFromStyle(tag.attrs.style || "", "height");
        const widthPx = widthFromStyle || widthFromAttr || null;
        const heightPx = heightFromStyle || heightFromAttr || null;
        flushParagraph();
        blocks.push({
          type: "image",
          token: tokenValue,
          src,
          alt: altText,
          widthPx,
          heightPx
        });
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName === "td" || tagName === "th") {
        ensureParagraph("Normal");
        const hasText = currentParagraph.runs.some((run) => run.type === "text");
        if (hasText) {
          currentParagraph.runs.push({ type: "text", text: " | " });
        }
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (tagName in HEADING_STYLE_MAP) {
        flushParagraph();
        ensureParagraph(HEADING_STYLE_MAP[tagName]);
        tokenMatch = tokenRegex.exec(cleanedHtml);
        continue;
      }

      if (BLOCK_TAGS.has(tagName) && tagName !== "td" && tagName !== "th") {
        flushParagraph();
      }
    }

    if (isEnd || tag.type === "self") {
      if (tagName === "a") {
        if (linkStack.length) linkStack.pop();
      }
      if (tagName === "strong" || tagName === "b") boldDepth = Math.max(0, boldDepth - 1);
      if (tagName === "em" || tagName === "i") italicDepth = Math.max(0, italicDepth - 1);
      if (tagName === "u") underlineDepth = Math.max(0, underlineDepth - 1);
      if (tagName === "s" || tagName === "strike" || tagName === "del") {
        strikeDepth = Math.max(0, strikeDepth - 1);
      }

      if (tagName === "ul" || tagName === "ol") {
        flushParagraph();
        if (listStack.length) {
          listStack.pop();
        }
      }

      if (
        tagName === "p" ||
        tagName === "div" ||
        tagName === "section" ||
        tagName === "article" ||
        tagName === "blockquote" ||
        tagName === "pre" ||
        tagName === "figure" ||
        tagName === "figcaption" ||
        tagName === "li" ||
        tagName === "tr" ||
        tagName === "table" ||
        tagName in HEADING_STYLE_MAP
      ) {
        flushParagraph();
      }
    }

    tokenMatch = tokenRegex.exec(cleanedHtml);
  }

  flushParagraph();
  return blocks;
}

function buildRunPropertiesXml(run, hyperlink = false) {
  const parts = [];
  if (hyperlink) {
    parts.push('<w:rStyle w:val="Hyperlink"/>');
  }
  if (run.bold) parts.push("<w:b/>");
  if (run.italic) parts.push("<w:i/>");
  if (run.underline || hyperlink) parts.push('<w:u w:val="single"/>');
  if (run.strike) parts.push("<w:strike/>");
  if (!parts.length) return "";
  return `<w:rPr>${parts.join("")}</w:rPr>`;
}

function buildTextRunXml(run, hyperlink = false) {
  if (run.type === "break") {
    return "<w:r><w:br/></w:r>";
  }
  const textValue = String(run.text || "");
  if (!textValue) return "";
  const preserve = /^\s|\s$|\s{2,}/.test(textValue);
  return `<w:r>${buildRunPropertiesXml(run, hyperlink)}<w:t${preserve ? ' xml:space="preserve"' : ""
    }>${escapeXml(textValue)}</w:t></w:r>`;
}

function resolveDataUrlMimeType(dataUrl, ext = "png") {
  const match = /^data:([^;,]+)[;,]/i.exec(String(dataUrl || ""));
  if (match?.[1]) return match[1].toLowerCase();
  return MIME_BY_EXT[ext] || "application/octet-stream";
}

function createWordContext(payload, imageMap) {
  const hyperlinkRelMap = new Map();
  const mediaRelMap = new Map();
  const mediaEntries = [];
  let nextRelIndex = 2; // rId1 -> styles.xml
  let nextDocPrId = 1;

  function getHyperlinkRelId(href) {
    const normalized = sanitizeUrl(href, payload.currentUrl || payload.url || "");
    if (!normalized) return "";
    if (hyperlinkRelMap.has(normalized)) {
      return hyperlinkRelMap.get(normalized);
    }
    const relId = `rId${nextRelIndex}`;
    nextRelIndex += 1;
    hyperlinkRelMap.set(normalized, relId);
    return relId;
  }

  function resolveImageEntry(imageBlock) {
    const token = String(imageBlock?.token || "").trim();
    const source = String(imageBlock?.src || "").trim();
    let dataUrl = "";
    if (token && imageMap.has(token)) {
      dataUrl = imageMap.get(token) || "";
    } else if (source.startsWith("data:")) {
      dataUrl = source;
    }
    if (!dataUrl) return null;

    const keySource = token || source || dataUrl.slice(0, 1024);
    const cacheKey = `${token ? "token" : "src"}:${hashText(keySource)}:${dataUrl.length}`;
    if (mediaRelMap.has(cacheKey)) {
      return mediaRelMap.get(cacheKey);
    }

    const ext = guessExtensionFromDataUrl(dataUrl, "png").toLowerCase();
    const mimeType = resolveDataUrlMimeType(dataUrl, ext);
    const data = dataUrlToUint8Array(dataUrl);
    const parsed = detectImageSize(data);
    const mediaName = `image-${String(mediaEntries.length + 1).padStart(4, "0")}.${ext}`;
    const relId = `rId${nextRelIndex}`;
    nextRelIndex += 1;
    const entry = {
      relId,
      mediaName,
      mediaPath: `word/media/${mediaName}`,
      mimeType,
      ext,
      data,
      naturalWidth: parsed?.width || null,
      naturalHeight: parsed?.height || null
    };
    mediaRelMap.set(cacheKey, entry);
    mediaEntries.push(entry);
    return entry;
  }

  function buildParagraphXml(paragraph) {
    const runs = Array.isArray(paragraph?.runs) ? paragraph.runs : [];
    const runXml = [];
    for (const run of runs) {
      if (run.type === "break") {
        runXml.push(buildTextRunXml(run));
        continue;
      }
      const text = String(run.text || "");
      if (!text) continue;
      const linkRelId = run.href ? getHyperlinkRelId(run.href) : "";
      if (linkRelId) {
        runXml.push(
          `<w:hyperlink r:id="${escapeXmlAttr(linkRelId)}" w:history="1">${buildTextRunXml(
            run,
            true
          )}</w:hyperlink>`
        );
      } else {
        runXml.push(buildTextRunXml(run, false));
      }
    }
    const styleId = String(paragraph?.styleId || "Normal");
    const pPr =
      styleId && styleId !== "Normal"
        ? `<w:pPr><w:pStyle w:val="${escapeXmlAttr(styleId)}"/></w:pPr>`
        : "";
    const finalRuns = runXml.join("");
    if (!pPr && !finalRuns) return "<w:p/>";
    return `<w:p>${pPr}${finalRuns || "<w:r/>"}</w:p>`;
  }

  function buildImageParagraphXml(imageBlock) {
    const entry = resolveImageEntry(imageBlock);
    if (!entry) {
      return "";
    }

    const normalizedSize = normalizeImageSize(
      imageBlock?.widthPx,
      imageBlock?.heightPx,
      entry.naturalWidth,
      entry.naturalHeight
    );
    const cx = Math.max(1, Math.round(normalizedSize.widthPx * PX_TO_EMU));
    const cy = Math.max(1, Math.round(normalizedSize.heightPx * PX_TO_EMU));
    const docPrId = nextDocPrId;
    nextDocPrId += 1;
    const descr = escapeXmlAttr(String(imageBlock?.alt || entry.mediaName || "image"));
    const imageName = escapeXmlAttr(entry.mediaName);
    const relId = escapeXmlAttr(entry.relId);

    return `<w:p><w:r><w:drawing><wp:inline xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:docPr id="${docPrId}" name="${imageName}" descr="${descr}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${imageName}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${relId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
  }

  function buildDocumentXml(blocks) {
    const bodyXml = [];
    const docTitle = normalizeExportTitle(String(payload?.title || "").trim());
    const currentUrl = sanitizeUrl(payload?.currentUrl || payload?.url || "", payload?.currentUrl);

    if (docTitle) {
      bodyXml.push(
        buildParagraphXml({
          type: "paragraph",
          styleId: "Heading1",
          runs: [{ type: "text", text: docTitle }]
        })
      );
    }
    if (currentUrl) {
      bodyXml.push(
        buildParagraphXml({
          type: "paragraph",
          styleId: "Normal",
          runs: [
            { type: "text", text: "原文链接: " },
            { type: "text", text: currentUrl, href: currentUrl }
          ]
        })
      );
      bodyXml.push("<w:p/>");
    }

    for (const block of blocks) {
      if (block.type === "image") {
        const imageXml = buildImageParagraphXml(block);
        if (imageXml) {
          bodyXml.push(imageXml);
        }
      } else {
        bodyXml.push(buildParagraphXml(block));
      }
    }

    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="${OOXML_NS}" xmlns:r="${REL_NS}">
  <w:body>
    ${bodyXml.join("")}
    <w:sectPr>
      <w:pgSz w:w="11906" w:h="16838"/>
      <w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;
  }

  function buildDocumentRelsXml() {
    const relItems = [
      `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`
    ];
    for (const [href, relId] of hyperlinkRelMap.entries()) {
      relItems.push(
        `<Relationship Id="${escapeXmlAttr(relId)}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${escapeXmlAttr(
          href
        )}" TargetMode="External"/>`
      );
    }
    for (const media of mediaEntries) {
      relItems.push(
        `<Relationship Id="${escapeXmlAttr(
          media.relId
        )}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/${escapeXmlAttr(
          media.mediaName
        )}"/>`
      );
    }
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  ${relItems.join("")}
</Relationships>`;
  }

  return {
    buildDocumentXml,
    buildDocumentRelsXml,
    mediaEntries
  };
}

function buildStylesXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:styles xmlns:w="${OOXML_NS}">
  <w:docDefaults>
    <w:rPrDefault>
      <w:rPr>
        <w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="Microsoft YaHei"/>
        <w:sz w:val="22"/>
        <w:szCs w:val="22"/>
      </w:rPr>
    </w:rPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal">
    <w:name w:val="Normal"/>
    <w:qFormat/>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading1">
    <w:name w:val="heading 1"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="36"/>
      <w:szCs w:val="36"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading2">
    <w:name w:val="heading 2"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="30"/>
      <w:szCs w:val="30"/>
    </w:rPr>
  </w:style>
  <w:style w:type="paragraph" w:styleId="Heading3">
    <w:name w:val="heading 3"/>
    <w:basedOn w:val="Normal"/>
    <w:next w:val="Normal"/>
    <w:qFormat/>
    <w:rPr>
      <w:b/>
      <w:sz w:val="26"/>
      <w:szCs w:val="26"/>
    </w:rPr>
  </w:style>
  <w:style w:type="character" w:default="1" w:styleId="DefaultParagraphFont">
    <w:name w:val="Default Paragraph Font"/>
    <w:uiPriority w:val="1"/>
    <w:semiHidden/>
    <w:unhideWhenUsed/>
  </w:style>
  <w:style w:type="character" w:styleId="Hyperlink">
    <w:name w:val="Hyperlink"/>
    <w:basedOn w:val="DefaultParagraphFont"/>
    <w:uiPriority w:val="99"/>
    <w:unhideWhenUsed/>
    <w:rPr>
      <w:color w:val="0563C1"/>
      <w:u w:val="single"/>
    </w:rPr>
  </w:style>
</w:styles>`;
}

function buildPackageRelsXml() {
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="${PACKAGE_REL_NS}">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;
}

function buildContentTypesXml(mediaEntries) {
  const defaults = [
    '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>',
    '<Default Extension="xml" ContentType="application/xml"/>'
  ];
  const extSet = new Set();
  for (const entry of mediaEntries) {
    const ext = String(entry.ext || "").toLowerCase();
    if (!ext || extSet.has(ext)) continue;
    extSet.add(ext);
    defaults.push(
      `<Default Extension="${escapeXmlAttr(ext)}" ContentType="${escapeXmlAttr(
        entry.mimeType || MIME_BY_EXT[ext] || "application/octet-stream"
      )}"/>`
    );
  }
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="${CONTENT_TYPES_NS}">
  ${defaults.join("")}
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>
</Types>`;
}

export async function buildWordExport(payload, imageMap) {
  const baseName = buildExportBaseName(payload?.title || "feishu-document");
  const htmlWithPlaceholders = String(payload?.htmlWithPlaceholders || payload?.htmlContent || "");
  const blocks = parseHtmlToBlocks(htmlWithPlaceholders, payload?.currentUrl || payload?.url || "");
  const context = createWordContext(payload, imageMap || new Map());
  const documentXml = context.buildDocumentXml(blocks);
  const documentRelsXml = context.buildDocumentRelsXml();
  const contentTypesXml = buildContentTypesXml(context.mediaEntries);
  const stylesXml = buildStylesXml();
  const files = [
    { name: "[Content_Types].xml", data: contentTypesXml },
    { name: "_rels/.rels", data: buildPackageRelsXml() },
    { name: "word/document.xml", data: documentXml },
    { name: "word/styles.xml", data: stylesXml },
    { name: "word/_rels/document.xml.rels", data: documentRelsXml }
  ];

  for (const media of context.mediaEntries) {
    const safeName = sanitizeFilename(media.mediaName, "image.bin", 100);
    files.push({
      name: `word/media/${safeName}`,
      data: media.data
    });
  }

  const zipBlob = await createZipBlob(files);
  const blob = new Blob([zipBlob], { type: DOCX_MIME });
  return {
    blob,
    filename: ensureExtension(baseName, "docx"),
    exportType: "WORD"
  };
}
