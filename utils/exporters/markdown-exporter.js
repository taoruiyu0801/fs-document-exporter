import {
  buildExportBaseName,
  sanitizeFilename,
  ensureExtension,
  dataUrlToUint8Array,
  guessExtensionFromDataUrl
} from "../filename.js";
import { createZipBlob } from "../zip.js";

function decodeHtmlEntities(text) {
  return String(text || "")
    .replace(/&nbsp;/gi, " ")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&amp;/gi, "&")
    .replace(/&quot;/gi, "\"")
    .replace(/&#39;/gi, "'");
}

function stripTags(text) {
  return decodeHtmlEntities(String(text || "").replace(/<[^>]*>/g, ""));
}

function normalizeMarkdown(text) {
  return String(text || "")
    .replace(/\r/g, "")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function splitMarkdownForBigMode(markdown, maxChars = 4 * 1024 * 1024) {
  const text = String(markdown || "");
  if (!text || text.length <= maxChars) {
    return [text];
  }

  const parts = [];
  let start = 0;
  while (start < text.length) {
    let end = Math.min(start + maxChars, text.length);
    if (end < text.length) {
      const breakAt = text.lastIndexOf("\n## ", end);
      if (breakAt > start + 1024) {
        end = breakAt;
      }
    }
    parts.push(text.slice(start, end).trim());
    start = end;
  }
  return parts.filter(Boolean);
}

function convertHtmlToMarkdown(html, imageResolver) {
  let md = String(html || "");

  md = md
    .replace(/<script[\s\S]*?<\/script>/gi, "")
    .replace(/<style[\s\S]*?<\/style>/gi, "")
    .replace(/<!--[\s\S]*?-->/g, "");

  md = md.replace(
    /<pre\b[^>]*>\s*<code\b[^>]*>([\s\S]*?)<\/code>\s*<\/pre>/gi,
    (_, code) => `\n\`\`\`\n${decodeHtmlEntities(code).trim()}\n\`\`\`\n`
  );

  md = md.replace(/<h1\b[^>]*>([\s\S]*?)<\/h1>/gi, (_, c) => `\n# ${stripTags(c)}\n`);
  md = md.replace(/<h2\b[^>]*>([\s\S]*?)<\/h2>/gi, (_, c) => `\n## ${stripTags(c)}\n`);
  md = md.replace(/<h3\b[^>]*>([\s\S]*?)<\/h3>/gi, (_, c) => `\n### ${stripTags(c)}\n`);
  md = md.replace(/<h4\b[^>]*>([\s\S]*?)<\/h4>/gi, (_, c) => `\n#### ${stripTags(c)}\n`);
  md = md.replace(/<h5\b[^>]*>([\s\S]*?)<\/h5>/gi, (_, c) => `\n##### ${stripTags(c)}\n`);
  md = md.replace(/<h6\b[^>]*>([\s\S]*?)<\/h6>/gi, (_, c) => `\n###### ${stripTags(c)}\n`);

  md = md.replace(/<blockquote\b[^>]*>([\s\S]*?)<\/blockquote>/gi, (_, c) => {
    const lines = stripTags(c).split(/\n+/).map((line) => `> ${line.trim()}`);
    return `\n${lines.join("\n")}\n`;
  });

  md = md.replace(/<li\b[^>]*>([\s\S]*?)<\/li>/gi, (_, c) => `\n- ${stripTags(c)}`);
  md = md.replace(/<\/?(ul|ol)\b[^>]*>/gi, "\n");

  md = md.replace(/<strong\b[^>]*>([\s\S]*?)<\/strong>/gi, (_, c) => `**${stripTags(c)}**`);
  md = md.replace(/<b\b[^>]*>([\s\S]*?)<\/b>/gi, (_, c) => `**${stripTags(c)}**`);
  md = md.replace(/<em\b[^>]*>([\s\S]*?)<\/em>/gi, (_, c) => `*${stripTags(c)}*`);
  md = md.replace(/<i\b[^>]*>([\s\S]*?)<\/i>/gi, (_, c) => `*${stripTags(c)}*`);
  md = md.replace(/<code\b[^>]*>([\s\S]*?)<\/code>/gi, (_, c) => `\`${stripTags(c)}\``);

  md = md.replace(
    /<a\b[^>]*href="([^"]*)"[^>]*>([\s\S]*?)<\/a>/gi,
    (_, href, text) => `[${stripTags(text)}](${href})`
  );

  md = md.replace(
    /<img\b[^>]*data-token="([^"]+)"[^>]*>/gi,
    (full, token) => {
      const nameMatch = /data-name="([^"]*)"/i.exec(full);
      const altText = nameMatch ? nameMatch[1] : token;
      return `\n${imageResolver(token, altText)}\n`;
    }
  );

  md = md.replace(
    /<img\b[^>]*src="([^"]+)"[^>]*>/gi,
    (_, src) => `\n![](${src})\n`
  );

  md = md.replace(/<br\s*\/?>/gi, "\n");
  md = md.replace(/<\/p>/gi, "\n\n");
  md = md.replace(/<p\b[^>]*>/gi, "");
  md = md.replace(/<\/div>/gi, "\n");
  md = md.replace(/<div\b[^>]*>/gi, "");
  md = md.replace(/<[^>]*>/g, "");

  return normalizeMarkdown(decodeHtmlEntities(md));
}

export async function buildMarkdownExport(payload, imageMap, options = {}) {
  const baseName = buildExportBaseName(payload.title || "feishu-document");
  const imageFiles = [];
  const imagePathByToken = new Map();
  const addedImagePath = new Set();

  const markdown = convertHtmlToMarkdown(
    payload.htmlWithPlaceholders,
    (token, altText) => {
      const dataUrl = imageMap.get(token);
      if (!dataUrl) {
        return `![${altText}](missing://${token})`;
      }
      const existingPath = imagePathByToken.get(token);
      if (existingPath) {
        return `![${altText}](${existingPath})`;
      }
      const ext = guessExtensionFromDataUrl(dataUrl, "png");
      const safeToken = sanitizeFilename(token, "image");
      const fileName = `images/${safeToken}.${ext}`;
      imagePathByToken.set(token, fileName);
      if (!addedImagePath.has(fileName)) {
        addedImagePath.add(fileName);
        imageFiles.push({
          name: fileName,
          data: dataUrlToUint8Array(dataUrl)
        });
      }
      return `![${altText}](${fileName})`;
    }
  );

  const splitParts = options.bigMode
    ? splitMarkdownForBigMode(markdown)
    : [markdown];
  const markdownFiles = splitParts.map((part, index) => ({
    name:
      splitParts.length > 1
        ? `${baseName}.part-${String(index + 1).padStart(3, "0")}.md`
        : ensureExtension(baseName, "md"),
    data: `${part}\n`
  }));
  const zipBlob = await createZipBlob([...markdownFiles, ...imageFiles]);
  return {
    blob: zipBlob,
    filename: ensureExtension(baseName, "zip"),
    exportType: "MARKDOWN"
  };
}
