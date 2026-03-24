#!/usr/bin/env node
/**
 * ChatGPT Conversation Exporter v2 — Node.js with local web UI.
 * Starts a local server, opens a browser with a nice UI,
 * user pastes their session JSON, and conversations are exported.
 *
 * Improvements over v1:
 *   • Resume:    skips conversations whose .json file already exists
 *   • Speed:     5 parallel downloads + 100ms delay (was sequential 500ms)
 *   • Projects:  exports ChatGPT project conversations to projects/{name}/
 *   • No deps:   only built-in Node.js modules
 */

import { createServer }                       from "node:http";
import { writeFileSync, mkdirSync, existsSync, readFileSync, appendFileSync } from "node:fs";
import { join, dirname }                      from "node:path";
import { homedir }                            from "node:os";
import { randomUUID }                         from "node:crypto";
import { Buffer }                             from "node:buffer";
import { execSync }                           from "node:child_process";
import { fileURLToPath }                      from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));

const API_BASE   = "https://chatgpt.com/backend-api";
const PAGE_SIZE  = 100;
const DELAY      = 100;   // ms between requests per slot (was 500ms)
const HOST       = "127.0.0.1";
const PORT       = 8523;
const DEVICE_ID  = randomUUID();

const HEADERS = {
  "Content-Type":    "application/json",
  "User-Agent":      "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36",
  Accept:            "application/json",
  "Accept-Language": "en-US,en;q=0.9",
  Referer:           "https://chatgpt.com/",
  Origin:            "https://chatgpt.com",
  "Oai-Device-Id":   DEVICE_ID,
  "Oai-Language":    "en-US",
  "Sec-Ch-Ua":       '"Google Chrome";v="131", "Chromium";v="131", "Not_A Brand";v="24"',
  "Sec-Ch-Ua-Mobile":   "?0",
  "Sec-Ch-Ua-Platform": '"macOS"',
  "Sec-Fetch-Dest":     "empty",
  "Sec-Fetch-Mode":     "cors",
  "Sec-Fetch-Site":     "same-origin",
};

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

// ── API helpers ──────────────────────────────────────────────────────

async function apiGet(path, token) {
  const resp = await fetch(`${API_BASE}/${path}`, {
    headers: { ...HEADERS, Authorization: `Bearer ${token}` },
  });
  if (!resp.ok) {
    const body = await resp.text().catch(() => "");
    throw new Error(`HTTP ${resp.status}: ${body.slice(0, 300)}`);
  }
  return resp.json();
}

async function apiFetchBinary(url, token) {
  const h = { ...HEADERS, Accept: "*/*" };
  if (token) h.Authorization = `Bearer ${token}`;
  const resp = await fetch(url, { headers: h });
  if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
  const buffer = Buffer.from(await resp.arrayBuffer());
  const contentType = resp.headers.get("content-type") || "";
  return { buffer, contentType };
}

const MIME_TO_EXT = {
  "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
  "image/webp": ".webp", "image/svg+xml": ".svg", "application/pdf": ".pdf",
  "text/plain": ".txt", "text/html": ".html", "text/csv": ".csv",
  "application/json": ".json", "application/zip": ".zip",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
  "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
};

// ── Semaphore ────────────────────────────────────────────────────────

class Semaphore {
  constructor(n) { this.n = n; this.queue = []; }
  acquire() {
    if (this.n > 0) { this.n--; return Promise.resolve(); }
    return new Promise((r) => this.queue.push(r));
  }
  release() {
    if (this.queue.length > 0) { this.queue.shift()(); } else { this.n++; }
  }
}

// ── File utilities ───────────────────────────────────────────────────

function sanitizeFilename(name, maxLen = 80) {
  return name.replace(/[<>:"/\\|?*]/g, "_").replace(/^[. ]+|[. ]+$/g, "").slice(0, maxLen) || "untitled";
}

function stripCitations(str) {
  return str.replace(/\u3010[^\u3011]*\u3011/g, "");
}

function escapeHtml(str) {
  return str.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

function deduplicateFilename(name, usedNames) {
  if (!usedNames.has(name)) { usedNames.add(name); return name; }
  const dot  = name.lastIndexOf(".");
  const base = dot > 0 ? name.slice(0, dot) : name;
  const ext  = dot > 0 ? name.slice(dot)    : "";
  let i = 1;
  while (usedNames.has(`${base}_${i}${ext}`)) i++;
  const deduped = `${base}_${i}${ext}`;
  usedNames.add(deduped);
  return deduped;
}

// ── File reference extraction ────────────────────────────────────────

function extractFileReferences(convo) {
  const refs = [];
  const seen = new Set();
  const mapping = convo.mapping || {};
  for (const node of Object.values(mapping)) {
    const msg = node.message;
    if (!msg) continue;
    if (msg.content?.parts) {
      for (const part of msg.content.parts) {
        if (part?.content_type === "image_asset_pointer" && part.asset_pointer) {
          const match = part.asset_pointer.match(/^(?:file-service|sediment):\/\/(.+)$/);
          if (match && !seen.has(match[1])) {
            seen.add(match[1]);
            refs.push({ fileId: match[1], filename: part.metadata?.dalle?.prompt ? "dalle_image.png" : "image.png", type: "image" });
          }
        }
      }
    }
    if (msg.metadata?.attachments) {
      for (const att of msg.metadata.attachments) {
        if (att.id && !seen.has(att.id)) {
          seen.add(att.id);
          refs.push({ fileId: att.id, filename: att.name || "attachment", type: "attachment" });
        }
      }
    }
    if (msg.metadata?.citations) {
      for (const cit of msg.metadata.citations) {
        const fileId = cit.metadata?.file_id || cit.file_id;
        const title  = cit.metadata?.title    || cit.title   || "citation";
        if (fileId && !seen.has(fileId)) {
          seen.add(fileId);
          refs.push({ fileId, filename: title, type: "citation" });
        }
      }
    }
  }
  return refs;
}

async function downloadFile(fileId, token, fallbackName) {
  const meta = await apiGet(`files/download/${fileId}`, token);
  const url = meta.download_url;
  if (!url) throw new Error("No download_url returned");
  const { buffer, contentType } = await apiFetchBinary(url, token);
  let filename = meta.file_name || fallbackName || fileId;
  if (!filename.includes(".") && contentType) {
    const mime = contentType.split(";")[0].trim();
    const ext  = MIME_TO_EXT[mime];
    if (ext) filename += ext;
  }
  return { filename, buffer };
}

// ── Markdown converter ───────────────────────────────────────────────

function conversationToMarkdown(convo, fileMap = {}) {
  const title   = convo.title || "Untitled";
  const ct      = convo.create_time;
  const dateStr = ct
    ? new Date(ct * 1000).toISOString().replace("T", " ").slice(0, 16) + " UTC"
    : "";

  const lines   = [`# ${title}`, ""];
  if (dateStr) lines.push(`*${dateStr}*\n`);

  const mapping = convo.mapping || {};
  const rootId  = Object.keys(mapping).find((k) => mapping[k].parent == null);

  if (rootId) {
    const queue = [rootId];
    while (queue.length) {
      const nid  = queue.shift();
      const node = mapping[nid] || {};
      const msg  = node.message;
      if (msg?.content?.parts) {
        const role        = msg.author?.role || "unknown";
        const contentType = msg.content?.content_type || "text";
        if (role === "system" || role === "tool") { queue.push(...(node.children || [])); continue; }
        if (role === "assistant" && contentType !== "text") { queue.push(...(node.children || [])); continue; }

        const textParts = [];
        for (const part of msg.content.parts) {
          if (typeof part === "string") {
            textParts.push(part);
          } else if (part?.content_type === "image_asset_pointer" && part.asset_pointer) {
            const match = part.asset_pointer.match(/^(?:file-service|sediment):\/\/(.+)$/);
            textParts.push(match && fileMap[match[1]] ? `![image](${fileMap[match[1]]})` : "[image]");
          } else {
            textParts.push(JSON.stringify(part));
          }
        }
        if (msg.metadata?.attachments) {
          for (const att of msg.metadata.attachments) {
            if (att.id && fileMap[att.id]) {
              textParts.push(`\n📎 [${att.name || "attachment"}](${fileMap[att.id]})`);
            }
          }
        }
        const text = stripCitations(textParts.join("\n")).trim();
        if (text) lines.push(`## ${role.charAt(0).toUpperCase() + role.slice(1)}\n\n${text}\n`);
      }
      queue.push(...(node.children || []));
    }
  }
  return lines.join("\n");
}

// ── HTML converter ───────────────────────────────────────────────────

function conversationToHtml(convo, fileMap = {}, allConversations = [], currentFname = "") {
  const title   = escapeHtml(convo.title || "Untitled");
  const ct      = convo.create_time;
  const dateStr = ct
    ? new Date(ct * 1000).toISOString().replace("T", " ").slice(0, 16) + " UTC"
    : "";

  const messages = [];
  const mapping  = convo.mapping || {};
  const rootId   = Object.keys(mapping).find((k) => mapping[k].parent == null);

  if (rootId) {
    const queue = [rootId];
    while (queue.length) {
      const nid  = queue.shift();
      const node = mapping[nid] || {};
      const msg  = node.message;
      if (msg?.content?.parts) {
        const role        = msg.author?.role || "unknown";
        const contentType = msg.content?.content_type || "text";
        if (role === "system") { queue.push(...(node.children || [])); continue; }

        const isInternal = role === "tool" ||
          (role === "assistant" && contentType !== "text") ||
          (role === "user" && contentType === "user_editable_context");

        const textParts  = [];
        const imageParts = [];
        for (const part of msg.content.parts) {
          if (typeof part === "string") {
            textParts.push(part);
          } else if (part?.content_type === "image_asset_pointer" && part.asset_pointer) {
            const match = part.asset_pointer.match(/^(?:file-service|sediment):\/\/(.+)$/);
            if (match && fileMap[match[1]]) imageParts.push(fileMap[match[1]]);
          }
        }
        const attachments = [];
        if (msg.metadata?.attachments) {
          for (const att of msg.metadata.attachments) {
            if (att.id && fileMap[att.id]) {
              attachments.push({ name: att.name || "attachment", path: fileMap[att.id] });
            }
          }
        }
        const text = stripCitations(textParts.join("\n")).trim();
        if (text || imageParts.length || attachments.length) {
          messages.push({ role, text, images: imageParts, attachments, isInternal, contentType });
        }
      }
      queue.push(...(node.children || []));
    }
  }

  const LOGO = `<svg viewBox="0 0 41 41" fill="none" xmlns="http://www.w3.org/2000/svg" width="24" height="24"><path d="M37.532 16.87a9.963 9.963 0 0 0-.856-8.184 10.078 10.078 0 0 0-10.855-4.835A9.964 9.964 0 0 0 18.306.5a10.079 10.079 0 0 0-9.614 6.977 9.967 9.967 0 0 0-6.664 4.834 10.08 10.08 0 0 0 1.24 11.817 9.965 9.965 0 0 0 .856 8.185 10.079 10.079 0 0 0 10.855 4.835 9.965 9.965 0 0 0 7.516 3.35 10.078 10.078 0 0 0 9.617-6.981 9.967 9.967 0 0 0 6.663-4.834 10.079 10.079 0 0 0-1.243-11.813ZM22.498 37.886a7.474 7.474 0 0 1-4.799-1.735c.061-.033.168-.091.237-.134l7.964-4.6a1.294 1.294 0 0 0 .655-1.134V19.054l3.366 1.944a.12.12 0 0 1 .066.092v9.299a7.505 7.505 0 0 1-7.49 7.496ZM6.392 31.006a7.471 7.471 0 0 1-.894-5.023c.06.036.162.099.237.141l7.964 4.6a1.297 1.297 0 0 0 1.308 0l9.724-5.614v3.888a.12.12 0 0 1-.048.103l-8.051 4.649a7.504 7.504 0 0 1-10.24-2.744ZM4.297 13.62A7.469 7.469 0 0 1 8.2 10.333c0 .068-.004.19-.004.274v9.201a1.294 1.294 0 0 0 .654 1.132l9.723 5.614-3.366 1.944a.12.12 0 0 1-.114.012L7.044 23.86a7.504 7.504 0 0 1-2.747-10.24Zm27.658 6.437-9.724-5.615 3.367-1.943a.121.121 0 0 1 .114-.012l8.048 4.648a7.498 7.498 0 0 1-1.158 13.528V21.36a1.293 1.293 0 0 0-.647-1.132v-.17Zm3.35-5.043c-.059-.037-.162-.099-.236-.141l-7.965-4.6a1.298 1.298 0 0 0-1.308 0l-9.723 5.614v-3.888a.12.12 0 0 1 .048-.103l8.05-4.645a7.497 7.497 0 0 1 11.135 7.763Zm-21.063 6.929-3.367-1.944a.12.12 0 0 1-.065-.092v-9.299a7.497 7.497 0 0 1 12.293-5.756 6.94 6.94 0 0 0-.236.134l-7.965 4.6a1.294 1.294 0 0 0-.654 1.132l-.006 11.225Zm1.829-3.943 4.33-2.501 4.332 2.5v5l-4.331 2.5-4.331-2.5V18Z" fill="currentColor"/></svg>`;

  const INTERNAL_LABELS = {
    multimodal_text: "File context", code: "Code", execution_output: "Output",
    computer_output: "Output", tether_browsing_display: "Web browsing",
    system_error: "Error", text: "Tool output",
  };

  const messagesHtml = messages.map((m) => {
    if (m.isInternal) {
      const label = INTERNAL_LABELS[m.contentType] || "Internal context";
      const b64   = Buffer.from(m.text, "utf8").toString("base64");
      return `<details class="thinking"><summary>${label}</summary><div class="thinking-content md-content" dir="auto" data-md="${b64}"></div></details>`;
    }
    const roleClass = m.role === "user" ? "user" : "assistant";
    let content = "";
    if (m.role === "user") {
      content = `<div class="bubble" dir="auto">${escapeHtml(m.text).replace(/\n/g, "<br>")}</div>`;
    } else {
      const b64 = Buffer.from(m.text, "utf8").toString("base64");
      content = `<div class="avatar">${LOGO}</div><div class="content"><div class="md-content" dir="auto" data-md="${b64}"></div></div>`;
    }
    if (m.images.length) {
      content += `<div class="images">${m.images.map((s) =>
        `<a href="${escapeHtml(s)}" target="_blank"><img src="${escapeHtml(s)}" alt="image" loading="lazy"></a>`
      ).join("")}</div>`;
    }
    if (m.attachments.length) {
      content += `<div class="attachments">${m.attachments.map((a) =>
        `<a class="attachment" href="${escapeHtml(a.path)}" target="_blank"><span class="att-icon">📎</span><span class="att-name">${escapeHtml(a.name)}</span></a>`
      ).join("")}</div>`;
    }
    return `<div class="message ${roleClass}">${content}</div>`;
  }).join("\n");

  const sidebarItems = allConversations.map((c) => {
    const cls = c.fname === currentFname ? "sidebar-item active" : "sidebar-item";
    return `<a class="${cls}" href="${c.fname}.html" title="${escapeHtml(c.title)}">${escapeHtml(c.title)}</a>`;
  }).join("\n");

  return `<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/highlightjs/cdn-release/build/styles/github-dark.min.css">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif; background: #fff; color: #0d0d0d; line-height: 1.65; font-size: 16px; display: flex; height: 100vh; }
  .sidebar { width: 260px; min-width: 260px; height: 100vh; background: #f9f9f9; border-right: 1px solid #e5e5e5; overflow-y: auto; padding: 16px 0; flex-shrink: 0; position: sticky; top: 0; }
  .sidebar-header { padding: 8px 16px 16px; font-size: 14px; font-weight: 600; color: #6b6b6b; border-bottom: 1px solid #e5e5e5; margin-bottom: 8px; }
  .sidebar-item { display: block; padding: 8px 16px; font-size: 13px; color: #0d0d0d; text-decoration: none; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; border-radius: 8px; margin: 2px 8px; }
  .sidebar-item:hover { background: #ececec; }
  .sidebar-item.active { background: #e5e5e5; font-weight: 600; }
  .main { flex: 1; overflow-y: auto; }
  .header { max-width: 768px; margin: 0 auto; padding: 32px 24px 16px; border-bottom: 1px solid #e5e5e5; }
  .header h1 { font-size: 22px; font-weight: 600; }
  .header .date { font-size: 13px; color: #6b6b6b; margin-top: 4px; }
  .chat { max-width: 768px; margin: 0 auto; padding: 24px; }
  .message { margin-bottom: 24px; }
  .message.user { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px; }
  .message.user .bubble { background: #f4f4f4; border-radius: 18px; padding: 10px 16px; max-width: 85%; white-space: pre-wrap; word-break: break-word; }
  .message.assistant { display: flex; gap: 12px; align-items: flex-start; }
  .message.assistant .avatar { width: 28px; height: 28px; border-radius: 50%; background: #00a67e; color: #fff; display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 2px; }
  .message.assistant .content { flex: 1; min-width: 0; }
  .message.assistant .content h1, .message.assistant .content h2, .message.assistant .content h3 { margin: 16px 0 8px; font-weight: 600; }
  .message.assistant .content h1 { font-size: 20px; } .message.assistant .content h2 { font-size: 18px; } .message.assistant .content h3 { font-size: 16px; }
  .message.assistant .content p { margin: 8px 0; }
  .message.assistant .content ul, .message.assistant .content ol { margin: 8px 0; padding-left: 24px; }
  .message.assistant .content li { margin: 4px 0; }
  .message.assistant .content a { color: #1a7f64; }
  .message.assistant .content code { background: #f0f0f0; border-radius: 4px; padding: 2px 5px; font-family: monospace; font-size: 14px; }
  .message.assistant .content pre { margin: 12px 0; border-radius: 8px; overflow: hidden; }
  .message.assistant .content pre code { display: block; background: #0d0d0d; color: #f8f8f2; padding: 16px; overflow-x: auto; font-size: 13px; line-height: 1.5; }
  .code-block { position: relative; }
  .code-block .copy-btn { position: absolute; top: 8px; right: 8px; background: #333; border: none; color: #999; cursor: pointer; font-size: 12px; padding: 4px 10px; border-radius: 4px; opacity: 0; transition: opacity 0.2s; }
  .code-block:hover .copy-btn { opacity: 1; }
  .images img { max-width: 100%; border-radius: 8px; margin: 4px 0; display: block; }
  .message.user .images img { max-width: 300px; }
  .attachments { margin-top: 8px; display: flex; flex-wrap: wrap; gap: 8px; }
  .attachment { display: inline-flex; align-items: center; gap: 8px; background: #f4f4f4; border: 1px solid #e5e5e5; border-radius: 8px; padding: 8px 12px; text-decoration: none; color: #0d0d0d; font-size: 14px; }
  .att-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 200px; }
  .thinking { margin-bottom: 24px; border-left: 3px solid #d4d4d4; padding-left: 16px; font-size: 14px; }
  .thinking summary { color: #8e8e8e; font-style: italic; cursor: pointer; padding: 4px 0; }
  .thinking-content { color: #6b6b6b; padding: 8px 0; font-style: italic; }
</style>
</head>
<body>
<nav class="sidebar"><div class="sidebar-header">Conversations</div>${sidebarItems}</nav>
<div class="main">
  <div class="header"><h1>${title}</h1>${dateStr ? `<div class="date">${dateStr}</div>` : ""}</div>
  <div class="chat">${messagesHtml}</div>
</div>
<script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/gh/highlightjs/cdn-release/build/highlight.min.js"><\/script>
<script>
document.addEventListener('DOMContentLoaded', () => {
  const renderer = new marked.Renderer();
  renderer.code = function({ text, lang }) {
    const hl = lang && hljs.getLanguage(lang) ? hljs.highlight(text, { language: lang }).value : hljs.highlightAuto(text).value;
    return '<div class="code-block"><button class="copy-btn" onclick="navigator.clipboard.writeText(this.nextElementSibling.querySelector(\\'code\\').textContent);this.textContent=\\'Copied!\\';setTimeout(()=>this.textContent=\\'Copy\\',1500)">Copy</button><pre><code class="hljs">' + hl + '</code></pre></div>';
  };
  marked.use({ renderer, breaks: true });
  document.querySelectorAll('.md-content').forEach(el => {
    try { el.innerHTML = marked.parse(decodeURIComponent(escape(atob(el.dataset.md)))); } catch(e) {}
  });
  const active = document.querySelector('.sidebar-item.active');
  if (active) active.scrollIntoView({ block: 'center', behavior: 'instant' });
});
<\/script>
</body>
</html>`;
}

// ── ZIP builder ──────────────────────────────────────────────────────

function crc32(buf) {
  const table = new Uint32Array(256);
  for (let i = 0; i < 256; i++) {
    let c = i;
    for (let j = 0; j < 8; j++) c = c & 1 ? 0xedb88320 ^ (c >>> 1) : c >>> 1;
    table[i] = c;
  }
  let crc = 0xffffffff;
  for (let i = 0; i < buf.length; i++) crc = table[(crc ^ buf[i]) & 0xff] ^ (crc >>> 8);
  return (crc ^ 0xffffffff) >>> 0;
}

function buildZip(files) {
  const entries = [];
  let offset = 0;
  for (const file of files) {
    const pathBuf = Buffer.from(file.path, "utf8");
    const data    = Buffer.isBuffer(file.data) ? file.data : Buffer.from(file.data, "utf8");
    const crc     = crc32(data);
    const header  = Buffer.alloc(30);
    header.writeUInt32LE(0x04034b50, 0);
    header.writeUInt16LE(20, 4);
    header.writeUInt32LE(crc, 14);
    header.writeUInt32LE(data.length, 18);
    header.writeUInt32LE(data.length, 22);
    header.writeUInt16LE(pathBuf.length, 26);
    entries.push({ header, pathBuf, data, crc, offset });
    offset += 30 + pathBuf.length + data.length;
  }
  const cdParts = [];
  for (const e of entries) {
    const cd = Buffer.alloc(46);
    cd.writeUInt32LE(0x02014b50, 0); cd.writeUInt16LE(20, 4); cd.writeUInt16LE(20, 6);
    cd.writeUInt32LE(e.crc, 16); cd.writeUInt32LE(e.data.length, 20); cd.writeUInt32LE(e.data.length, 24);
    cd.writeUInt16LE(e.pathBuf.length, 28); cd.writeUInt32LE(e.offset, 42);
    cdParts.push(cd, e.pathBuf);
  }
  const cdBuf = Buffer.concat(cdParts);
  const eocd  = Buffer.alloc(22);
  eocd.writeUInt32LE(0x06054b50, 0); eocd.writeUInt16LE(entries.length, 8); eocd.writeUInt16LE(entries.length, 10);
  eocd.writeUInt32LE(cdBuf.length, 12); eocd.writeUInt32LE(offset, 16);
  const parts = [];
  for (const e of entries) parts.push(e.header, e.pathBuf, e.data);
  parts.push(cdBuf, eocd);
  return Buffer.concat(parts);
}

// ── GPT markdown converter ────────────────────────────────────────────

function gptToMarkdown(gizmo) {
  const display  = gizmo.display || {};
  const name     = display.name        || gizmo.name || "Untitled GPT";
  const desc     = display.description || "";
  const starters = display.prompt_starters || [];
  const tools    = gizmo.tools || [];
  const created  = gizmo.created_at
    ? new Date(gizmo.created_at * 1000).toISOString().slice(0, 10)
    : "";
  const visibility = gizmo.visibility || display.visibility || "";

  let md = `# ${name}\n\n`;
  if (desc)       md += `**Description:** ${desc}\n\n`;
  if (created)    md += `**Created:** ${created}\n\n`;
  if (visibility) md += `**Visibility:** ${visibility}\n\n`;

  if (gizmo.instructions) {
    md += `## Instructions\n\n${gizmo.instructions}\n\n`;
  }

  if (starters.length) {
    md += `## Conversation Starters\n\n`;
    for (const s of starters) md += `- ${s}\n`;
    md += "\n";
  }

  if (tools.length) {
    md += `## Tools\n\n`;
    for (const t of tools) md += `- ${t.type || t.name || JSON.stringify(t)}\n`;
    md += "\n";
  }

  return md;
}

// ── Export logic ─────────────────────────────────────────────────────

async function runExport(token, outputDir, concurrency, sendEvent) {
  try {
    const CONCURRENCY = Math.max(1, Math.min(10, concurrency || 5));
    const rootDir     = outputDir.startsWith("~")
      ? join(homedir(), outputDir.slice(1))
      : outputDir;
    const zipPath     = rootDir + ".zip";

    // ── Logging setup ─────────────────────────────────────────────────
    mkdirSync(rootDir, { recursive: true });
    const logTs   = new Date().toISOString().replace(/[:.]/g, "-").slice(0, 19);
    const logPath = join(rootDir, `export-${logTs}.log`);
    const log = (msg) => appendFileSync(logPath, `[${new Date().toISOString()}] ${msg}\n`, "utf8");

    log(`Export started — output: ${rootDir} — concurrency: ${CONCURRENCY}`);

    const _send = sendEvent;
    sendEvent = (type, data) => {
      _send(type, data);
      if      (type === "status")   log(`STATUS   ${data}`);
      else if (type === "progress") { const d = JSON.parse(data); log(`PROGRESS ${d.current}/${d.total} — ${d.title}`); }
      else if (type === "done")     { const d = JSON.parse(data); log(`DONE     ${d.succeeded} new · ${d.skipped} resumed · ${d.failed} failed · output: ${d.output}`); }
      else if (type === "error_msg") log(`ERROR    ${data}`);
    };

    // ── Fetch conversations ──────────────────────────────────────────
    sendEvent("status", "Fetching conversation list...");
    let allConversations = [];

    let offset = 0;
    while (true) {
      const data  = await apiGet(`conversations?offset=${offset}&limit=${PAGE_SIZE}`, token);
      const items = data.items || [];
      if (!items.length) break;
      allConversations.push(...items.map((c) => ({ ...c, projectName: null })));
      const total = data.total || 0;
      sendEvent("status", `Fetching conversation list... ${allConversations.length}/${total}`);
      offset += PAGE_SIZE;
      if (offset >= total) break;
      await sleep(DELAY);
    }

    // ── Fetch projects ───────────────────────────────────────────────
    try {
      sendEvent("status", "Fetching projects...");
      let projOffset = 0;
      const allProjects = [];
      while (true) {
        const data  = await apiGet(`projects?offset=${projOffset}&limit=50`, token);
        const items = data.projects || data.items || (Array.isArray(data) ? data : []);
        if (!items.length) break;
        allProjects.push(...items);
        projOffset += items.length;
        if (!data.has_more && !data.next) break;
        await sleep(DELAY);
      }

      const seenIds = new Set(allConversations.map((c) => c.id));
      for (const proj of allProjects) {
        const projName = proj.name || proj.title || proj.id;
        sendEvent("status", `Fetching project: ${projName}...`);
        let off = 0;
        while (true) {
          const data  = await apiGet(`projects/${proj.id}/conversations?offset=${off}&limit=${PAGE_SIZE}`, token);
          const items = data.items || [];
          if (!items.length) break;
          for (const c of items) {
            if (!seenIds.has(c.id)) {
              seenIds.add(c.id);
              allConversations.push({ ...c, projectName: projName });
            } else {
              const ex = allConversations.find((x) => x.id === c.id);
              if (ex) ex.projectName = projName;
            }
          }
          const total = data.total || items.length;
          off += PAGE_SIZE;
          if (off >= total) break;
          await sleep(DELAY);
        }
      }
    } catch {
      // Projects endpoint may not exist — continue without it
    }

    const total = allConversations.length;
    if (total === 0) {
      sendEvent("done", JSON.stringify({ total: 0, succeeded: 0, skipped: 0, failed: 0, output: rootDir }));
      return;
    }

    sendEvent("status", `Found ${total} conversations. Starting download...`);

    const jsonDir  = join(rootDir, "json");
    const mdDir    = join(rootDir, "markdown");
    const htmlDir  = join(rootDir, "html");
    const filesDir = join(rootDir, "files");
    mkdirSync(jsonDir,  { recursive: true });
    mkdirSync(mdDir,    { recursive: true });
    mkdirSync(htmlDir,  { recursive: true });
    mkdirSync(filesDir, { recursive: true });

    // Pre-compute conversation list per group so HTML sidebar is correct
    // even when generating HTML inline during download.
    const convosByGroup = {};
    for (const item of allConversations) {
      const key = item.projectName
        ? `projects/${sanitizeFilename(item.projectName)}`
        : "__root__";
      if (!convosByGroup[key]) convosByGroup[key] = [];
      const s = sanitizeFilename(item.title || "Untitled");
      convosByGroup[key].push({ fname: `${s}_${item.id.slice(0, 8)}`, title: item.title || "Untitled" });
    }

    const zipFiles   = [];
    const failed     = [];
    let totalFiles   = 0;
    let failedFiles  = 0;
    const failedFileDetails = [];
    let newCount     = 0;
    let skipped      = 0;
    const startTime  = Date.now();

    const sem = new Semaphore(CONCURRENCY);

    const tasks = allConversations.map((item) => async () => {
      await sem.acquire();
      try {
        const { id: cid, title: rawTitle, projectName } = item;
        const title  = rawTitle || "Untitled";
        const safe   = sanitizeFilename(title);
        const fname  = `${safe}_${cid.slice(0, 8)}`;

        const baseDirParts = projectName ? ["projects", sanitizeFilename(projectName)] : [];
        const convJsonDir  = baseDirParts.length ? join(rootDir, ...baseDirParts, "json")     : jsonDir;
        const convMdDir    = baseDirParts.length ? join(rootDir, ...baseDirParts, "markdown") : mdDir;
        const convHtmlDir  = baseDirParts.length ? join(rootDir, ...baseDirParts, "html")     : htmlDir;
        const convFilesDir = baseDirParts.length ? join(rootDir, ...baseDirParts, "files")    : filesDir;
        const zipPrefix    = baseDirParts.length ? baseDirParts.join("/") + "/" : "";

        const jsonPath = join(convJsonDir, `${fname}.json`);

        // Resume: skip without sleeping — no API call was made
        if (existsSync(jsonPath)) {
          skipped++;
          const done = newCount + skipped + failed.length;
          // Throttle progress events for skipped (every 25 or last item)
          if (skipped % 25 === 0 || done === total) {
            const elapsed = (Date.now() - startTime) / 1000 / 60;
            const rate    = elapsed > 0 ? Math.round(done / elapsed) : 0;
            sendEvent("progress", JSON.stringify({
              current: done, total, title,
              stats: `${newCount} new · ${skipped} resumed · ${failed.length} failed · ${rate}/min`,
            }));
          }
          return;
        }

        const convo   = await apiGet(`conversation/${cid}`, token);
        const jsonStr = JSON.stringify(convo, null, 2);

        // Download attached files
        const fileRefs  = extractFileReferences(convo);
        const fileMap   = {};
        const usedNames = new Set();

        if (fileRefs.length) {
          const convFilesDirFull = join(convFilesDir, fname);
          mkdirSync(convFilesDirFull, { recursive: true });
          for (const ref of fileRefs) {
            totalFiles++;
            try {
              const { filename: dlName, buffer } = await downloadFile(ref.fileId, token, ref.filename);
              const actualName = deduplicateFilename(dlName || ref.filename, usedNames);
              writeFileSync(join(convFilesDirFull, actualName), buffer);
              zipFiles.push({ path: `${zipPrefix}files/${fname}/${actualName}`, data: buffer });
              fileMap[ref.fileId] = `../files/${fname}/${actualName}`;
              await sleep(DELAY);
            } catch (e) {
              failedFiles++;
              failedFileDetails.push(`${ref.filename} (${ref.fileId}): ${e.message}`);
            }
          }
        }

        // Write JSON + Markdown (markdown computed once, reused for ZIP)
        const mdStr = conversationToMarkdown(convo, fileMap);
        mkdirSync(convJsonDir, { recursive: true });
        mkdirSync(convMdDir,   { recursive: true });
        writeFileSync(jsonPath, jsonStr, "utf8");
        writeFileSync(join(convMdDir, `${fname}.md`), mdStr, "utf8");
        zipFiles.push({ path: `${zipPrefix}json/${fname}.json`,     data: jsonStr });
        zipFiles.push({ path: `${zipPrefix}markdown/${fname}.md`,   data: mdStr   });

        // Write HTML inline using pre-computed sidebar list
        try {
          const groupKey = baseDirParts.length ? baseDirParts.join("/") : "__root__";
          mkdirSync(convHtmlDir, { recursive: true });
          const htmlStr = conversationToHtml(convo, fileMap, convosByGroup[groupKey] || [], fname);
          writeFileSync(join(convHtmlDir, `${fname}.html`), htmlStr, "utf8");
          zipFiles.push({ path: `${zipPrefix}html/${fname}.html`, data: htmlStr });
        } catch { /* non-fatal */ }

        newCount++;
        const done    = newCount + skipped + failed.length;
        const elapsed = (Date.now() - startTime) / 1000 / 60;
        const rate    = elapsed > 0 ? Math.round(done / elapsed) : 0;
        sendEvent("progress", JSON.stringify({
          current: done, total, title,
          stats: `${newCount} new · ${skipped} resumed · ${failed.length} failed · ${rate}/min`,
        }));

        await sleep(DELAY); // rate-limit only after an actual API call

      } catch (e) {
        failed.push(item.title || "Untitled");
        console.error(`[error] "${item.title}": ${e.message}`);
        log(`FAILED   "${item.title}": ${e.message}`);
      } finally {
        sem.release(); // no sleep here — skipped items must not be delayed
      }
    });

    await Promise.all(tasks.map((t) => t()));

    // ── Export GPTs ──────────────────────────────────────────────────
    let gptCount = 0;
    try {
      sendEvent("status", "Fetching your GPTs...");
      const gptsDir = join(rootDir, "gpts");
      mkdirSync(gptsDir, { recursive: true });
      let cursor = null;
      while (true) {
        const qs   = cursor ? `?cursor=${encodeURIComponent(cursor)}&limit=50` : "?limit=50";
        const data = await apiGet(`gizmos/discovery/mine${qs}`, token);
        const items = data.items || data.gizmos || [];
        if (!items.length) break;
        for (const item of items) {
          const gizmo  = item.gizmo || item;
          const id     = gizmo.id || "";
          const name   = gizmo.display?.name || gizmo.name || id;
          const safe   = sanitizeFilename(name);
          const fname  = `${safe}_${id.slice(0, 8)}`;
          const mdPath = join(gptsDir, `${fname}.md`);
          if (!existsSync(mdPath)) {
            const mdStr = gptToMarkdown(gizmo);
            writeFileSync(mdPath, mdStr, "utf8");
            zipFiles.push({ path: `gpts/${fname}.md`, data: mdStr });
          }
          gptCount++;
        }
        cursor = data.cursor || data.next_cursor || null;
        if (!cursor) break;
        await sleep(DELAY);
      }
      sendEvent("status", `Exported ${gptCount} GPT(s).`);
      log(`GPT export done — ${gptCount} GPT(s)`);
    } catch (e) {
      log(`GPT export skipped: ${e.message}`);
    }

    // ── Build ZIP ────────────────────────────────────────────────────
    sendEvent("status", "Creating ZIP archive...");
    const zipBuf = buildZip(zipFiles);
    writeFileSync(zipPath, zipBuf);

    const projectCount = Object.keys(convosByGroup).filter((k) => k !== "__root__").length;
    sendEvent("done", JSON.stringify({
      total,
      succeeded: newCount,
      skipped,
      failed: failed.length,
      failedTitles: failed,
      output: rootDir,
      zip: zipPath,
      totalFiles,
      failedFiles,
      failedFileDetails,
      projects: projectCount,
      gpts: gptCount,
    }));

  } catch (e) {
    sendEvent("error_msg", `Export failed: ${e.message}`);
  }
}

// ── HTTP server ──────────────────────────────────────────────────────

const exportJobs = new Map(); // exportId → { events: [], done: false }

const server = createServer((req, res) => {
  // Serve UI
  if (req.method === "GET" && (req.url === "/" || req.url === "")) {
    const htmlPath = join(__dirname, "public", "index.html");
    try {
      const html = readFileSync(htmlPath, "utf8");
      res.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      res.end(html);
    } catch {
      res.writeHead(500);
      res.end("Could not read public/index.html");
    }

  // Start export
  } else if (req.method === "POST" && req.url === "/start-export") {
    let body = "";
    req.on("data", (c) => (body += c));
    req.on("end", () => {
      let token, outputDir, concurrency;
      try {
        const parsed = JSON.parse(body);
        token       = parsed.token;
        outputDir   = parsed.outputDir   || join(homedir(), "Desktop", "chatgpt-export");
        concurrency = parsed.concurrency || 5;
      } catch {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "Invalid JSON body" }));
        return;
      }
      if (!token) {
        res.writeHead(400);
        res.end(JSON.stringify({ error: "No token provided" }));
        return;
      }

      const exportId = randomUUID().slice(0, 8);
      exportJobs.set(exportId, { events: [], done: false });

      const sendEvent = (type, data) => {
        const job = exportJobs.get(exportId);
        if (job) job.events.push({ type, data: String(data).replace(/\n/g, "\\n") });
        // Terminal output
        if (type === "progress") {
          const d = JSON.parse(data);
          const pct = String(Math.round((d.current / d.total) * 100)).padStart(3);
          process.stdout.write(`\r[${pct}%] ${d.current}/${d.total}  ${(d.title || "").slice(0, 50).padEnd(50)}`);
        } else if (type === "status") {
          process.stdout.write(`\r\x1b[2K${data}\n`);
        } else if (type === "done") {
          const d = JSON.parse(data);
          process.stdout.write(`\r\x1b[2K`);
          console.log(`\n✓ Done — ${d.succeeded} new · ${d.skipped} resumed · ${d.failed} failed`);
          console.log(`  Saved to: ${d.output}\n`);
        } else if (type === "error_msg") {
          console.error(`\n✗ ${data}\n`);
        }
      };

      runExport(token, outputDir, concurrency, sendEvent).finally(() => {
        const job = exportJobs.get(exportId);
        if (job) job.done = true;
      });

      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ exportId }));
    });

  // SSE progress stream
  } else if (req.method === "GET" && req.url.startsWith("/progress/")) {
    const exportId = req.url.split("/progress/")[1];
    const job      = exportJobs.get(exportId);
    if (!job) { res.writeHead(404); res.end(); return; }

    res.writeHead(200, {
      "Content-Type":  "text/event-stream",
      "Cache-Control": "no-cache",
      "Connection":    "keep-alive",
    });

    let sent = 0;
    const interval = setInterval(() => {
      while (sent < job.events.length) {
        const evt = job.events[sent++];
        res.write(`event: ${evt.type}\ndata: ${evt.data}\n\n`);
      }
      if (job.done && sent >= job.events.length) {
        clearInterval(interval);
        exportJobs.delete(exportId);
        res.end();
      }
    }, 200);

    req.on("close", () => clearInterval(interval));

  // Folder picker (macOS only)
  } else if (req.method === "GET" && req.url === "/pick-folder") {
    try {
      const result = execSync(
        `osascript -e 'POSIX path of (choose folder with prompt "Kies export map:")'`,
        { encoding: "utf8" }
      ).trim();
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ path: result }));
    } catch {
      // User cancelled the dialog
      res.writeHead(200, { "Content-Type": "application/json" });
      res.end(JSON.stringify({ path: null }));
    }

  } else {
    res.writeHead(404);
    res.end();
  }
});

// ── Start ────────────────────────────────────────────────────────────

const url = `http://${HOST}:${PORT}`;
server.listen(PORT, HOST, () => {
  console.log(`\nChatGPT Exporter v2 running at ${url}`);
  console.log("Press Ctrl+C to stop.\n");
  try {
    if      (process.platform === "darwin") execSync(`open "${url}"`);
    else if (process.platform === "linux")  execSync(`xdg-open "${url}"`);
    else if (process.platform === "win32")  execSync(`start "${url}"`);
  } catch {}
});
