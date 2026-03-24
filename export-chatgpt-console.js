// ── ChatGPT Conversation Exporter (v2) ───────────────────────────────
// Paste this into your browser console while on chatgpt.com
// Requires Chrome or Edge (File System Access API).
//
// Improvements over v1:
//   • Writes files directly to disk — no ZIP in memory (supports >2GB)
//   • Resume: skips conversations whose .json file already exists
//   • Parallel: downloads 5 conversations concurrently
//   • Projects: exports ChatGPT project conversations to projects/{name}/
//   • Faster: 100ms delay (was 500ms)
// ─────────────────────────────────────────────────────────────────────

(async () => {
  const API       = "/backend-api";
  const PAGE_SIZE = 100;
  const DELAY     = 100;   // ms between requests per slot
  const CONCURRENCY = 5;   // parallel conversation downloads
  const DEVICE_ID = crypto.randomUUID();

  const HEADERS = {
    "Content-Type": "application/json",
    Accept:         "application/json",
    "Oai-Device-Id":  DEVICE_ID,
    "Oai-Language":   "en-US",
  };

  const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

  // ── Sanity check ─────────────────────────────────────────────────────

  if (!window.showDirectoryPicker) {
    alert("This script requires Chrome or Edge (File System Access API is not available in this browser).");
    return;
  }

  // ── UI overlay ────────────────────────────────────────────────────────

  const overlay = document.createElement("div");
  overlay.id = "cge-overlay";
  overlay.innerHTML = `
    <div style="position:fixed;inset:0;background:rgba(0,0,0,0.75);z-index:99999;
      display:flex;align-items:center;justify-content:center;
      font-family:-apple-system,BlinkMacSystemFont,sans-serif">
      <div style="background:#1e293b;border-radius:16px;padding:40px;max-width:520px;
        width:90%;color:#e2e8f0;box-shadow:0 25px 50px rgba(0,0,0,0.5)">
        <h2 style="margin:0 0 6px;font-size:20px;color:#f8fafc">ChatGPT Exporter v2</h2>
        <p id="cge-status" style="color:#94a3b8;font-size:14px;margin:0 0 18px">Starting…</p>
        <div style="width:100%;height:8px;background:#334155;border-radius:4px;overflow:hidden;margin-bottom:8px">
          <div id="cge-bar" style="height:100%;background:#3b82f6;border-radius:4px;transition:width 0.3s;width:0%"></div>
        </div>
        <p id="cge-detail"  style="color:#64748b;font-size:13px;margin:0 0 6px;
          white-space:nowrap;overflow:hidden;text-overflow:ellipsis"></p>
        <p id="cge-stats"   style="color:#475569;font-size:12px;margin:0"></p>
      </div>
    </div>`;
  document.body.appendChild(overlay);

  const ui = {
    status: overlay.querySelector("#cge-status"),
    bar:    overlay.querySelector("#cge-bar"),
    detail: overlay.querySelector("#cge-detail"),
    stats:  overlay.querySelector("#cge-stats"),
    set(status, pct, detail, stats) {
      if (status != null) this.status.textContent = status;
      if (pct    != null) this.bar.style.width = pct + "%";
      if (detail != null) this.detail.textContent = detail;
      if (stats  != null) this.stats.textContent  = stats;
    },
    done(msg) {
      this.status.textContent = msg;
      this.bar.style.width    = "100%";
      this.bar.style.background = "#22c55e";
      this.detail.textContent = "Click anywhere to close.";
      overlay.querySelector("div").style.cursor = "pointer";
      overlay.addEventListener("click", () => overlay.remove());
    },
    error(msg) {
      this.status.textContent   = "Error: " + msg;
      this.bar.style.background = "#ef4444";
      this.detail.textContent   = "Click anywhere to close.";
      overlay.addEventListener("click", () => overlay.remove());
    },
  };

  // ── Get session token ─────────────────────────────────────────────────

  ui.set("Getting session token…");
  let token;
  try {
    const session = await fetch("/api/auth/session").then((r) => r.json());
    token = session.accessToken;
    if (!token) throw new Error("No accessToken in session");
  } catch (e) {
    ui.error("Failed to get session token. Are you logged in?");
    return;
  }

  // ── API helpers ───────────────────────────────────────────────────────

  async function apiGet(path) {
    const resp = await fetch(`${API}/${path}`, {
      headers: { ...HEADERS, Authorization: `Bearer ${token}` },
    });
    if (!resp.ok) throw new Error(`HTTP ${resp.status} on ${path}`);
    return resp.json();
  }

  async function apiFetchBinary(url) {
    const resp = await fetch(url);
    if (!resp.ok) throw new Error(`HTTP ${resp.status}`);
    const data = new Uint8Array(await resp.arrayBuffer());
    const contentType = resp.headers.get("content-type") || "";
    return { data, contentType };
  }

  const MIME_TO_EXT = {
    "image/png": ".png", "image/jpeg": ".jpg", "image/gif": ".gif",
    "image/webp": ".webp", "image/svg+xml": ".svg", "application/pdf": ".pdf",
    "text/plain": ".txt", "text/html": ".html", "text/csv": ".csv",
    "application/json": ".json", "application/zip": ".zip",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document": ".docx",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": ".xlsx",
  };

  // ── Semaphore for concurrency control ─────────────────────────────────

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

  // ── File System helpers ───────────────────────────────────────────────

  async function getSubDir(root, ...parts) {
    let dir = root;
    for (const p of parts) {
      dir = await dir.getDirectoryHandle(p, { create: true });
    }
    return dir;
  }

  // Returns true if the file at path (array of name segments) exists under root.
  async function fileExists(root, ...parts) {
    try {
      let dir = root;
      for (const p of parts.slice(0, -1)) {
        dir = await dir.getDirectoryHandle(p);
      }
      await dir.getFileHandle(parts.at(-1));
      return true;
    } catch {
      return false;
    }
  }

  async function writeFsFile(dir, filename, data) {
    const fh = await dir.getFileHandle(filename, { create: true });
    const wr = await fh.createWritable();
    await wr.write(data);
    await wr.close();
  }

  // ── Text utilities (from original) ───────────────────────────────────

  function sanitize(name) {
    return name.replace(/[<>:"/\\|?*]/g, "_").replace(/^[. ]+|[. ]+$/g, "").slice(0, 80) || "untitled";
  }

  function stripCitations(str) {
    return str.replace(/\u3010[^\u3011]*\u3011/g, "");
  }

  function escapeHtml(str) {
    return str
      .replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
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

  // ── File reference extraction & download (from original) ─────────────

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
            const m = part.asset_pointer.match(/^(?:file-service|sediment):\/\/(.+)$/);
            if (m && !seen.has(m[1])) {
              seen.add(m[1]);
              refs.push({ fileId: m[1], filename: part.metadata?.dalle?.prompt ? "dalle_image.png" : "image.png", type: "image" });
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

  async function downloadFile(fileId, fallbackName) {
    const meta = await apiGet(`files/download/${fileId}`);
    if (!meta.download_url) throw new Error("No download_url");
    const { data, contentType } = await apiFetchBinary(meta.download_url);
    let filename = meta.file_name || fallbackName || fileId;
    if (!filename.includes(".") && contentType) {
      const mime = contentType.split(";")[0].trim();
      const ext  = MIME_TO_EXT[mime];
      if (ext) filename += ext;
    }
    return { filename, data };
  }

  // ── Fetch conversation list ───────────────────────────────────────────

  ui.set("Fetching conversation list…");
  let allConversations = []; // { id, title, projectName? }

  // Regular (non-project) conversations
  try {
    let offset = 0;
    while (true) {
      const data  = await apiGet(`conversations?offset=${offset}&limit=${PAGE_SIZE}`);
      const items = data.items || [];
      if (!items.length) break;
      allConversations.push(...items.map((c) => ({ ...c, projectName: null })));
      const total = data.total || allConversations.length;
      ui.set(`Fetching conversations… ${allConversations.length}/${total}`);
      offset += PAGE_SIZE;
      if (offset >= total) break;
      await sleep(DELAY);
    }
  } catch (e) {
    ui.error(`Failed to fetch conversations: ${e.message}`);
    return;
  }

  // Projects + their conversations
  const projectMap = {}; // projectId → projectName
  try {
    ui.set("Fetching projects…");
    let projOffset = 0;
    const allProjects = [];
    while (true) {
      const data = await apiGet(`projects?offset=${projOffset}&limit=50`);
      // API may return { projects: [...] } or an array directly
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
      projectMap[proj.id] = projName;
      ui.set(`Fetching project: ${projName}…`);
      let offset = 0;
      while (true) {
        const data  = await apiGet(`projects/${proj.id}/conversations?offset=${offset}&limit=${PAGE_SIZE}`);
        const items = data.items || [];
        if (!items.length) break;
        for (const c of items) {
          if (!seenIds.has(c.id)) {
            seenIds.add(c.id);
            allConversations.push({ ...c, projectName: projName });
          } else {
            // Conversation also exists in main list — tag it with the project name
            const existing = allConversations.find((x) => x.id === c.id);
            if (existing) existing.projectName = projName;
          }
        }
        const total = data.total || items.length;
        offset += PAGE_SIZE;
        if (offset >= total) break;
        await sleep(DELAY);
      }
    }
  } catch (e) {
    // Projects endpoint may not exist for all account types — continue without it
    console.warn("[cge] Projects fetch failed:", e.message);
    ui.set(null, null, `Note: projects skipped (${e.message})`);
    await sleep(1500);
  }

  if (!allConversations.length) {
    ui.done("No conversations found.");
    return;
  }

  // ── Pick output directory ─────────────────────────────────────────────

  ui.set(`Found ${allConversations.length} conversations. Choose output folder…`);
  let rootDir;
  try {
    rootDir = await window.showDirectoryPicker({
      id:    "chatgpt-export",
      mode:  "readwrite",
      startIn: "downloads",
    });
  } catch (e) {
    ui.error("No folder selected — export cancelled.");
    return;
  }

  // ── Pass 1: Download conversations + files ────────────────────────────

  const total   = allConversations.length;
  let newCount  = 0;
  let skipped   = 0;
  let failed    = 0;
  let startTime = Date.now();

  // downloaded entries for HTML generation (only newly processed ones)
  // { fname, title, convo, fileMap, projectName }
  const downloaded = [];

  const sem = new Semaphore(CONCURRENCY);

  const tasks = allConversations.map((item, i) => async () => {
    await sem.acquire();
    try {
      const { id: cid, title: rawTitle, projectName } = item;
      const title  = rawTitle || "Untitled";
      const fname  = `${sanitize(title)}_${cid.slice(0, 8)}`;

      // Determine base directory segments for this conversation
      const baseDirParts = projectName
        ? ["projects", sanitize(projectName)]
        : [];

      // Resume check — if json file already exists, skip
      const alreadyDone = await fileExists(rootDir, ...baseDirParts, "json", `${fname}.json`);
      if (alreadyDone) {
        skipped++;
        const done = newCount + skipped + failed;
        const pct  = Math.round((done / total) * 100);
        ui.set(null, pct, `[skip] ${title}`, `${newCount} new · ${skipped} skipped · ${failed} failed`);
        return;
      }

      // Download conversation JSON
      const convo   = await apiGet(`conversation/${cid}`);
      const jsonStr = JSON.stringify(convo, null, 2);

      // Download attached files
      const fileRefs  = extractFileReferences(convo);
      const fileMap   = {};
      const usedNames = new Set();

      for (const ref of fileRefs) {
        try {
          const { filename: dlName, data } = await downloadFile(ref.fileId, ref.filename);
          const actualName = deduplicateFilename(dlName || ref.filename, usedNames);
          const filesDir   = await getSubDir(rootDir, ...baseDirParts, "files", fname);
          await writeFsFile(filesDir, actualName, data);
          // fileMap paths are relative from html/ to files/
          fileMap[ref.fileId] = `../files/${fname}/${actualName}`;
          await sleep(DELAY);
        } catch {
          // File download failed — continue
        }
      }

      // Write JSON + Markdown
      const jsonDir = await getSubDir(rootDir, ...baseDirParts, "json");
      await writeFsFile(jsonDir, `${fname}.json`, jsonStr);

      const mdDir = await getSubDir(rootDir, ...baseDirParts, "markdown");
      await writeFsFile(mdDir, `${fname}.md`, toMarkdown(convo, fileMap));

      downloaded.push({ fname, title, convo, fileMap, projectName });
      newCount++;

      const done    = newCount + skipped + failed;
      const pct     = Math.round((done / total) * 100);
      const elapsed = (Date.now() - startTime) / 1000 / 60;
      const rate    = elapsed > 0 ? Math.round(done / elapsed) : "…";
      ui.set(
        `Downloading ${done}/${total} (${pct}%)`,
        pct,
        title,
        `${newCount} new · ${skipped} skipped · ${failed} failed · ${rate}/min`
      );
    } catch {
      failed++;
    } finally {
      await sleep(DELAY);
      sem.release();
    }
  });

  await Promise.all(tasks.map((t) => t()));

  // ── Pass 2: Generate HTML ─────────────────────────────────────────────

  ui.set("Generating HTML pages…", 100);

  // Group downloaded conversations by project (null = regular)
  const groups = {};
  for (const d of downloaded) {
    const key = d.projectName || "__root__";
    if (!groups[key]) groups[key] = [];
    groups[key].push(d);
  }

  for (const [key, items] of Object.entries(groups)) {
    const allConvos    = items.map((d) => ({ fname: d.fname, title: d.title }));
    const baseDirParts = key === "__root__" ? [] : ["projects", sanitize(key)];

    for (const d of items) {
      try {
        const htmlStr = toHtml(d.convo, d.fileMap, allConvos, d.fname);
        const htmlDir = await getSubDir(rootDir, ...baseDirParts, "html");
        await writeFsFile(htmlDir, `${d.fname}.html`, htmlStr);
      } catch {
        // HTML generation failure is non-fatal
      }
    }
  }

  // ── Done ─────────────────────────────────────────────────────────────

  const totalFiles  = downloaded.reduce((s, d) => s + Object.keys(d.fileMap).length, 0);
  const projectCount = Object.keys(groups).filter((k) => k !== "__root__").length;
  let doneMsg = `Done! ${newCount} new · ${skipped} resumed · ${failed} failed`;
  if (totalFiles)    doneMsg += ` · ${totalFiles} files`;
  if (projectCount)  doneMsg += ` · ${projectCount} project(s)`;
  ui.done(doneMsg);

  // ── Markdown converter ────────────────────────────────────────────────

  function toMarkdown(convo, fileMap = {}) {
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

          if (role === "system" || role === "tool") {
            queue.push(...(node.children || []));
            continue;
          }
          if (role === "assistant" && contentType !== "text") {
            queue.push(...(node.children || []));
            continue;
          }

          const textParts = [];
          for (const part of msg.content.parts) {
            if (typeof part === "string") {
              textParts.push(part);
            } else if (part?.content_type === "image_asset_pointer" && part.asset_pointer) {
              const m = part.asset_pointer.match(/^(?:file-service|sediment):\/\/(.+)$/);
              textParts.push(m && fileMap[m[1]] ? `![image](${fileMap[m[1]]})` : "[image]");
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
          if (text) {
            lines.push(`## ${role.charAt(0).toUpperCase() + role.slice(1)}\n\n${text}\n`);
          }
        }
        queue.push(...(node.children || []));
      }
    }

    return lines.join("\n");
  }

  // ── HTML converter ────────────────────────────────────────────────────

  function toHtml(convo, fileMap = {}, allConversations = [], currentFname = "") {
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
              const m = part.asset_pointer.match(/^(?:file-service|sediment):\/\/(.+)$/);
              if (m && fileMap[m[1]]) imageParts.push(fileMap[m[1]]);
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

    const LOGO = '<svg viewBox="0 0 41 41" fill="none" xmlns="http://www.w3.org/2000/svg" width="24" height="24"><path d="M37.532 16.87a9.963 9.963 0 0 0-.856-8.184 10.078 10.078 0 0 0-10.855-4.835A9.964 9.964 0 0 0 18.306.5a10.079 10.079 0 0 0-9.614 6.977 9.967 9.967 0 0 0-6.664 4.834 10.08 10.08 0 0 0 1.24 11.817 9.965 9.965 0 0 0 .856 8.185 10.079 10.079 0 0 0 10.855 4.835 9.965 9.965 0 0 0 7.516 3.35 10.078 10.078 0 0 0 9.617-6.981 9.967 9.967 0 0 0 6.663-4.834 10.079 10.079 0 0 0-1.243-11.813ZM22.498 37.886a7.474 7.474 0 0 1-4.799-1.735c.061-.033.168-.091.237-.134l7.964-4.6a1.294 1.294 0 0 0 .655-1.134V19.054l3.366 1.944a.12.12 0 0 1 .066.092v9.299a7.505 7.505 0 0 1-7.49 7.496ZM6.392 31.006a7.471 7.471 0 0 1-.894-5.023c.06.036.162.099.237.141l7.964 4.6a1.297 1.297 0 0 0 1.308 0l9.724-5.614v3.888a.12.12 0 0 1-.048.103l-8.051 4.649a7.504 7.504 0 0 1-10.24-2.744ZM4.297 13.62A7.469 7.469 0 0 1 8.2 10.333c0 .068-.004.19-.004.274v9.201a1.294 1.294 0 0 0 .654 1.132l9.723 5.614-3.366 1.944a.12.12 0 0 1-.114.012L7.044 23.86a7.504 7.504 0 0 1-2.747-10.24Zm27.658 6.437-9.724-5.615 3.367-1.943a.121.121 0 0 1 .114-.012l8.048 4.648a7.498 7.498 0 0 1-1.158 13.528V21.36a1.293 1.293 0 0 0-.647-1.132v-.17Zm3.35-5.043c-.059-.037-.162-.099-.236-.141l-7.965-4.6a1.298 1.298 0 0 0-1.308 0l-9.723 5.614v-3.888a.12.12 0 0 1 .048-.103l8.05-4.645a7.497 7.497 0 0 1 11.135 7.763Zm-21.063 6.929-3.367-1.944a.12.12 0 0 1-.065-.092v-9.299a7.497 7.497 0 0 1 12.293-5.756 6.94 6.94 0 0 0-.236.134l-7.965 4.6a1.294 1.294 0 0 0-.654 1.132l-.006 11.225Zm1.829-3.943 4.33-2.501 4.332 2.5v5l-4.331 2.5-4.331-2.5V18Z" fill="currentColor"/></svg>';

    const INTERNAL_LABELS = {
      multimodal_text: "File context", code: "Code", execution_output: "Output",
      computer_output: "Output", tether_browsing_display: "Web browsing",
      system_error: "Error", text: "Tool output",
    };

    const messagesHtml = messages.map((m) => {
      if (m.isInternal) {
        const label = INTERNAL_LABELS[m.contentType] || "Internal context";
        const b64   = btoa(unescape(encodeURIComponent(m.text)));
        return `<details class="thinking"><summary>${label}</summary><div class="thinking-content md-content" dir="auto" data-md="${b64}"></div></details>`;
      }

      const roleClass = m.role === "user" ? "user" : "assistant";
      let content = "";

      if (m.role === "user") {
        content = `<div class="bubble" dir="auto">${escapeHtml(m.text).replace(/\n/g, "<br>")}</div>`;
      } else {
        const b64 = btoa(unescape(encodeURIComponent(m.text)));
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
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${title}</title>
<link rel="stylesheet" href="https://cdn.jsdelivr.net/gh/highlightjs/cdn-release/build/styles/github-dark.min.css">
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, "Helvetica Neue", sans-serif;
    background: #ffffff; color: #0d0d0d; line-height: 1.65; font-size: 16px; display: flex; height: 100vh; }
  .sidebar { width: 260px; min-width: 260px; height: 100vh; background: #f9f9f9; border-right: 1px solid #e5e5e5;
    overflow-y: auto; padding: 16px 0; flex-shrink: 0; position: sticky; top: 0; }
  .sidebar-header { padding: 8px 16px 16px; font-size: 14px; font-weight: 600; color: #6b6b6b;
    border-bottom: 1px solid #e5e5e5; margin-bottom: 8px; }
  .sidebar-item { display: block; padding: 8px 16px; font-size: 13px; color: #0d0d0d; text-decoration: none;
    white-space: nowrap; overflow: hidden; text-overflow: ellipsis; border-radius: 8px; margin: 2px 8px; }
  .sidebar-item:hover { background: #ececec; }
  .sidebar-item.active { background: #e5e5e5; font-weight: 600; }
  .sidebar-toggle { display: none; position: fixed; top: 12px; left: 12px; z-index: 100;
    background: #f4f4f4; border: 1px solid #e5e5e5; border-radius: 8px; width: 36px; height: 36px;
    cursor: pointer; align-items: center; justify-content: center; font-size: 20px; }
  @media (max-width: 768px) {
    .sidebar { position: fixed; left: -280px; z-index: 99; transition: left 0.2s; box-shadow: 2px 0 8px rgba(0,0,0,0.1); }
    .sidebar.open { left: 0; }
    .sidebar-toggle { display: flex; }
    .main { margin-left: 0 !important; }
  }
  .main { flex: 1; overflow-y: auto; }
  .header { max-width: 768px; margin: 0 auto; padding: 32px 24px 16px; border-bottom: 1px solid #e5e5e5; }
  .header h1 { font-size: 22px; font-weight: 600; }
  .header .date { font-size: 13px; color: #6b6b6b; margin-top: 4px; }
  .chat { max-width: 768px; margin: 0 auto; padding: 24px; }
  .message { margin-bottom: 24px; }
  .message.user { display: flex; flex-wrap: wrap; justify-content: flex-end; gap: 8px; }
  .message.user .bubble { background: #f4f4f4; border-radius: 18px; padding: 10px 16px;
    max-width: 85%; white-space: pre-wrap; word-break: break-word; }
  .message.user .images { width: 100%; display: flex; justify-content: flex-end; }
  .message.assistant { display: flex; gap: 12px; align-items: flex-start; }
  .message.assistant .avatar { width: 28px; height: 28px; border-radius: 50%; background: #00a67e; color: #fff;
    display: flex; align-items: center; justify-content: center; flex-shrink: 0; margin-top: 2px; }
  .message.assistant .content { flex: 1; min-width: 0; }
  .message.assistant .content h1, .message.assistant .content h2,
  .message.assistant .content h3 { margin: 16px 0 8px; font-weight: 600; }
  .message.assistant .content h1 { font-size: 20px; }
  .message.assistant .content h2 { font-size: 18px; }
  .message.assistant .content h3 { font-size: 16px; }
  .message.assistant .content p { margin: 8px 0; }
  .message.assistant .content ul, .message.assistant .content ol { margin: 8px 0; padding-left: 24px; }
  .message.assistant .content li { margin: 4px 0; }
  .message.assistant .content a { color: #1a7f64; }
  .message.assistant .content code { background: #f0f0f0; border-radius: 4px; padding: 2px 5px;
    font-family: "SFMono-Regular", Consolas, "Liberation Mono", Menlo, monospace; font-size: 14px; }
  .message.assistant .content pre { margin: 12px 0; border-radius: 8px; overflow: hidden; }
  .message.assistant .content pre code { display: block; background: #0d0d0d; color: #f8f8f2;
    padding: 16px; overflow-x: auto; border-radius: 0; font-size: 13px; line-height: 1.5; }
  .code-block { position: relative; }
  .code-block .copy-btn { position: absolute; top: 8px; right: 8px; background: #333; border: none;
    color: #999; cursor: pointer; font-size: 12px; padding: 4px 10px; border-radius: 4px;
    opacity: 0; transition: opacity 0.2s; }
  .code-block:hover .copy-btn { opacity: 1; }
  .code-block .copy-btn:hover { color: #fff; background: #555; }
  .images img { max-width: 100%; border-radius: 8px; margin: 4px 0; display: block; cursor: pointer; }
  .images img:hover { opacity: 0.9; }
  .message.user .images img { max-width: 300px; }
  .attachments { margin-top: 8px; display: flex; flex-wrap: wrap; gap: 8px; }
  .attachment { display: inline-flex; align-items: center; gap: 8px; background: #f4f4f4;
    border: 1px solid #e5e5e5; border-radius: 8px; padding: 8px 12px; text-decoration: none;
    color: #0d0d0d; font-size: 14px; }
  .attachment:hover { background: #ececec; }
  .att-icon { font-size: 16px; }
  .att-name { overflow: hidden; text-overflow: ellipsis; white-space: nowrap; max-width: 200px; }
  .thinking { margin-bottom: 24px; border-left: 3px solid #d4d4d4; padding-left: 16px; font-size: 14px; }
  .thinking summary { color: #8e8e8e; font-style: italic; cursor: pointer; padding: 4px 0; user-select: none; }
  .thinking summary:hover { color: #555; }
  .thinking-content { color: #6b6b6b; padding: 8px 0; font-style: italic; }
  .thinking-content p, .thinking-content li { color: #6b6b6b; }
  .thinking-content pre code { opacity: 0.7; }
  .thinking-content h1, .thinking-content h2, .thinking-content h3 { color: #6b6b6b; font-size: 15px; }
</style>
</head>
<body>
<button class="sidebar-toggle" onclick="document.querySelector('.sidebar').classList.toggle('open')">&#9776;</button>
<nav class="sidebar">
  <div class="sidebar-header">Conversations</div>
  ${sidebarItems}
</nav>
<div class="main">
  <div class="header">
    <h1>${title}</h1>
    ${dateStr ? `<div class="date">${dateStr}</div>` : ""}
  </div>
  <div class="chat">
  ${messagesHtml}
  </div>
</div>
<script src="https://cdn.jsdelivr.net/npm/marked/marked.min.js"><\/script>
<script src="https://cdn.jsdelivr.net/gh/highlightjs/cdn-release/build/highlight.min.js"><\/script>
<script>
document.addEventListener('DOMContentLoaded', () => {
  marked.setOptions({
    highlight: (code, lang) => {
      if (lang && hljs.getLanguage(lang)) return hljs.highlight(code, { language: lang }).value;
      return hljs.highlightAuto(code).value;
    },
    breaks: true,
  });
  const renderer = new marked.Renderer();
  renderer.code = function({ text, lang }) {
    const highlighted = lang && hljs.getLanguage(lang)
      ? hljs.highlight(text, { language: lang }).value
      : hljs.highlightAuto(text).value;
    return '<div class="code-block"><button class="copy-btn" onclick="navigator.clipboard.writeText(this.nextElementSibling.querySelector(\\'code\\').textContent);this.textContent=\\'Copied!\\';setTimeout(()=>this.textContent=\\'Copy\\',1500)">Copy</button>'
      + '<pre><code class="hljs">' + highlighted + '</code></pre></div>';
  };
  marked.use({ renderer });
  document.querySelectorAll('.md-content').forEach(el => {
    const md = decodeURIComponent(escape(atob(el.dataset.md)));
    el.innerHTML = marked.parse(md);
  });
  const active = document.querySelector('.sidebar-item.active');
  if (active) active.scrollIntoView({ block: 'center', behavior: 'instant' });
});
<\/script>
</body>
</html>`;
  }
})();
