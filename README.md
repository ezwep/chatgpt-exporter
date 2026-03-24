# ChatGPT Exporter

A zero-dependency Node.js tool that exports **all** your ChatGPT data — conversations, projects (with uploaded files), and custom GPTs — to JSON, Markdown, and browseable HTML.

## Why this tool?

OpenAI's built-in data export is slow, gives you a single massive JSON blob, and **doesn't include** custom GPT configurations or project resource files. This tool fills that gap.

## Features

- **Conversations** — full message history with code blocks, images, and file attachments
- **Projects** — metadata, instructions, uploaded resources (PDFs, documents), and all project conversations organized in subfolders
- **Custom GPTs** — name, description, instructions, tools, and conversation starters
- **3 output formats** — JSON (raw API data), Markdown (readable), HTML (browseable with sidebar)
- **HTML viewer** — generated `index.html` with ChatGPT-style dark sidebar, project folders, and search
- **Resume support** — re-run at any time; already-exported conversations are skipped instantly
- **Parallel downloads** — configurable concurrency (1–10)
- **Date filter** — export only conversations within a specific date range
- **ZIP archive** — optionally bundle everything into a single `.zip`
- **Keep awake** — prevents sleep on macOS (`caffeinate`), Windows (`SetThreadExecutionState`), and Linux (`systemd-inhibit`)
- **Zero dependencies** — only built-in Node.js modules
- **Web UI** — clean local interface, no terminal required after startup

## Output structure

```
chatgpt-export/
├── index.html        ← HTML viewer with search and iframe navigation
├── json/             ← raw conversation JSON
├── markdown/         ← readable Markdown per conversation
├── html/             ← browseable HTML with conversation sidebar
├── files/            ← images and attachments per conversation
├── projects/
│   └── Project Name/
│       ├── project-info.md
│       ├── resources/    ← uploaded project files (PDFs, images, etc.)
│       ├── json/
│       ├── markdown/
│       └── html/
├── gpts/             ← your custom GPTs as Markdown
└── logs/             ← export run logs
```

## Quick start

```bash
git clone https://github.com/ezwep/chatgpt-exporter.git
cd chatgpt-exporter
npm install
npm start
```

The browser opens automatically at `http://127.0.0.1:8523`.

### Get your session token

1. Open [chatgpt.com/api/auth/session](https://chatgpt.com/api/auth/session) in a new tab (log in first if needed)
2. Select all (`Cmd+A` / `Ctrl+A`), copy, and paste into the text area
3. Click **Export conversations**

The export runs in the background. You can close the browser tab — progress is saved and resumable.

## Export options

Click **Export options** to expand:

| Option | Description |
|--------|-------------|
| Conversations / Projects / GPTs | Toggle which categories to export |
| JSON / Markdown / HTML | Toggle output formats |
| From / Until | Date range filter (leave empty = export all) |
| Keep computer awake | Prevent sleep during long exports |
| Create ZIP archive | Bundle output into a `.zip` file |

> **Note:** JSON is always written to disk (used for resume detection), but only included in the ZIP when the JSON format is enabled.

## Requirements

- Node.js 18 or later
- macOS, Linux, or Windows

## How it works

The tool uses ChatGPT's internal API endpoints (the same ones the web app uses) with your session token. It fetches conversation lists, individual conversations with full message trees, project metadata and resources, and custom GPT configurations. No third-party APIs or services are involved — everything stays on your local machine.

## Disclaimer

**Use at your own risk.** This tool is provided as-is, without warranty of any kind.

- **Unofficial API** — This tool uses ChatGPT's internal web API endpoints, which are undocumented and not part of the official OpenAI API. These endpoints may change, break, or be blocked at any time without notice.
- **Terms of Service** — Using internal API endpoints may violate [OpenAI's Terms of Use](https://openai.com/policies/terms-of-use). You are solely responsible for ensuring your use complies with applicable terms and laws.
- **Session token security** — Your session token grants full access to your ChatGPT account. This tool processes it locally and never transmits it to any third party, but you should **never share your token** with anyone. Treat it like a password.
- **Your own data** — This tool only accesses your own conversations, projects, and GPTs. It is intended for personal data portability (consistent with GDPR Article 20), not for accessing other users' data.
- **No affiliation** — This project is not affiliated with, endorsed by, or associated with OpenAI in any way.
- **Rate limiting** — Excessive use may trigger rate limits or temporary blocks on your ChatGPT account. The tool includes built-in delays to minimize this risk, but no guarantees are made.
- **Data accuracy** — Exported data is provided as-is from the API. The author is not responsible for missing, incomplete, or corrupted exports.

## Credits

Inspired by the original browser console export script by [**ocombe**](https://gist.github.com/ocombe/1d7604bd29a91ceb716304ef8b5aa4b5).

This project extends that approach into a full local server with resume support, project & GPT export, HTML browsing, search, and a web UI.

## Support

If you find this tool useful, consider buying me a coffee:

[![ko-fi](https://ko-fi.com/img/githubbutton_sm.svg)](https://ko-fi.com/ezwep)

## License

MIT
