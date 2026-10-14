# Teams Excavator

A Chrome / Edge extension (Manifest V3) that exports Microsoft Teams chats and file attachments straight from the Teams web client.

- **Exports:** CSV, HTML, Markdown, and ZIP archives of attachments
- **Works on:** `teams.microsoft.com`, `teams.live.com`, `teams.cloud.microsoft`
- **Private:** everything runs locally in your browser. No server, no telemetry, no external dependencies.

## Features

- Auto-scrolls a chat and extracts every message: author, timestamp, text, reactions, quoted replies
- Channel threads: drills into replies
- File attachments downloaded into a single streamed ZIP (ZIP64, DEFLATE) that handles large archives without loading them fully in memory
- Inline images captured as embedded `data:` URLs (150 MB memory cap)
- Live progress with message and file counts, start/stop control, and total extraction time
- Batch mode for several chats
- Readable, dated export file names; unnamed group chats are titled with all participants
- Popup state survives closing and reopening; downloads keep running after the popup closes
- 20 locales, including right-to-left (Arabic, Urdu)

## Installation

There is no build step. Load the folder as an unpacked extension.

1. Clone or download this repository.
2. Open `chrome://extensions` (or `edge://extensions`).
3. Enable **Developer mode**.
4. Click **Load unpacked** and select the project folder.

## Usage

1. Open [Teams on the web](https://teams.microsoft.com) and open the chat or channel you want to export.
2. Click the extension icon.
3. Start the extraction and wait while the chat is scrolled and parsed. You can stop at any time.
4. Choose an export: **CSV**, **HTML**, **Markdown**, or **Download files (zip)**.

Keep the Teams tab open while extraction runs. Results live in that tab until you start a new extraction or reload it.

## Permissions

| Permission | Why |
|---|---|
| `activeTab`, `scripting` | Inject the extraction script into the Teams tab |
| `storage` | Remember popup state and settings |
| `downloads` | Save exports and the ZIP |
| Host access to Teams, SharePoint, OneDrive | Fetch attachments referenced by messages |

## Project structure

| Path | Role |
|---|---|
| `manifest.json` | MV3 manifest |
| `src/content/` | Extraction engine injected into Teams; holds the results |
| `service-worker.js`, `src/background/` | Exports (CSV/HTML/MD), streamed ZIP64 downloads |
| `src/popup/` | UI and language selection |
| `src/shared/` | Helpers shared by all contexts |
| `_locales/` | Translations |

Plain vanilla JavaScript: no npm, no bundler, no TypeScript. See `AGENTS.md` for the architecture.

## Limitations

- Extraction relies on the Teams web DOM, so Microsoft DOM changes in Teams Web can break it.
- Only what the web client loads can be exported; very old history requires scrolling to load.
- Attachments must be accessible to your signed-in account.

## Disclaimer

This project is not affiliated with or endorsed by Microsoft. Only export data you are authorized to access, and respect your organization's policies.
