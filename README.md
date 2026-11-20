# SharePoint Class Downloader

Download class recordings from a SharePoint / OneDrive shared folder using **your existing browser login**, with resumable downloads and verification.

## What it does

1. Takes a SharePoint folder URL (anonymous `/:f:/g/...` share link, `/shared?listurl=...&id=...` link, or a direct `onedrive.aspx?id=...` URL).
2. Reuses your browser session (cookies are extracted from your real Chromium profile, including OS-keyring decryption).
3. Enumerates the folder through SharePoint's own `RenderListDataAsStream` API — this yields filenames, sizes and per-file share tokens without scraping the UI.
4. Builds a `stream.aspx` player URL for each video and downloads it with `yt-dlp`, merging DASH video + audio into an `.mp4`.
5. Verifies every result with `ffprobe` (video + audio streams present) and skips files that are already verified.

## Why it exists

Microsoft SharePoint hides or denies the direct file-download endpoint for many shared folders:

- `download.aspx?share=...` returns HTTP 200 with an **HTML error page** (or 403) even for a logged-in user.
- The `source` format reported by `yt-dlp` is often that same error page saved as `.mp4`.
- Guest shares expose a "More Actions" menu with no **Download** entry, but the streaming player works.

The streaming manifests do work. They are AES-encrypted DASH fragments, which `yt-dlp` + `ffmpeg` can resolve when given the browser cookies. The downloader therefore deliberately uses the *player* endpoint, retries rate-limited fragments (`HTTP 429`), and verifies the bytes instead of trusting an HTTP status code.

It also handles the awkward bit nobody thinks about first: the browser session you can automate (a managed/CDP browser) is often **not** the browser where you are logged in. This tool rips cookies from the real Chromium profile via `yt-dlp.cookies.extract_cookies_from_browser('chromium')`, which handles Linux keyring decryption.

## Requirements

- Python 3.10+
- `yt-dlp` (`pip install -r requirements.txt`)
- `ffmpeg` / `ffprobe` on `PATH` (system package: `pacman -S ffmpeg`, `apt install ffmpeg`, `brew install ffmpeg`, …)
- A browser profile that is logged into the SharePoint tenant, or an explicit Netscape cookie jar.

No credentials are ever typed or stored by this tool. It only reads the session that already exists in your browser.

## How to use

### 1. List the videos in a folder first

```bash
python spclassdl.py 'https://ipbejapt-my.sharepoint.com/:f:/g/personal/teacher/OTHER_ID?e=TOKEN' --list-only
```

Or with an authenticated `/shared?...` style link:

```bash
python spclassdl.py 'https://…/shared?listurl=…&id=%2Fpersonal%2Fteacher%2FDocuments%2FRecordings' --list-only
```

If the URL is a `/people?emailId=...` link with no folder path, add `--folder`:

```bash
python spclassdl.py 'https://…/people?emailId=teacher%40example.edu' \
  --folder '/personal/teacher/Documents/Recordings' --list-only
```

### 2. Download everything

```bash
python spclassdl.py 'SHARE_URL' --out ~/Downloads/Recordings
```

Useful options:

| Option | Purpose |
|---|---|
| `--format SPEC[,SPEC…]` | yt-dlp format specs, tried in order. Default: `dash-v1080p+dash-audhigh,dash-vcopy+dash-audcopy,dash-v720p+dash-audhigh,bestvideo+bestaudio` |
| `--limit N` | Only process the first N videos (smoke test) |
| `--out DIR` | Output directory (default `./downloads`) |
| `--fragments N` | Concurrent fragments (default 3; lower it if the server throws 429s) |
| `--fragment-retries N`, `--retry-sleep SEC` | Rate-limit backoff knobs |
| `--attempts N` | Retry whole files N times |
| `--cookies FILE` | Use a Netscape cookie jar instead of browser extraction |
| `--browser NAME`, `--browser-profile NAME` | Which browser profile to rip cookies from |
| `--no-browser-cookies` | Use anonymous/no cookies |
| `--match REGEX` | Only process filenames matching this regex |
| `--all-files` | Include non-video files in the listing |
| `--verbose` | More output |

### 3. Resume / re-run

Re-running the same command skips files that already exist and pass `ffprobe`. Interrupted downloads leave `yt-dlp` fragment state, so the next attempt resumes rather than starting over.

### 4. Use an explicit cookie jar

If automatic extraction fails (locked keyring, unusual profile layout), export a jar yourself and pass it:

```bash
python - <<'PY'
from yt_dlp.cookies import extract_cookies_from_browser
extract_cookies_from_browser('chromium').save('cookies.txt')
PY
python spclassdl.py 'SHARE_URL' --cookies cookies.txt --list-only
```

Keep that file private — it contains session credentials. It is gitignored.

## Typical output

```
Exported 536 cookies from the chromium profile
Site: https://tenant-my.sharepoint.com/personal/teacher  Folder: /personal/teacher/Documents/Recordings
Found 4 video(s) (2.15 GiB source)
[1/4] recording-a.mp4 (658 MiB source)
  yt-dlp -f dash-vcopy+dash-audcopy -> recording-a.mp4 (share=***)
  verified: 88.8 min, 658.1 MiB
[2/4] recording-b.mp4 SKIP ... (already verified, 111.0 min)
Done. 4 verified, 0 failed.
```

Share tokens and `tempauth` values are redacted in logs.

## Troubleshooting

- **`Could not determine folder path`** — pass `--folder` explicitly.
- **List API returns `FileNotFoundException`** — the `--folder` path or owner in the URL is wrong; re-check `id=` in the browser address bar.
- **Everything fails with 429** — set `--fragments 1 --fragment-retries 60 --retry-sleep 30`. Slow, but it eventually gets through.
- **"yt-dlp finished but ffprobe rejected the output"** — the service served an error page as media; the next format spec is tried automatically.
- **`Session cookies are required for this URL`** from yt-dlp — the active jar has no cookie for the SharePoint host. Re-export cookies from the logged-in browser, or pass `--cookies cookies.txt` with a jar that contains `FedAuth`.
- **Cookies exported but still access denied** — the logged-in account does not have permission on that folder; there is no workaround in the tool.

## Legal / ethical use

Only use this on material you are authorised to download. Respect your institution's policies and the privacy of others. The tool is a convenience wrapper around your own browser session; it does not bypass authentication or access controls.

## License

Unlicense — see [UNLICENSE](UNLICENSE). Do what you want; no warranty.
