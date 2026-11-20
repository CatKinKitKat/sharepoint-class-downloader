#!/usr/bin/env python3
"""Download class-recording videos from SharePoint / OneDrive shared folders.

Uses the logged-in browser session, the site's own list API for enumeration,
and yt-dlp + ffmpeg for the actual media. See README.md for the rationale.
"""

from __future__ import annotations

import argparse
import json
import os
import re
import subprocess
import sys
import tempfile
import time
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, quote, unquote, urlsplit
from urllib.request import HTTPCookieProcessor, Request, build_opener

VIDEO_EXTENSIONS = {".mp4", ".mkv", ".mov", ".m4v", ".webm", ".avi", ".ts"}
DEFAULT_FORMATS = (
    "dash-v1080p+dash-audhigh",
    "dash-vcopy+dash-audcopy",
    "dash-v720p+dash-audhigh",
    "bestvideo+bestaudio",
)


def log(message: str) -> None:
    print(message, flush=True)


def redact(url: str) -> str:
    """Keep share tokens and tempauth values out of logs."""
    return re.sub(r"(share|tempauth|access_token)=[^&\s]+", r"\1=***", url)


def require(command: str) -> None:
    if not shutil_which(command):
        raise SystemExit(f"Required command not found: {command}")


def shutil_which(command: str) -> str | None:
    from shutil import which

    return which(command)


def export_browser_cookies(browser: str, profile: str | None) -> tuple[str | None, str | None]:
    """Rip cookies from the user's real browser, including OS-keyring decryption."""
    try:
        from yt_dlp.cookies import extract_cookies_from_browser
    except ImportError as exc:
        log(f"WARNING: yt-dlp cookie module unavailable: {exc}")
        return None, None
    try:
        jar = extract_cookies_from_browser(browser, profile=profile)
        handle = tempfile.NamedTemporaryFile(
            prefix="spclassdl-cookies-", suffix=".txt", delete=False
        )
        handle.close()
        jar.save(handle.name)
        log(f"Exported {len(jar)} cookies from the {browser} profile")
        return handle.name, handle.name
    except Exception as exc:  # keyring locked, profile missing, etc.
        log(f"WARNING: could not export {browser} cookies: {exc}")
        return None, None


def has_session_cookie(jar_path: str | None, host: str) -> bool:
    """FedAuth / rtFA are the cookies SharePoint player requests really need."""
    session_names = {"FedAuth", "rtFA", "SignIn", "ESTSAUTH", "ESTSAUTHPERSISTENT"}
    if not jar_path:
        return False
    from http.cookiejar import MozillaCookieJar

    jar = MozillaCookieJar(jar_path)
    try:
        jar.load(ignore_discard=True, ignore_expires=True)
    except OSError:
        return False
    return any(
        host.endswith(cookie.domain.lstrip(".")) and cookie.name in session_names
        for cookie in jar
    )


def make_opener(jar_path: str | None):
    from http.cookiejar import MozillaCookieJar

    if not jar_path:
        return build_opener()
    jar = MozillaCookieJar(jar_path)
    jar.load(ignore_discard=True, ignore_expires=True)
    return build_opener(HTTPCookieProcessor(jar))


def _query(url: str) -> dict[str, list[str]]:
    return parse_qs(urlsplit(url).query, keep_blank_values=True)


def _first(params: dict[str, list[str]], key: str) -> str | None:
    values = params.get(key)
    return values[0] if values else None


def resolve_target(url: str, opener, folder_override: str | None) -> dict:
    """Turn any supported SharePoint URL into host/owner/doclib/folder."""
    parts = urlsplit(url)
    host = parts.netloc
    if not host:
        raise SystemExit(f"Invalid URL: {url}")

    folder = folder_override
    match = re.search(r"^/personal/([^/]+)/", parts.path)
    owner = match.group(1) if match else None

    if not folder:
        params = _query(url)
        folder = _first(params, "id")

    # Anonymous /:f:/g/... share links redirect to onedrive.aspx?id=<path>.
    if not folder and "/:f:" in parts.path:
        try:
            response = opener.open(Request(url, headers={"User-Agent": "Mozilla/5.0"}), timeout=60)
            final_url = response.geturl()
            folder = _first(_query(final_url), "id")
            response.close()
        except (HTTPError, URLError) as exc:
            raise SystemExit(f"Could not resolve share link: {exc}")

    if not folder:
        raise SystemExit(
            "Could not determine folder path. Supply --folder /personal/<owner>/Documents/<folder>."
        )
    folder = unquote(folder)
    if not folder.startswith("/personal/"):
        raise SystemExit(f"Unexpected folder path: {folder}")

    if not owner:
        owner = folder.split("/")[2]
    doclib = f"/personal/{owner}/Documents"
    site = f"https://{host}/personal/{owner}"
    return {"host": host, "owner": owner, "site": site, "doclib": doclib, "folder": folder}


def list_rows(target: dict, opener, recursive: bool = False) -> list[dict]:
    """Replay SharePoint's own RenderListDataAsStream endpoint."""
    api = f"{target['site']}/_api/web/GetListUsingPath(DecodedUrl=@a1)/RenderListDataAsStream"
    params = {
        "@a1": f"'{target['doclib']}'",
        "RootFolder": target["folder"],
        "TryNewExperienceSingle": "TRUE",
    }
    query = "&".join(f"{key}={quote(value, safe='')}" for key, value in params.items())
    body = json.dumps(
        {
            "parameters": {
                "__metadata": {"type": "SP.RenderListDataParameters"},
                "RenderOptions": 547,
                "ViewXml": "<View Scope='RecursiveAll' ><Query/></View>"
                if recursive
                else "<View><Query/></View>",
            }
        }
    ).encode("utf-8")
    request = Request(
        f"{api}?{query}",
        data=body,
        method="POST",
        headers={
            "Content-Type": "application/json;odata=verbose",
            "Accept": "application/json;odata=verbose",
            "User-Agent": "Mozilla/5.0",
        },
    )
    try:
        with opener.open(request, timeout=60) as response:
            payload = json.loads(response.read().decode("utf-8", "replace"))
    except (HTTPError, URLError, json.JSONDecodeError) as exc:
        log(f"ERROR: list API failed for {target['folder']}: {exc}")
        return []

    rows = []
    for row in payload.get("ListData", {}).get("Row", []):
        token = ""
        try:
            token = json.loads(row.get(".mainLinkInfo", "{}")).get("ShareTokenString", "")
        except (TypeError, ValueError):
            pass
        rows.append(
            {
                "name": row.get("FileLeafRef", ""),
                "size": int(row.get("File_x0020_Size") or 0),
                "path": f"{row.get('FileDirRef', '').rstrip('/')}/{row.get('FileLeafRef', '')}",
                "token": token,
                "is_dir": str(row.get("FSObjType", "0")) == "1",
            }
        )
    return rows


def stream_url(target: dict, row: dict) -> str:
    return (
        f"{target['site']}/_layouts/15/stream.aspx"
        f"?id={quote(row['path'], safe='')}&share={row['token']}"
    )


def probe(path: Path) -> tuple[bool, float]:
    if not path.exists() or path.stat().st_size < 10_000:
        return False, 0.0
    if not shutil_which("ffprobe"):
        return path.stat().st_size > 10_000_000, 0.0
    result = subprocess.run(
        [
            "ffprobe", "-v", "error",
            "-show_entries", "format=duration",
            "-show_entries", "stream=codec_type",
            "-of", "default=nw=1",
            str(path),
        ],
        capture_output=True,
        text=True,
    )
    text = result.stdout
    duration_match = re.search(r"duration=([0-9.]+)", text)
    duration = float(duration_match.group(1)) if duration_match else 0.0
    ok = result.returncode == 0 and "codec_type=video" in text and "codec_type=audio" in text
    return ok, duration


def download_one(url: str, target_path: Path, formats: list[str], args, jar_path: str | None) -> bool:
    for fmt in formats:
        output = target_path.parent / (target_path.stem + ".%(ext)s")
        command = [
            "yt-dlp", "--no-warnings", "--newline",
            "-f", fmt, "--merge-output-format", "mp4",
            "--concurrent-fragments", str(args.fragments),
            "--fragment-retries", str(args.fragment_retries),
            "--retry-sleep", str(args.retry_sleep),
        ]
        if jar_path:
            command += ["--cookies", jar_path]
        command += ["-o", str(output), url]
        log(f"  yt-dlp -f {fmt} -> {target_path.name} ({redact(url)})")
        result = subprocess.run(command)
        if result.returncode == 0:
            ok, duration = probe(target_path)
            if ok:
                log(f"  verified: {duration / 60:.1f} min, {target_path.stat().st_size / 2**20:.1f} MiB")
                return True
            log("  WARNING: yt-dlp finished but ffprobe rejected the output; trying next format")
    return False


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("url", nargs="?", help="SharePoint share/folder URL")
    parser.add_argument("--url", dest="url_opt", help="SharePoint URL (alternative to positional)")
    parser.add_argument("--out", default="downloads", help="output directory (default: ./downloads)")
    parser.add_argument("--folder", help="override folder path, e.g. /personal/user/Documents/Folder")
    parser.add_argument("--list-only", action="store_true", help="enumerate and exit")
    parser.add_argument("--limit", type=int, help="process at most N videos (handy for a test run)")
    parser.add_argument("--format", default=",".join(DEFAULT_FORMATS),
                        help="comma-separated yt-dlp format specs, tried in order")
    parser.add_argument("--fragments", type=int, default=3, help="concurrent fragments (default: 3)")
    parser.add_argument("--fragment-retries", type=int, default=25)
    parser.add_argument("--retry-sleep", type=int, default=10)
    parser.add_argument("--attempts", type=int, default=3, help="retries per file")
    parser.add_argument("--cookies", help="Netscape cookie jar to use instead of browser extraction")
    parser.add_argument("--browser", default="chromium", help="browser for cookie extraction (default: chromium)")
    parser.add_argument("--browser-profile", help="browser profile name/path (passed to yt-dlp)")
    parser.add_argument("--no-browser-cookies", action="store_true", help="do not extract browser cookies")
    parser.add_argument("--all-files", action="store_true", help="include non-video files")
    parser.add_argument("--match", help="regex filter on filenames (case-insensitive)")
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args()

    url = args.url_opt or args.url
    if not url:
        parser.error("a SharePoint URL is required")
    require("yt-dlp")

    jar_path = args.cookies
    temporary_jar = None
    if not jar_path and not args.no_browser_cookies:
        jar_path, temporary_jar = export_browser_cookies(args.browser, args.browser_profile)
    opener = make_opener(jar_path)

    target = resolve_target(url, opener, args.folder)
    log(f"Site: {target['site']}  Folder: {target['folder']}")
    if not has_session_cookie(jar_path, target["host"]):
        log(
            "WARNING: no SharePoint session cookie (FedAuth/rtFA) in the active jar. "
            "Listings may work anonymously, but yt-dlp needs a session; "
            "pass --cookies <jar> if downloads fail."
        )
    rows = list_rows(target, opener)
    videos = [row for row in rows if not row["is_dir"] and
              (args.all_files or Path(row["name"]).suffix.lower() in VIDEO_EXTENSIONS)]
    if args.match:
        videos = [row for row in videos if re.search(args.match, row["name"], re.IGNORECASE)]
    log(f"Found {len(videos)} video(s) ({sum(row['size'] for row in videos) / 2**30:.2f} GiB source)")
    if args.list_only:
        for row in videos:
            print(f"{row['name']}\t{row['size'] / 2**20:.1f} MiB")
        return 0
    if not videos:
        return 1

    out_dir = Path(args.out).expanduser()
    out_dir.mkdir(parents=True, exist_ok=True)
    formats = [item.strip() for item in args.format.split(",") if item.strip()]
    failures: list[str] = []

    for index, row in enumerate(videos, 1):
        if args.limit and index > args.limit:
            break
        destination = out_dir / (Path(row["name"]).stem + ".mp4")
        ok, duration = probe(destination)
        if ok:
            log(f"[{index}/{len(videos)}] SKIP {row['name']} (already verified, {duration / 60:.1f} min)")
            continue
        log(f"[{index}/{len(videos)}] {row['name']} ({row['size'] / 2**20:.1f} MiB source)")
        url_for_file = stream_url(target, row)
        success = False
        for attempt in range(1, args.attempts + 1):
            log(f"  attempt {attempt}/{args.attempts}")
            if download_one(url_for_file, destination, formats, args, jar_path):
                success = True
                break
            if attempt < args.attempts:
                time.sleep(15)
        if not success:
            failures.append(row["name"])
            log(f"  FAILED: {row['name']}")

    if temporary_jar:
        try:
            os.unlink(temporary_jar)
        except OSError:
            pass

    log(f"Done. {len(videos) - len(failures)} verified, {len(failures)} failed.")
    if failures:
        for name in failures:
            log(f"  failed: {name}")
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
