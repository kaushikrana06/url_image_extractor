# backend/app.py
import io
import os
import re
import json
import base64
import zipfile
import asyncio
import unicodedata
from typing import List, Optional, Dict, Any, Tuple
from urllib.parse import urlparse, parse_qs, quote

import httpx
from PIL import Image, UnidentifiedImageError, ImageFile
from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.responses import StreamingResponse
from fastapi.middleware.cors import CORSMiddleware
from fastapi.staticfiles import StaticFiles
from openpyxl import load_workbook

# Pillow safety
ImageFile.LOAD_TRUNCATED_IMAGES = True

# -----------------------------------------------------------------------------
# App setup
# -----------------------------------------------------------------------------
app = FastAPI(title="XLSX Image Zipper")

# CORS (harmless even if serving UI + API from same origin)
ALLOWED = os.getenv("CORS_ORIGIN", "*")
origins = [o.strip() for o in ALLOWED.split(",") if o.strip()]
app.add_middleware(
    CORSMiddleware,
    allow_origins=origins,
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
    max_age=3600,
)

# A browser-y UA helps some CDNs (Drive/Dropbox)
DEFAULT_HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0 Safari/537.36"
    )
}

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def slugify(text: str, max_len: int = 70) -> str:
    text = unicodedata.normalize("NFKD", str(text))
    text = text.encode("ascii", "ignore").decode("ascii")
    text = re.sub(r"[^A-Za-z0-9._ -]+", "_", text)
    text = re.sub(r"\s+", "_", text)
    text = re.sub(r"_+", "_", text).strip("._ ")
    return (text or "untitled")[:max_len]

def ascii_fallback_filename(s: str) -> str:
    norm = unicodedata.normalize("NFKD", s)
    ascii_only = norm.encode("ascii", "ignore").decode("ascii")
    ascii_only = re.sub(r"[^A-Za-z0-9._ -]+", "_", ascii_only).strip()
    if not ascii_only:
        ascii_only = "images.zip"
    if not ascii_only.lower().endswith(".zip"):
        ascii_only += ".zip"
    return ascii_only

def extract_url_from_cell(cell) -> str:
    """openpyxl: hyperlink.target, HYPERLINK() formula, or plain text URL."""
    try:
        if getattr(cell, "hyperlink", None) and getattr(cell.hyperlink, "target", None):
            return str(cell.hyperlink.target).strip()
    except Exception:
        pass

    v = cell.value
    if not isinstance(v, str):
        return ""

    s = v.strip()
    if not s:
        return ""

    m = re.search(r'HYPERLINK\(\s*"([^"]+)"', s, re.I)
    if m:
        return m.group(1).strip()

    if s.startswith(("http://", "https://")):
        return s

    return ""

def strip_tracking_params(url: str) -> str:
    try:
        p = urlparse(url)
        if not p.query:
            return url
        qs = parse_qs(p.query)
        for k in ["utm_source", "utm_medium", "utm_campaign", "utm_term", "utm_content", "fbclid", "gclid"]:
            qs.pop(k, None)
        new_q = "&".join(f"{k}={v[0]}" for k, v in qs.items())
        return p._replace(query=new_q).geturl()
    except Exception:
        return url

def drive_fetch_plan(parsed, original_qs) -> List[str]:
    """
    Candidates for Google Drive, preserving resourcekey if present.
    Order: original -> direct download.
    """
    candidates: List[str] = [parsed.geturl()]

    file_id = None
    if parsed.path.startswith("/file/d/"):
        parts = parsed.path.split("/")
        if len(parts) > 3:
            file_id = parts[3]
    elif parsed.path.startswith("/open") and "id" in original_qs:
        file_id = original_qs["id"][0]
    elif parsed.path.startswith("/uc") and "id" in original_qs:
        file_id = original_qs["id"][0]

    if file_id:
        rk = original_qs.get("resourcekey", [None])[0]
        q = f"export=download&id={file_id}"
        if rk:
            q += f"&resourcekey={rk}"
        candidates.append(f"https://drive.google.com/uc?{q}")

    # de-dup
    seen = set()
    uniq = []
    for u in candidates:
        if u not in seen:
            uniq.append(u)
            seen.add(u)
    return uniq

def build_fetch_plan(raw_url: str) -> Tuple[str, List[str]]:
    """
    (display_url, fetch_candidates[])
    display_url: original link (opens in user's logged-in browser)
    candidates: server-side URLs to try for downloading content
    """
    display = raw_url.strip()
    if not display:
        return "", []

    try:
        p = urlparse(display)
        host = (p.netloc or "").lower()
        qs = parse_qs(p.query or "")

        if "drive.google.com" in host:
            return display, drive_fetch_plan(p, qs)

        if "dropbox.com" in host:
            cand = p
            if "dl=" in (p.query or ""):
                q = re.sub(r"dl=\d", "dl=1", p.query)
                cand = p._replace(query=q)
            return display, [display, cand.geturl()]

        stripped = strip_tracking_params(display)
        if stripped != display:
            return display, [display, stripped]
        return display, [display]
    except Exception:
        return display, [display]

async def fetch_one(client: httpx.AsyncClient, urls: List[str], sem: asyncio.Semaphore) -> bytes | None:
    """Try candidates in order; return first successful content, else None."""
    for u in urls:
        async with sem:
            try:
                r = await client.get(u)
                r.raise_for_status()
                return r.content
            except Exception:
                continue
    return None

def find_final_col(ws) -> Optional[Tuple[int, int]]:
    """
    Find (header_row, col_idx) where text starts with 'final' (case-insensitive).
    Scan up to 500 rows to tolerate banners/multi-row headings.
    """
    max_r = min(ws.max_row or 1, 500)
    for r in range(1, max_r + 1):
        for row in ws.iter_rows(min_row=r, max_row=r, values_only=False):
            for cell in row:
                val = cell.value
                if isinstance(val, str) and val.strip().lower().startswith("final"):
                    return (r, cell.column)
    return None

# -----------------------------------------------------------------------------
# API
# -----------------------------------------------------------------------------
@app.get("/healthz")
def healthz():
    return {"ok": True}

@app.post("/api/upload")
async def upload_excel(file: UploadFile = File(...)):
    if not file.filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Please upload an .xlsx file.")

    data = await file.read()
    if not data:
        raise HTTPException(status_code=400, detail="Empty file.")

    try:
        # not read_only so hyperlinks + formulas are available
        wb = load_workbook(io.BytesIO(data), data_only=False, read_only=False, keep_links=True)
    except Exception:
        raise HTTPException(status_code=400, detail="Unable to read the Excel file. Is it valid .xlsx?")

    workbook_name = os.path.splitext(os.path.basename(file.filename))[0]
    zip_buf = io.BytesIO()
    failed: List[Dict[str, Any]] = []  # {sheet, name, url (original/display)}

    with zipfile.ZipFile(zip_buf, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # Reuse one HTTP client across all sheets
        async with httpx.AsyncClient(
            follow_redirects=True,
            headers=DEFAULT_HEADERS,
            timeout=45.0
        ) as client:

            for sheet_name in wb.sheetnames:
                ws = wb[sheet_name]
                loc = find_final_col(ws)
                if not loc:
                    continue
                header_row, col_idx = loc

                # Collect (display_url, fetch_candidates[])
                plans: List[Tuple[str, List[str]]] = []
                for row in ws.iter_rows(
                    min_row=header_row + 1,
                    max_row=ws.max_row,
                    min_col=col_idx,
                    max_col=col_idx
                ):
                    raw = extract_url_from_cell(row[0])
                    if not raw:
                        continue
                    display, candidates = build_fetch_plan(raw)
                    if candidates:
                        plans.append((display, candidates))

                if not plans:
                    continue

                inner_dir = f"{slugify(workbook_name)}/{slugify(sheet_name)}/images"

                # concurrent downloads (per-sheet)
                sem = asyncio.Semaphore(10)
                results = await asyncio.gather(
                    *(fetch_one(client, cand, sem) for _, cand in plans),
                    return_exceptions=False,
                )

                # Save as final_image_N.jpg per sheet
                seq = 1
                for (display_url, _), content in zip(plans, results):
                    name = f"final_image_{seq}.jpg"
                    if not content:
                        failed.append({"sheet": sheet_name, "name": name, "url": display_url})
                        seq += 1
                        continue
                    try:
                        img = Image.open(io.BytesIO(content))
                        if img.mode not in ("RGB", "L"):
                            img = img.convert("RGB")
                        out = io.BytesIO()
                        img.save(out, format="JPEG", quality=95)
                        zf.writestr(f"{inner_dir}/{name}", out.getvalue())
                    except (UnidentifiedImageError, Exception):
                        failed.append({"sheet": sheet_name, "name": name, "url": display_url})
                    finally:
                        seq += 1

        # include failures inside the zip
        if failed:
            zf.writestr(
                f"{slugify(workbook_name)}/failed.json",
                json.dumps(failed, ensure_ascii=False, indent=2),
            )

    # Response with Unicode-safe filename and failures header
    zip_buf.seek(0)
    original_zip = f"{workbook_name}_images.zip"
    ascii_name = ascii_fallback_filename(original_zip)
    filename_star = quote(original_zip, safe="")
    failed_b64 = base64.b64encode(json.dumps(failed).encode("utf-8")).decode("ascii")
    cd = f'attachment; filename="{ascii_name}"; filename*=UTF-8\'\'{filename_star}'

    return StreamingResponse(
        zip_buf,
        media_type="application/zip",
        headers={
            "Content-Disposition": cd,
            "X-Failed-Json": failed_b64,
        },
    )

# -----------------------------------------------------------------------------
# Static UI (served by FastAPI) - mount AFTER API routes
# -----------------------------------------------------------------------------
if os.path.isdir("static"):
    app.mount("/", StaticFiles(directory="static", html=True), name="static")
