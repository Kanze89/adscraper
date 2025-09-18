# -*- coding: utf-8 -*-
"""
shipping.py — packaging, Excel export, email, and GitHub push.

- Builds XLSX from combined CSV with clickable links ("example_link" column).
  It tries, in order:
    1) PUBLIC_BASE_URL + a relative path computed from image_path (if possible)
    2) file:// local path (works if the file is on the viewer's machine or shared drive)
- Zips last 7 days of banner_screenshots.
- Sends weekly email (Monday only) with zip + ledger + xlsx attached.
- Commits and pushes to Git after each run if env configured.

Environment (set in Windows “User variables” or similar):
  GIT_REPO_DIR            e.g., C:/Users/tuguldur.kh/Downloads/adscraper-full-code
  GIT_REMOTE_NAME         default 'origin'
  GIT_BRANCH              default 'main'

  SMTP_HOST               e.g., smtp.gmail.com
  SMTP_PORT               e.g., 587
  SMTP_USER               your_smtp_username
  SMTP_PASS               your_app_password
  MAIL_FROM               "Ad Bot <bot@example.com>"
  MAIL_TO                 comma-separated list (e.g., "a@b.com,c@d.com")

  PUBLIC_BASE_URL         e.g., https://github.com/<USER>/<REPO>/blob/main
  OUTPUT_ROOT             (optional) override root for building relative file paths
"""

import os, csv, zipfile, smtplib, mimetypes, traceback, subprocess
from datetime import datetime, timedelta
from email.message import EmailMessage
from pathlib import Path

from openpyxl import Workbook
from openpyxl.utils import get_column_letter

def _guess_output_root() -> str:
    # Prefer explicit env; else infer from typical structure
    env_root = os.getenv("OUTPUT_ROOT", "").strip()
    if env_root:
        return os.path.abspath(env_root)
    # Fallback: assume folder next to run.py called banner_screenshots
    here = Path(__file__).resolve().parent
    candidate = here / "banner_screenshots"
    return str(candidate)

def _to_rel(path_str: str, output_root: str) -> str:
    try:
        rel = os.path.relpath(path_str, output_root)
        return rel.replace("\\", "/")
    except Exception:
        return ""

def _public_url_from_rel(rel_path: str) -> str:
    base = (os.getenv("PUBLIC_BASE_URL") or "").rstrip("/")
    if not base or not rel_path:
        return ""
    return base + "/" + rel_path.lstrip("/")

def _file_url(local_path: str) -> str:
    # Excel will open file:// links when accessible
    p = Path(local_path).resolve()
    return "file:///" + str(p).replace("\\", "/")

def build_xlsx_from_csv(csv_path: str, xlsx_path: str) -> None:
    """
    Convert the combined CSV into XLSX, making the *existing* path column
    clickable for recipients (no local C:\ paths shown).

    Rules:
      - Display: a nice relative path like  news.mn/2025-09-18/file.png
                 (or just filename if we can't compute rel)
      - Hyperlink: public URL, preferring RAW GitHub if available.
      - We do NOT modify your scrapers; only the Excel output.
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "banners"

    os.makedirs(os.path.dirname(xlsx_path), exist_ok=True)

    if not os.path.exists(csv_path):
        wb.save(xlsx_path)
        print("[XLSX] CSV not found; created empty workbook.")
        return

    # Where screenshots live locally (to compute a relative path)
    output_root = _guess_output_root()

    # Public link bases (pick RAW if provided; else derive RAW from PUBLIC_BASE_URL if possible)
    raw_base = (os.getenv("RAW_BASE_URL") or "").rstrip("/")
    public_base = (os.getenv("PUBLIC_BASE_URL") or "").rstrip("/")
    if not raw_base and public_base and "/blob/" in public_base:
        # Convert https://github.com/u/r/blob/branch  -> https://raw.githubusercontent.com/u/r/branch
        try:
            parts = public_base.split("/blob/")
            left = parts[0]                      # https://github.com/<USER>/<REPO>
            branch = parts[1]                    # <BRANCH>
            # transform domain + path
            after = left.replace("https://github.com/", "https://raw.githubusercontent.com/")
            raw_base = after + "/" + branch
        except Exception:
            raw_base = ""  # fallback to public_base if conversion fails

    # Which column should be clickable?
    # Priority: example_path, then image_path
    clickable_cols = ["example_path", "image_path"]

    with open(csv_path, "r", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        fieldnames = reader.fieldnames or []
        ws.append(fieldnames)  # keep the same columns; we’ll decorate one as hyperlink

        # find the path column that exists in this CSV
        click_col_idx = None
        click_col_name = None
        for name in clickable_cols:
            if name in fieldnames:
                click_col_name = name
                click_col_idx = fieldnames.index(name) + 1  # 1-based for Excel
                break

        for row in reader:
            # First append raw values
            values = [row.get(k, "") for k in fieldnames]
            ws.append(values)

            # If we have a path column, rewrite its cell: display rel, hyperlink public URL
            if click_col_idx:
                original_path = row.get(click_col_name, "") or ""
                display_text = original_path
                hyperlink = ""

                # Try to compute a relative path under output_root
                if original_path:
                    rel = _to_rel(original_path, output_root)
                    if rel:
                        display_text = rel  # nicer than C:\...
                    else:
                        # as a last resort, just show the filename
                        try:
                            display_text = os.path.basename(original_path)
                        except Exception:
                            display_text = original_path

                    # Prefer RAW base (direct file bytes) if set
                    if raw_base and rel:
                        hyperlink = raw_base.rstrip("/") + "/" + rel.lstrip("/")
                    elif public_base and rel:
                        # Fall back to PUBLIC_BASE_URL (blob URLs load GitHub viewer)
                        hyperlink = public_base.rstrip("/") + "/" + rel.lstrip("/")
                    else:
                        # As a last fallback (not ideal for other people),
                        # keep a local file:// link if nothing public configured
                        hyperlink = _file_url(original_path)

                # Write back into the sheet: replace the visible text + hyperlink
                cell = ws.cell(row=ws.max_row, column=click_col_idx)
                cell.value = display_text
                if hyperlink:
                    cell.hyperlink = hyperlink
                    cell.style = "Hyperlink"

        # Make it readable
        for i, col in enumerate(fieldnames, 1):
            ws.column_dimensions[get_column_letter(i)].width = 30

    wb.save(xlsx_path)
    print(f"[XLSX] Wrote {xlsx_path} with clickable '{click_col_name or 'N/A'}' links")


def zip_last_7_days(root_screenshots: str, out_zip_path: str) -> None:
    cutoff = datetime.now().date() - timedelta(days=7)
    os.makedirs(os.path.dirname(out_zip_path), exist_ok=True)
    with zipfile.ZipFile(out_zip_path, "w", zipfile.ZIP_DEFLATED) as zf:
        for site in ("gogo.mn", "ikon.mn", "news.mn"):
            site_root = os.path.join(root_screenshots, site)
            if not os.path.isdir(site_root):
                continue
            for day in os.listdir(site_root):
                # day folder like YYYY-MM-DD
                try:
                    d = datetime.strptime(day, "%Y-%m-%d").date()
                except Exception:
                    continue
                if d >= cutoff:
                    folder = os.path.join(site_root, day)
                    for dirpath, _, filenames in os.walk(folder):
                        for fn in filenames:
                            full = os.path.join(dirpath, fn)
                            rel  = os.path.relpath(full, root_screenshots)
                            zf.write(full, rel)
    print(f"[ZIP] Wrote {out_zip_path}")

def _attach_file(msg: EmailMessage, file_path: str):
    ctype, encoding = mimetypes.guess_type(file_path)
    if ctype is None or encoding is not None:
        ctype = "application/octet-stream"
    maintype, subtype = ctype.split("/", 1)
    with open(file_path, "rb") as f:
        msg.add_attachment(f.read(), maintype=maintype, subtype=subtype, filename=os.path.basename(file_path))

def send_email(subject: str, body: str, attachments: list[str]) -> None:
    smtp_host = os.getenv("SMTP_HOST", "")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASS", "")
    mail_from = os.getenv("MAIL_FROM", smtp_user)
    mail_to   = [x.strip() for x in os.getenv("MAIL_TO", "").split(",") if x.strip()]

    if not (smtp_host and smtp_user and smtp_pass and mail_to):
        print("[WARN] Email not sent (SMTP env missing)")
        return

    msg = EmailMessage()
    msg["From"] = mail_from
    msg["To"] = ", ".join(mail_to)
    msg["Subject"] = subject
    msg.set_content(body)

    for p in attachments:
        try:
            if os.path.exists(p):
                _attach_file(msg, p)
        except Exception:
            traceback.print_exc()

    with smtplib.SMTP(smtp_host, smtp_port) as s:
        s.starttls()
        s.login(smtp_user, smtp_pass)
        s.send_message(msg)
    print("[MAIL] Sent email to:", msg["To"])

def git_commit_and_push(repo_dir: str, message: str) -> None:
    """
    Uses local git in PATH. Repo should already have remote + auth set.
    Env (optional): GIT_REMOTE_NAME, GIT_BRANCH
    """
    remote = os.getenv("GIT_REMOTE_NAME", "origin")
    branch = os.getenv("GIT_BRANCH", "main")
    def run(*cmd):
        subprocess.run(cmd, cwd=repo_dir, check=False)
    run("git", "add", "-A")
    run("git", "commit", "-m", message)
    run("git", "push", remote, branch)
    print("[GIT] Push attempted to", remote, branch)
