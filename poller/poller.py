#!/usr/bin/env python3
"""print-poller: email-to-print bridge.

Watches SOURCE_FOLDER (a Proton-via-Bridge mailbox) for messages addressed to
PRINT_TO from allow-listed senders and prints them to a CUPS queue:
  - PDF/image attachments print natively
  - Office-doc attachments convert via LibreOffice headless
  - if no printable attachment and PRINT_BODY=true, the email body (HTML/text)
    is rendered to PDF and printed (forward-an-email-to-print-it)
Every processed message is MOVED out of SOURCE_FOLDER (-> PROCESSED_FOLDER on
success, REJECTED_FOLDER otherwise) which is the idempotency guard.
Stdlib only + external `lp` and `soffice`. Fail-closed on the allow-list.
"""
import os, ssl, sys, time, email, subprocess, tempfile, logging, mimetypes
from email.header import decode_header, make_header
from email.utils import parseaddr, getaddresses, formatdate, make_msgid
from email.message import EmailMessage
import imaplib, smtplib, threading, json as _json
from http.server import BaseHTTPRequestHandler, HTTPServer

def env(k, d=None, req=False):
    v = os.environ.get(k, d)
    if req and not v:
        logging.critical("missing required env %s", k); sys.exit(2)
    return v

logging.basicConfig(level=os.environ.get("LOG_LEVEL", "INFO").upper(),
                    format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("print-poller")

IMAP_HOST = env("IMAP_HOST", "127.0.0.1"); IMAP_PORT = int(env("IMAP_PORT", "1143"))
IMAP_USER = env("IMAP_USER", req=True);    IMAP_PASS = env("IMAP_PASS", req=True)
SMTP_HOST = env("SMTP_HOST", "127.0.0.1"); SMTP_PORT = int(env("SMTP_PORT", "1025"))
PRINT_TO  = env("PRINT_TO", req=True).lower()
SOURCE_FOLDER = env("SOURCE_FOLDER", "INBOX")
ALLOWED   = set(a.strip().lower() for a in env("ALLOWED_SENDERS", "").split(",") if a.strip())
PRINTER   = env("PRINTER", req=True)
CUPS_SERVER = env("CUPS_SERVER", "127.0.0.1:631")
PROCESSED_FOLDER = env("PROCESSED_FOLDER", "Folders/Printed")
REJECTED_FOLDER  = env("REJECTED_FOLDER", "Folders/Print-Rejected")
POLL_INTERVAL = int(env("POLL_INTERVAL", "60"))
MAX_MB    = float(env("MAX_ATTACH_MB", "25"))
CONFIRM_REPLY = env("CONFIRM_REPLY", "true").lower() == "true"
PRINT_BODY = env("PRINT_BODY", "true").lower() == "true"
DRY_RUN   = env("DRY_RUN", "false").lower() == "true"
REQUIRE_AUTH_PASS = env("REQUIRE_AUTH_PASS", "false").lower() == "true"
TLS_VERIFY = env("TLS_VERIFY", "true").lower() == "true"
IMAP_SSL  = env("IMAP_SSL", "false").lower() == "true"
REPLY_ON_REJECT = env("REPLY_ON_REJECT", "false").lower() == "true"
PRINT_OPTS = {"sides": env("SIDES", "one-sided"), "media": env("MEDIA", "letter")}
HEALTH_BIND = env("HEALTH_BIND", "127.0.0.1")

_LOCAL_HOSTS = {"127.0.0.1", "localhost", "::1"}
if not TLS_VERIFY and not (IMAP_HOST in _LOCAL_HOSTS and SMTP_HOST in _LOCAL_HOSTS):
    logging.critical("TLS_VERIFY=false is only allowed for localhost (bridge). "
                     "IMAP_HOST=%s SMTP_HOST=%s", IMAP_HOST, SMTP_HOST); sys.exit(2)

os.environ["CUPS_SERVER"] = CUPS_SERVER
HEALTH_PORT = int(env("HEALTH_PORT", "2631"))
_state = {"started": time.time(), "last_poll": None, "last_poll_ok": False, "printed_total": 0, "rejected_total": 0, "errors_total": 0}

class _Health(BaseHTTPRequestHandler):
    def log_message(self, *a): pass
    def do_GET(self):
        if self.path.rstrip("/") not in ("/health", ""):
            self.send_response(404); self.end_headers(); return
        ok = _state["last_poll_ok"] and _state["last_poll"] is not None
        status = "ok" if ok else ("starting" if _state["last_poll"] is None else "degraded")
        lp = _state["last_poll"]
        body = _json.dumps({"status": status, "printer": PRINTER, "source": SOURCE_FOLDER,
            "last_poll": None if lp is None else time.strftime("%Y-%m-%dT%H:%M:%S", time.localtime(lp)),
            "printed_total": _state["printed_total"], "rejected_total": _state["rejected_total"],
            "errors_total": _state["errors_total"], "uptime_s": int(time.time() - _state["started"])}).encode()
        self.send_response(200); self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body))); self.end_headers(); self.wfile.write(body)

def _start_health():
    try: HTTPServer((HEALTH_BIND, HEALTH_PORT), _Health).serve_forever()
    except Exception as e: log.warning("health server failed: %s", e)

NATIVE_CTYPES = {"application/pdf","image/jpeg","image/png","image/gif","image/webp","image/bmp","image/tiff"}
NATIVE_EXT = {".pdf",".jpg",".jpeg",".png",".gif",".webp",".bmp",".tif",".tiff"}
OFFICE_EXT = {".doc",".docx",".odt",".rtf",".xls",".xlsx",".ods",".ppt",".pptx",".odp",".txt",".csv"}

def dh(s):
    try: return str(make_header(decode_header(s or "")))
    except Exception: return s or ""

def addrs(msg, *headers):
    vals = []
    for h in headers: vals += msg.get_all(h, [])
    return [a.lower() for _, a in getaddresses(vals) if a]

def _tls_ctx():
    ctx = ssl.create_default_context()
    if not TLS_VERIFY:
        ctx.check_hostname = False; ctx.verify_mode = ssl.CERT_NONE
    return ctx

def imap_connect():
    if IMAP_SSL:
        M = imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT, ssl_context=_tls_ctx())
    else:
        M = imaplib.IMAP4(IMAP_HOST, IMAP_PORT); M.starttls(ssl_context=_tls_ctx())
    M.login(IMAP_USER, IMAP_PASS)
    return M

def ensure_folder(M, name):
    try: M.create(name)
    except Exception: pass

def move(M, uid, dest):
    ensure_folder(M, dest)
    typ, _ = M.uid("MOVE", uid, dest)
    if typ != "OK":
        M.uid("COPY", uid, dest); M.uid("STORE", uid, "+FLAGS", r"(\Deleted)"); M.expunge()

def print_file(path, opts):
    cmd = ["lp", "-d", PRINTER]
    for k, v in opts.items(): cmd += ["-o", f"{k}={v}"]
    cmd += ["--", path]
    log.info("printing %s (%s)", os.path.basename(path), " ".join(cmd))
    if DRY_RUN: return True, "dry-run"
    try:
        r = subprocess.run(cmd, capture_output=True, text=True, timeout=120)
        return r.returncode == 0, (r.stdout + r.stderr).strip()
    except Exception as e:
        return False, str(e)

def to_pdf(path, workdir):
    try:
        r = subprocess.run(["soffice","--headless","--convert-to","pdf","--outdir",workdir,path],
                           capture_output=True, text=True, timeout=180)
    except Exception as e:
        log.error("soffice exec failed: %s", e); return None
    out = os.path.join(workdir, os.path.splitext(os.path.basename(path))[0] + ".pdf")
    if r.returncode == 0 and os.path.exists(out): return out
    log.error("soffice failed for %s: %s", path, (r.stdout + r.stderr).strip()); return None

def render_body(msg, wd):
    html = text = None
    for part in msg.walk():
        if part.get_content_maintype() == "multipart": continue
        if "attachment" in (part.get("Content-Disposition") or "").lower(): continue
        ct = part.get_content_type()
        if ct == "text/html" and html is None: html = part.get_payload(decode=True)
        elif ct == "text/plain" and text is None: text = part.get_payload(decode=True)
    if html:
        src = os.path.join(wd, "body.html"); open(src, "wb").write(html)
    elif text:
        src = os.path.join(wd, "body.txt"); open(src, "wb").write(text)
    else:
        return None
    return to_pdf(src, wd)

def auth_ok(msg):
    if not REQUIRE_AUTH_PASS: return True
    ar = " ".join(msg.get_all("Authentication-Results", [])).lower()
    return "spf=pass" in ar or "dkim=pass" in ar

def send_reply(orig, to_addr, status, detail):
    if not CONFIRM_REPLY or not to_addr: return
    try:
        m = EmailMessage(); m["From"] = IMAP_USER; m["To"] = to_addr
        m["Subject"] = f"Print {status}: " + dh(orig.get("Subject", "(no subject)"))
        if orig.get("Message-ID"): m["In-Reply-To"] = orig["Message-ID"]
        m["Date"] = formatdate(localtime=True); m["Message-ID"] = make_msgid(); m.set_content(detail)
        s = smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30); s.starttls(context=_tls_ctx()); s.login(IMAP_USER, IMAP_PASS)
        s.send_message(m); s.quit()
    except Exception as e:
        log.warning("confirm reply failed: %s", e)

def process(M, uid):
    typ, data = M.uid("FETCH", uid, "(RFC822)")
    if typ != "OK" or not data or not data[0]: return
    msg = email.message_from_bytes(data[0][1])
    if PRINT_TO not in addrs(msg, "To", "Cc", "Delivered-To", "X-Original-To", "X-Forwarded-To"):
        # In the watch folder but not addressed to the print alias (misfiled,
        # or Bcc which carries no header). Move it out so it isn't refetched
        # every poll. If you legitimately Bcc your printer, stop.
        log.warning("REJECT not addressed to %s (in %s anyway)", PRINT_TO, SOURCE_FOLDER)
        move(M, uid, REJECTED_FOLDER); _state["rejected_total"] += 1; return
    frm = parseaddr(msg.get("From", ""))[1].lower()
    subj = dh(msg.get("Subject", "(no subject)"))
    log.info("candidate from=%s subj=%r", frm, subj)
    if (not ALLOWED) or (frm not in ALLOWED):
        log.warning("REJECT sender not allowed: %s", frm)
        move(M, uid, REJECTED_FOLDER)
        if REPLY_ON_REJECT: send_reply(msg, frm, "rejected (sender not allowed)", "Your address is not on the print allow-list.")
        _state["rejected_total"] += 1; return
    if not auth_ok(msg):
        log.warning("REJECT failed SPF/DKIM: %s", frm); move(M, uid, REJECTED_FOLDER); return
    printed, errors = [], []
    with tempfile.TemporaryDirectory() as wd:
        idx = 0
        for part in msg.walk():
            if part.get_content_maintype() == "multipart": continue
            fn = dh(part.get_filename() or ""); disp = (part.get("Content-Disposition") or "").lower()
            ctype = (part.get_content_type() or "").lower()
            if not fn and "attachment" not in disp: continue
            payload = part.get_payload(decode=True)
            if not payload: continue
            if len(payload) > MAX_MB * 1024 * 1024: errors.append(f"{fn or ctype}: exceeds {MAX_MB}MB"); continue
            ext = os.path.splitext(fn)[1].lower()
            raw = os.path.basename(fn) if fn else "attachment" + (mimetypes.guess_extension(ctype) or ".bin")
            clean = "".join(c if c.isalnum() or c in "._- " else "_" for c in raw)[-120:].strip(". ") or "attachment.bin"
            idx += 1
            safe = os.path.join(wd, f"{idx:02d}-{clean}")
            with open(safe, "wb") as f: f.write(payload)
            if ctype in NATIVE_CTYPES or ext in NATIVE_EXT: target = safe
            elif ext in OFFICE_EXT:
                target = to_pdf(safe, wd)
                if not target: errors.append(f"{fn}: conversion failed"); continue
            else:
                log.info("skip unsupported attachment %s (%s)", fn, ctype); continue
            ok, detail = print_file(target, PRINT_OPTS)
            (printed if ok else errors).append(fn if ok else f"{fn}: {detail}")
        # No printable attachment -> print the email body itself (forward-to-print)
        if not printed and PRINT_BODY:
            body_pdf = render_body(msg, wd)
            if body_pdf:
                ok, detail = print_file(body_pdf, PRINT_OPTS)
                (printed if ok else errors).append("email body" if ok else f"email body: {detail}")
            else:
                errors.append("no printable attachment and no renderable body")
    if printed and not errors:
        move(M, uid, PROCESSED_FOLDER); send_reply(msg, frm, "ok", "Printed: " + ", ".join(printed)); _state["printed_total"] += 1; log.info("DONE printed=%s", printed)
    elif printed:
        move(M, uid, PROCESSED_FOLDER); send_reply(msg, frm, "partial", "Printed: " + ", ".join(printed) + "\nErrors: " + "; ".join(errors)); log.warning("PARTIAL printed=%s errors=%s", printed, errors)
    else:
        move(M, uid, REJECTED_FOLDER); send_reply(msg, frm, "failed", "Nothing could be printed. " + "; ".join(errors)); _state["rejected_total"] += 1; log.warning("NO-PRINT errors=%s", errors)

def poll_once(M):
    M.select(SOURCE_FOLDER)
    # Everything in the dedicated folder is a candidate; process() filters on
    # PRINT_TO. (Header SEARCH missed Cc/X-Original-To-only routing.)
    typ, data = M.uid("SEARCH", None, "ALL")
    uids = data[0].split() if typ == "OK" and data and data[0] else []
    if uids: log.info("%d candidate message(s) in %s", len(uids), SOURCE_FOLDER)
    for uid in uids:
        try: process(M, uid)
        except Exception as e:
            _state["errors_total"] += 1; log.exception("error on uid=%s: %s", uid, e)
    _state["last_poll"] = time.time(); _state["last_poll_ok"] = True

def main():
    log.info("print-poller up: user=%s source=%s printer=%s match=%s print_body=%s allow=%s dry=%s",
             IMAP_USER, SOURCE_FOLDER, PRINTER, PRINT_TO, PRINT_BODY, sorted(ALLOWED) or "(NONE-fail-closed)", DRY_RUN)
    if not ALLOWED: log.warning("ALLOWED_SENDERS empty -> fail-closed")
    threading.Thread(target=_start_health, daemon=True).start(); log.info("health endpoint on :%d/health", HEALTH_PORT)
    while True:
        try:
            M = imap_connect()
            try: poll_once(M)
            finally:
                try: M.logout()
                except Exception: pass
        except Exception as e:
            log.exception("poll cycle failed: %s", e)
        time.sleep(POLL_INTERVAL)

if __name__ == "__main__":
    main()
