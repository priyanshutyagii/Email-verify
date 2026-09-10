import csv
import io
import os
import re
import smtplib
import socket
import tempfile
import threading
import uuid
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date, datetime, time as dt_time

from flask import Flask, Request, jsonify, request, send_from_directory, send_file
from openpyxl import load_workbook, Workbook
from werkzeug.exceptions import RequestEntityTooLarge
import dns.resolver

MAX_UPLOAD_MB = 500
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024


class LargeUploadRequest(Request):
    # Flask 3.1 / Werkzeug applies this to the whole multipart body, including files.
    # None disables the in-memory cap so large XLSX/CSV uploads can stream to disk.
    max_form_memory_size = None
    max_content_length = MAX_UPLOAD_BYTES


app = Flask(__name__, static_folder="static")
app.request_class = LargeUploadRequest
app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_BYTES
app.config["MAX_FORM_MEMORY_SIZE"] = None
app.config["MAX_FORM_PARTS"] = 10_000

jobs = {}
jobs_lock = threading.Lock()
email_re = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
UPLOAD_DIR = os.path.join(tempfile.gettempdir(), "mailverify_jobs")
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Keep this list conservative. Add company-approved disposable domains as needed.
DISPOSABLE_DOMAINS = {
    "mailinator.com", "guerrillamail.com", "10minutemail.com",
    "temp-mail.org", "tempmail.com", "yopmail.com", "getnada.com",
    "trashmail.com", "sharklasers.com"
}


@app.errorhandler(413)
@app.errorhandler(RequestEntityTooLarge)
def too_large(_e):
    return jsonify({
        "error": f"File is too large. Maximum upload size is {MAX_UPLOAD_MB} MB."
    }), 413


def normalize_email(value):
    if value is None:
        return ""
    return str(value).strip().lower()


def mx_lookup(domain):
    try:
        answers = dns.resolver.resolve(domain, "MX", lifetime=5)
        records = sorted(
            [(int(r.preference), str(r.exchange).rstrip(".")) for r in answers],
            key=lambda x: x[0]
        )
        return records
    except Exception:
        # Some domains accept mail directly on A/AAAA when no MX exists.
        try:
            socket.gethostbyname(domain)
            return [(0, domain)]
        except Exception:
            return []


def smtp_check(email, mx_host, timeout=8):
    """Best-effort SMTP RCPT check. Providers may intentionally return inconclusive results."""
    try:
        with smtplib.SMTP(timeout=timeout) as smtp:
            smtp.connect(mx_host, 25)
            smtp.ehlo_or_helo_if_needed()
            try:
                if smtp.has_extn("starttls"):
                    smtp.starttls()
                    smtp.ehlo_or_helo_if_needed()
            except Exception:
                # Continue without TLS if the server rejects it.
                pass
            smtp.mail("verify@example.com")
            code, msg = smtp.rcpt(email)
            if 200 <= code < 300:
                return "VALID", "SMTP accepted recipient"
            if 500 <= code < 600:
                return "INVALID", f"SMTP rejected recipient ({code})"
            return "UNKNOWN", f"SMTP inconclusive ({code})"
    except (socket.timeout, TimeoutError):
        return "UNKNOWN", "SMTP timeout"
    except (ConnectionRefusedError, OSError) as e:
        return "UNKNOWN", f"SMTP unavailable: {type(e).__name__}"
    except Exception as e:
        return "UNKNOWN", f"SMTP unavailable: {type(e).__name__}"


def verify_email(email, do_smtp=True):
    email = normalize_email(email)
    if not email:
        return "INVALID", "Empty email"

    if not email_re.match(email):
        return "INVALID", "Invalid email format"

    local, domain = email.rsplit("@", 1)

    if len(email) > 254 or len(local) > 64:
        return "INVALID", "Email length is invalid"

    if domain in DISPOSABLE_DOMAINS:
        return "RISKY", "Disposable/temporary email domain"

    mx = mx_lookup(domain)
    if not mx:
        return "INVALID", "Domain has no reachable MX/A record"

    if not do_smtp:
        return "VALID", f"Mail domain has MX: {mx[0][1]}"

    # Try up to two MX servers.
    last_reason = "Domain accepts DNS mail but SMTP mailbox check was inconclusive"
    for _, host in mx[:2]:
        status, reason = smtp_check(email, host)
        if status in ("VALID", "INVALID"):
            return status, reason
        last_reason = reason

    return "UNKNOWN", last_reason


def process_job(job_id, emails, do_smtp):
    total = len(emails)
    ordered = [None] * total

    with jobs_lock:
        jobs[job_id]["status"] = "running"

    def work(item):
        idx, email = item
        status, reason = verify_email(email, do_smtp)
        return idx, email, status, reason

    # Conservative concurrency. Increase only after confirming your network/provider permits it.
    with ThreadPoolExecutor(max_workers=8) as pool:
        futures = [pool.submit(work, item) for item in enumerate(emails)]
        for future in as_completed(futures):
            idx, email, status, reason = future.result()
            row = {
                "email": email,
                "status": status,
                "reason": reason
            }
            ordered[idx] = row
            with jobs_lock:
                job = jobs[job_id]
                job["done"] += 1
                job["counts"][status] = job["counts"].get(status, 0) + 1
                job["recent"].append(row)
                job["ordered"] = ordered

    with jobs_lock:
        jobs[job_id]["status"] = "completed"
        jobs[job_id]["ordered"] = ordered


def excel_cell(value):
    if value is None:
        return None
    if isinstance(value, datetime):
        return value.replace(tzinfo=None) if value.tzinfo else value
    if isinstance(value, (date, dt_time, int, float, bool, str)):
        return value
    return str(value)


def copy_row(row):
    if row is None:
        return []
    return [excel_cell(c) for c in row]


def csv_cell(value):
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ", timespec="seconds")
    if isinstance(value, (date, dt_time)):
        return value.isoformat()
    return str(value)


def find_email_column(headers):
    for i, h in enumerate(headers):
        name = str(h).strip().lower() if h is not None else ""
        if name in ("email", "e-mail", "emails", "mail", "email address", "email_address"):
            return i
    # Fallback: any header that contains "email"
    for i, h in enumerate(headers):
        name = str(h).strip().lower() if h is not None else ""
        if "email" in name:
            return i
    return None


def parse_emails_from_path(path, file_ext):
    emails = []
    if file_ext == "xlsx":
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            rows = ws.iter_rows(values_only=True)
            first = next(rows, None)
            if first is None:
                return None, None, "Excel file is empty"
            headers = [str(x).strip() if x is not None else "" for x in first]
            email_idx = find_email_column(headers)
            if email_idx is None:
                return None, None, "Email column not found. Add a column named Email."
            for row in rows:
                row = row or []
                emails.append(normalize_email(row[email_idx] if email_idx < len(row) else ""))
        finally:
            wb.close()
        return headers, emails, None

    with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
        reader = csv.reader(fh)
        first = next(reader, None)
        if first is None:
            return None, None, "CSV file is empty"
        headers = [str(x).strip() if x is not None else "" for x in first]
        email_idx = find_email_column(headers)
        if email_idx is None:
            return None, None, "Email column not found. Add a column named Email."
        for row in reader:
            emails.append(normalize_email(row[email_idx] if email_idx < len(row) else ""))
    return headers, emails, None


@app.get("/")
def index():
    return send_from_directory("static", "index.html")


@app.post("/api/verify")
def start_verify():
    uploaded = request.files.get("file")
    if not uploaded:
        return jsonify({"error": "Please upload an XLSX or CSV file"}), 400

    do_smtp = request.form.get("smtp", "true").lower() == "true"
    filename = (uploaded.filename or "").lower()

    if filename.endswith(".xlsx"):
        file_ext = "xlsx"
    elif filename.endswith(".csv"):
        file_ext = "csv"
    else:
        return jsonify({"error": "Only XLSX and CSV files are supported"}), 400

    job_id = uuid.uuid4().hex
    source_path = os.path.join(UPLOAD_DIR, f"{job_id}.{file_ext}")

    try:
        uploaded.save(source_path)
        size = os.path.getsize(source_path)
        if size > MAX_UPLOAD_BYTES:
            os.remove(source_path)
            return jsonify({
                "error": f"File is too large. Maximum upload size is {MAX_UPLOAD_MB} MB."
            }), 413
        headers, emails, err = parse_emails_from_path(source_path, file_ext)
        if err:
            os.remove(source_path)
            return jsonify({"error": err}), 400
    except Exception as e:
        if os.path.exists(source_path):
            os.remove(source_path)
        return jsonify({"error": f"Could not read file: {e}"}), 400

    if not emails:
        os.remove(source_path)
        return jsonify({"error": "No email rows found in the file"}), 400

    with jobs_lock:
        jobs[job_id] = {
            "status": "queued",
            "total": len(emails),
            "done": 0,
            "recent": deque(maxlen=150),
            "ordered": [],
            "counts": {"VALID": 0, "INVALID": 0, "RISKY": 0, "UNKNOWN": 0},
            "headers": headers,
            "source_path": source_path,
            "file_ext": file_ext,
        }
    threading.Thread(target=process_job, args=(job_id, emails, do_smtp), daemon=True).start()

    return jsonify({"job_id": job_id, "total": len(emails)})


@app.get("/api/status/<job_id>")
def status(job_id):
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Job not found"}), 404
        payload = {
            "status": job["status"],
            "total": job["total"],
            "done": job["done"],
            "counts": dict(job["counts"]),
            "results": list(job["recent"]),
        }
    return jsonify(payload)


def send_bytes(data, filename, mimetype):
    return send_file(
        io.BytesIO(data),
        as_attachment=True,
        download_name=filename,
        mimetype=mimetype,
    )


def iter_source_rows(path, file_ext):
    if file_ext == "csv":
        with open(path, "r", encoding="utf-8-sig", errors="replace", newline="") as fh:
            reader = csv.reader(fh)
            next(reader, None)
            for row in reader:
                yield copy_row(row)
        return

    wb = load_workbook(path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = ws.iter_rows(values_only=True)
        next(rows, None)
        for row in rows:
            yield copy_row(row)
    finally:
        wb.close()


def build_original_file_from_rows(headers, rows_iter, file_ext):
    if file_ext == "csv":
        text = io.StringIO()
        writer = csv.writer(text, lineterminator="\n")
        writer.writerow(headers)
        for row in rows_iter:
            writer.writerow([csv_cell(c) for c in row])
        return (
            text.getvalue().encode("utf-8-sig"),
            "valid_emails.csv",
            "text/csv",
        )

    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Emails")
    ws.append(list(headers))
    for row in rows_iter:
        ws.append(row)
    output = io.BytesIO()
    wb.save(output)
    return (
        output.getvalue(),
        "valid_emails.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


def export_valid_file(job):
    """Download only VALID rows, keeping the original uploaded file columns."""
    headers = list(job.get("headers") or [])
    ordered = list(job.get("ordered") or [])
    file_ext = job.get("file_ext") or "xlsx"
    source_path = job.get("source_path")
    valid_count = job.get("counts", {}).get("VALID", 0)

    if not headers:
        return jsonify({"error": "Original columns were not saved. Please upload the file and run verification again."}), 400
    if not source_path or not os.path.exists(source_path):
        return jsonify({"error": "Original file is no longer available. Please upload and verify again."}), 400
    if valid_count < 1:
        return jsonify({"error": f"No valid emails to download ({valid_count} valid)."}), 400

    def valid_rows():
        for idx, row in enumerate(iter_source_rows(source_path, file_ext)):
            result = ordered[idx] if idx < len(ordered) else None
            if result and result.get("status") == "VALID":
                yield row

    data, filename, mimetype = build_original_file_from_rows(headers, valid_rows(), file_ext)
    if not data or (file_ext == "csv" and data.decode("utf-8-sig").count("\n") <= 1):
        return jsonify({"error": f"No valid emails to download ({valid_count} valid)."}), 400
    return send_bytes(data, filename, mimetype)


def export_all_file(job):
    rows = [r for r in (job.get("ordered") or []) if r]
    wb = Workbook(write_only=True)
    ws = wb.create_sheet("Verified Emails")
    ws.append(["Email", "Status", "Reason"])
    for item in rows:
        ws.append([item["email"], item["status"], item["reason"]])
    output = io.BytesIO()
    wb.save(output)
    return send_bytes(
        output.getvalue(),
        "verified_emails.xlsx",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )


@app.get("/api/export/<job_id>")
@app.get("/api/export/<job_id>/<kind>")
def export(job_id, kind=None):
    only = (kind or request.args.get("only") or "all").strip().lower()
    with jobs_lock:
        job = jobs.get(job_id)
        if not job:
            return jsonify({"error": "Job not found. Run verification again."}), 404
        if job["status"] != "completed":
            return jsonify({"error": "Verification is not completed"}), 400
        snapshot = {
            "headers": list(job.get("headers") or []),
            "ordered": list(job.get("ordered") or []),
            "source_path": job.get("source_path"),
            "file_ext": job.get("file_ext") or "xlsx",
            "counts": dict(job.get("counts") or {}),
        }

    if only == "valid":
        return export_valid_file(snapshot)
    return export_all_file(snapshot)


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
