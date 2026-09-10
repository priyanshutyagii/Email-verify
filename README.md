# Bulk Email Verification Tool

## Features
- Upload XLSX/CSV
- Automatically finds the `Email` column
- Syntax validation
- Domain validation
- MX record validation
- Optional SMTP verification
- Disposable-domain detection
- Progress and live results
- Export results as XLSX

## Run

Requires Python 3.10+.

```bash
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000

### Notes
SMTP verification is best-effort. Many providers intentionally hide mailbox existence, rate-limit requests, or reject SMTP probes. Therefore results can be `UNKNOWN`; this is not proof that an address does not exist.

For your 81k-row file, process in batches and respect DNS/SMTP rate limits.
