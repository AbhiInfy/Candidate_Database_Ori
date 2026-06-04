# IMAP CV Agent Steps

Use this version for custom-domain email accounts like:

```text
careers@emotifzone.com
```

This does not use Azure or Microsoft Graph.

## What You Need

Ask your email provider or hosting provider for:

- IMAP server
- IMAP port
- Username
- Password or app password

Common examples:

```text
imap.gmail.com
imap.zoho.com
imap.hostinger.com
mail.emotifzone.com
```

The port is usually:

```text
993
```

## Install Packages

In VS Code terminal:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Run The Agent

Replace the IMAP server if your provider gives a different one:

```powershell
python .\imap_cv_agent.py --imap-server "mail.emotifzone.com" --username "careers@emotifzone.com" --folder "INBOX"
```

The script will ask for the email password or app password.

It will create:

```text
candidate_details.xlsx
```

By default, the agent scans the latest 500 emails in the selected folder. To scan more emails, use:

```powershell
python .\imap_cv_agent.py --imap-server "mail.emotifzone.com" --username "careers@emotifzone.com" --folder "INBOX" --email-limit 2000
```

The agent now works incrementally. After a run, it saves the last checked email UID in:

```text
imap_cv_agent_state.json
```

Next time, the same command checks only newer emails:

```powershell
python .\imap_cv_agent.py --imap-server "mail.emotifzone.com" --username "careers@emotifzone.com" --folder "INBOX"
```

If you intentionally want to scan older emails again, use:

```powershell
python .\imap_cv_agent.py --imap-server "mail.emotifzone.com" --username "careers@emotifzone.com" --folder "INBOX" --email-limit 2000 --full-scan
```

Close `candidate_details.xlsx` before running the agent, otherwise Excel may block Python from saving updates.

The agent checks duplicates by:

- Candidate Name
- Email
- Contact Number

If the same candidate already exists, it updates missing blank fields instead of adding another row.

The agent skips only fully unreadable CVs. A row is saved when it finds any one of:

- Candidate Name
- Email
- Contact Number

If the name is not readable inside the CV, the agent also tries to guess the name from the attachment filename.

The Excel file is formatted automatically:

- Headings are bold
- Header row is frozen
- Filters are added
- Columns are resized
- Data is sorted by newest received email first

## If CV Emails Are In Another Folder

Example:

```powershell
python .\imap_cv_agent.py --imap-server "mail.emotifzone.com" --username "careers@emotifzone.com" --folder "CVs To Process"
```

Folder names depend on the email provider. If the folder name fails, start with:

```text
INBOX
```

## Supported CV Files

Supported:

- PDF
- DOCX

Scanned/image PDF support:

Some PDFs look readable but do not contain selectable text. For those, install Tesseract OCR:

```text
https://github.com/UB-Mannheim/tesseract/wiki
```

Install it in the default folder:

```text
C:\Program Files\Tesseract-OCR\tesseract.exe
```

Then run:

```powershell
pip install -r requirements.txt
```

The agent will automatically use OCR when normal PDF text extraction returns blank.

Not supported yet:

- Old DOC files
- ZIP files
