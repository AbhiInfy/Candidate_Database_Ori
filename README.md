# Naukri Job Link Agent

This Python agent opens Naukri in a real browser, searches for a technology, collects job post links from the last 24 hours, and stores only new links in an Excel sheet.

## 1. Install Python

Install Python 3.11 or newer from:

https://www.python.org/downloads/

During installation, select **Add python.exe to PATH**.

To check Python is installed, open PowerShell and run:

```powershell
python --version
```

## 2. Open This Project Folder

In PowerShell:

```powershell
cd "C:\Users\Nitin Maheshwari\Documents\Codex\2026-05-13\i-want-to-make-an-ai"
```

## 3. Create a Virtual Environment

```powershell
python -m venv .venv
```

Activate it:

```powershell
.\.venv\Scripts\Activate.ps1
```

If PowerShell blocks activation, run this once:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

Then activate again:

```powershell
.\.venv\Scripts\Activate.ps1
```

## 4. Install Required Packages

```powershell
pip install -r requirements.txt
```

Then install the browser used by the agent:

```powershell
python -m playwright install chromium
```

## 5. First Run With Manual Login

Run:

```powershell
python .\naukri_agent.py --login --keyword "Oracle Fusion Application" --pages 3
```

A browser window will open.

1. Log in to Naukri manually.
2. After login is complete, return to PowerShell.
3. Press Enter.
4. The agent will search Naukri and save links to Excel.

Your login is saved inside the local folder `naukri_browser_profile`, so you do not need to log in every time.

## 6. Normal Run After Login Is Saved

```powershell
python .\naukri_agent.py --keyword "Oracle Fusion Application" --pages 3
```

The output Excel file will be:

```text
naukri_job_links.xlsx
```

## 7. Change Technology

Example:

```powershell
python .\naukri_agent.py --keyword "Oracle Fusion HCM" --pages 5
```

Another example:

```powershell
python .\naukri_agent.py --keyword "Oracle Fusion Financials" --pages 5
```

## 8. Excel Columns

The agent writes these columns:

- Technology
- Job Title
- Company
- Posted
- Email
- Contact Number
- Job Link
- Collected At

Before adding a row, it checks whether the same job link already exists. If the link is already present, it skips it.

The agent first checks the visible search result card for email and contact number. If it does not find them there, it opens the job detail page and checks the visible text on that page.

If you want a faster run and do not need email/contact details, use:

```powershell
python .\naukri_agent.py --keyword "Oracle Fusion Application" --pages 3 --skip-contact-details
```

## 9. Common Problems

If `python` is not recognized, reinstall Python and select **Add python.exe to PATH**.

If activation is blocked, run:

```powershell
Set-ExecutionPolicy -Scope CurrentUser RemoteSigned
```

If Naukri asks you to log in again, run with `--login`:

```powershell
python .\naukri_agent.py --login --keyword "Oracle Fusion Application" --pages 3
```
