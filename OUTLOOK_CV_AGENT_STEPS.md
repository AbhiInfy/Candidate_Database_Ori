# Outlook CV Agent Steps

This agent reads CV attachments from Outlook Desktop and creates an Excel file with candidate details.

## What It Creates

Excel file:

```text
candidate_details.xlsx
```

Excel sheet:

```text
Candidates
```

Columns:

- Candidate Name
- Email
- Contact Number
- Experience
- Current Company
- Skills
- CV File
- Email Subject
- Sender
- Received Time
- Processed At

## Step 1: Create Outlook Folder

Open Outlook Desktop.

Create a folder under Inbox named:

```text
CVs To Process
```

Move only the emails with CV attachments into this folder.

This is safer than scanning your full Inbox.

## Step 2: Install Packages

In VS Code terminal, activate your environment:

```powershell
.\.venv\Scripts\Activate.ps1
```

Then install requirements:

```powershell
pip install -r requirements.txt
```

## Step 3: Run The Agent

```powershell
python .\outlook_cv_agent.py
```

It will read:

```text
Inbox/CVs To Process
```

and create:

```text
candidate_details.xlsx
```

## Step 4: If Your Outlook Folder Name Is Different

Example:

```powershell
python .\outlook_cv_agent.py --folder "Inbox/Oracle CVs"
```

## Supported CV Files

Supported now:

- PDF
- DOCX

Not supported in this first version:

- Old DOC files
- Scanned image PDFs
- ZIP files

Those can be added later if needed.
