# Outlook Cloud CV Agent Steps

Use this version when you do not want to configure classic Outlook 2016 on the computer.

This agent uses Microsoft Graph to read Outlook mail from the cloud.

## What It Creates

Excel file:

```text
candidate_details.xlsx
```

Sheet name:

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

## Step 1: Create Outlook Mail Folder

In Outlook Web or New Outlook, create a folder:

```text
CVs To Process
```

Move CV emails into this folder.

## Step 2: Create Microsoft App Registration

Open:

https://portal.azure.com/

Search:

```text
App registrations
```

Create a new registration:

- Name: `CV Extractor Agent`
- Supported account types: choose the option that matches your account. If unsure, choose accounts in any organizational directory and personal Microsoft accounts.
- Redirect URI: leave blank for now.

After creating it, copy:

```text
Application (client) ID
```

You will use it in the command as `--client-id`.

## Step 3: Enable Public Client Flow

In the app registration, open:

```text
Authentication
```

Enable:

```text
Allow public client flows
```

Save.

## Step 4: Add Mail Permission

Open:

```text
API permissions
```

Add permission:

```text
Microsoft Graph
Delegated permissions
Mail.Read
```

Save.

If your company account requires admin approval, ask IT admin to grant consent for `Mail.Read`.

## Step 5: Install Python Packages

In VS Code terminal:

```powershell
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

## Step 6: Run The Agent

Replace `YOUR_CLIENT_ID_HERE` with the Application client ID.

```powershell
python .\outlook_graph_cv_agent.py --client-id "YOUR_CLIENT_ID_HERE"
```

The first time, the terminal will show a Microsoft login code and URL.

Open the URL, enter the code, and approve access.

After that, the token is saved locally in:

```text
ms_graph_token_cache.json
```

Next runs should not ask you to approve again unless the token expires or permissions change.

## If Your Folder Is Under Inbox

Try:

```powershell
python .\outlook_graph_cv_agent.py --client-id "YOUR_CLIENT_ID_HERE" --folder "Inbox/CVs To Process"
```

## Supported CV Files

Supported:

- PDF
- DOCX

Not supported yet:

- Old DOC files
- Scanned image PDFs
- ZIP files
