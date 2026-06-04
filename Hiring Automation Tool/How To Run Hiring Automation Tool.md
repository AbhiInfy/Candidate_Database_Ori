# Run CV Agent From Icon

Use these files to run the CV agent without opening VS Code:

```text
Run_CV_Agent_UI.bat
Run_CV_Agent.bat
Run_CV_Agent_Full_Scan.bat
Set_CV_Agent_Checkpoint_Now.bat
```

## Recommended Client Run

Double-click:

```text
Run_CV_Agent_UI.bat
```

This opens a simple CV Agent window instead of showing raw command output.

The user can click:

- `Run New Emails` for daily incremental checking
- `Full Scan` only when old emails must be scanned again
- `Start From Now` to create a checkpoint and ignore old mailbox emails

No Python code is shown to the user.

## Normal Daily Run From Command Window

Double-click:

```text
Run_CV_Agent.bat
```

This checks only new emails after the previous run.

It should not scan 500 old emails after the checkpoint is created.

## Start Incremental From Now

If you already processed the old emails and want the agent to ignore all existing mailbox emails, double-click:

```text
Set_CV_Agent_Checkpoint_Now.bat
```

This does not read CVs or update Excel. It only saves the current latest email UID.

After that, use only:

```text
Run_CV_Agent.bat
```

and it will check only emails received after the checkpoint.

## Full Rescan

Double-click:

```text
Run_CV_Agent_Full_Scan.bat
```

This scans older emails again and is slower.

## Create Desktop Icon

1. Right-click `Run_CV_Agent_UI.bat`
2. Click `Show more options`
3. Click `Create shortcut`
4. Move the shortcut to Desktop
5. Rename it to:

```text
CV Agent
```

Now the user can double-click the Desktop icon.

## Important

Keep this Excel file closed before running:

```text
candidate_details.xlsx
```

In the GUI version, the password is typed into the password box.

In the command-window version, the password is typed into the black command window and is not visible while typing.

## Can The Python Code Be Completely Hidden?

The icon method hides the code from normal usage, but the `.py` file still exists in the project folder.

To fully hide the Python code, create a Windows `.exe` using PyInstaller. That can be added later.
