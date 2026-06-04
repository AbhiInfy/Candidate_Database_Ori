import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import messagebox, ttk


APP_DIR = Path(__file__).resolve().parent
PYTHON_EXE = APP_DIR / ".venv" / "Scripts" / "python.exe"
AGENT_SCRIPT = APP_DIR / "imap_cv_agent.py"
IMAP_SERVER = "mail.emotifzone.com"
USERNAME = "careers@emotifzone.com"
FOLDER = "INBOX"
OUTPUT_FILE = os.environ.get("CV_AGENT_OUTPUT", str(APP_DIR / "Hiring_Automation_Candidates.xlsx"))
ATTACHMENTS_DIR = os.environ.get("CV_AGENT_ATTACHMENTS_DIR", str(APP_DIR / "downloaded_cvs"))


VISIBLE_PREFIXES = (
    "Starting",
    "Incremental",
    "Recovered",
    "Initial/full",
    "Skipped",
    "Done.",
    "Emails checked",
    "CV attachments",
    "New candidates",
    "Missing fields",
    "Excel file",
    "Incremental checkpoint",
    "Next normal run",
)

HIDDEN_PATTERNS = (
    "FontBBox",
    "Data-loss while decompressing",
    "CropBox missing",
    "Syntax Warning",
    "cannot be parsed",
    "Traceback",
)


class CVAgentApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Motifzone Private Limited")
        self.geometry("760x520")
        self.minsize(680, 460)
        self.configure(bg="#f4f7fb")
        self.output_queue = queue.Queue()
        self.process = None

        self._build_ui()
        self.after(150, self._drain_queue)

    def _build_ui(self):
        header = tk.Frame(self, bg="#12324a", height=86)
        header.pack(fill="x")
        header.pack_propagate(False)

        tk.Label(
            header,
            text="Hiring Automation Tool",
            bg="#12324a",
            fg="white",
            font=("Segoe UI", 22, "bold"),
        ).pack(anchor="w", padx=28, pady=(16, 0))

        tk.Label(
            header,
            text="Automatically reads resume emails and updates the candidate Excel tracker",
            bg="#12324a",
            fg="#cde5f6",
            font=("Segoe UI", 10),
        ).pack(anchor="w", padx=30)

        body = tk.Frame(self, bg="#f4f7fb")
        body.pack(fill="both", expand=True, padx=28, pady=20)

        form = tk.Frame(body, bg="#f4f7fb")
        form.pack(fill="x")

        tk.Label(form, text="Email password / app password", bg="#f4f7fb", fg="#253746", font=("Segoe UI", 10, "bold")).grid(row=0, column=0, sticky="w")
        self.password_var = tk.StringVar()
        password_entry = ttk.Entry(form, textvariable=self.password_var, show="*", width=42)
        password_entry.grid(row=1, column=0, sticky="we", pady=(6, 0))
        form.columnconfigure(0, weight=1)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(form, textvariable=self.status_var, bg="#f4f7fb", fg="#526475", font=("Segoe UI", 10)).grid(row=1, column=1, sticky="e", padx=(18, 0))

        buttons = tk.Frame(body, bg="#f4f7fb")
        buttons.pack(fill="x", pady=18)

        self.run_button = ttk.Button(buttons, text="Run New Emails", command=lambda: self._start_agent([]))
        self.run_button.pack(side="left")

        self.full_button = ttk.Button(buttons, text="Full Scan", command=lambda: self._start_agent(["--email-limit", "2000", "--full-scan"]))
        self.full_button.pack(side="left", padx=10)

        self.mark_button = ttk.Button(buttons, text="Start From Now", command=lambda: self._start_agent(["--mark-current"]))
        self.mark_button.pack(side="left")

        self.progress = ttk.Progressbar(body, mode="indeterminate")
        self.progress.pack(fill="x", pady=(0, 14))

        log_frame = tk.Frame(body, bg="#ffffff", bd=1, relief="solid")
        log_frame.pack(fill="both", expand=True)

        self.log = tk.Text(log_frame, bg="#ffffff", fg="#203040", font=("Consolas", 10), relief="flat", wrap="word")
        self.log.pack(side="left", fill="both", expand=True, padx=12, pady=12)

        scrollbar = ttk.Scrollbar(log_frame, command=self.log.yview)
        scrollbar.pack(side="right", fill="y")
        self.log.configure(yscrollcommand=scrollbar.set)
        self._write("Ready. Close Hiring_Automation_Candidates.xlsx before running.\n")

    def _set_running(self, running: bool):
        state = "disabled" if running else "normal"
        self.run_button.configure(state=state)
        self.full_button.configure(state=state)
        self.mark_button.configure(state=state)
        if running:
            self.progress.start(12)
            self.status_var.set("Running...")
        else:
            self.progress.stop()
            self.status_var.set("Ready")

    def _start_agent(self, extra_args):
        password = self.password_var.get().strip()
        if not password:
            messagebox.showwarning("Password Required", "Please enter the email password or app password.")
            return
        if not PYTHON_EXE.exists():
            messagebox.showerror("Setup Missing", "Python virtual environment was not found. Please run setup first.")
            return
        if not AGENT_SCRIPT.exists():
            messagebox.showerror("Agent Missing", "imap_cv_agent.py was not found.")
            return

        self.log.delete("1.0", "end")
        self._write("Starting Hiring Automation Tool...\n")
        self._set_running(True)

        args = [
            str(PYTHON_EXE),
            str(AGENT_SCRIPT),
            "--imap-server",
            IMAP_SERVER,
            "--username",
            USERNAME,
            "--folder",
            FOLDER,
            "--output",
            OUTPUT_FILE,
            "--attachments-dir",
            ATTACHMENTS_DIR,
            *extra_args,
        ]
        env = os.environ.copy()
        env["CV_AGENT_PASSWORD"] = password

        thread = threading.Thread(target=self._run_process, args=(args, env), daemon=True)
        thread.start()

    def _run_process(self, args, env):
        try:
            self.process = subprocess.Popen(
                args,
                cwd=str(APP_DIR),
                env=env,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                creationflags=subprocess.CREATE_NO_WINDOW if sys.platform.startswith("win") else 0,
            )
            for line in self.process.stdout:
                self.output_queue.put(("line", line))
            code = self.process.wait()
            self.output_queue.put(("done", code))
        except Exception as exc:
            self.output_queue.put(("error", str(exc)))

    def _drain_queue(self):
        try:
            while True:
                kind, payload = self.output_queue.get_nowait()
                if kind == "line":
                    self._handle_output_line(payload)
                elif kind == "done":
                    self._set_running(False)
                    if payload == 0:
                        self._write("\nFinished successfully.\n")
                    else:
                        self._write("\nFinished with an issue. Please contact support.\n")
                elif kind == "error":
                    self._set_running(False)
                    self._write(f"\nUnable to start agent: {payload}\n")
        except queue.Empty:
            pass
        self.after(150, self._drain_queue)

    def _handle_output_line(self, line):
        clean = line.strip()
        if not clean:
            return
        if any(pattern in clean for pattern in HIDDEN_PATTERNS):
            return
        if clean.startswith(VISIBLE_PREFIXES):
            self._write(clean + "\n")

    def _write(self, text):
        self.log.insert("end", text)
        self.log.see("end")


if __name__ == "__main__":
    app = CVAgentApp()
    app.mainloop()
