import os
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk


DEFAULT_MODEL = "nvidia/nemotron-nano-12b-v2-vl:free"


class OcrRunnerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Caseware OCR Runner")
        self.geometry("860x640")
        self.resizable(True, True)

        self.process = None

        self._build_ui()

    def _build_ui(self):
        padx = 10
        pady = 6

        container = ttk.Frame(self)
        container.pack(fill="both", expand=True, padx=padx, pady=padx)

        # API Key
        ttk.Label(container, text="OpenRouter API Key").grid(row=0, column=0, sticky="w")
        self.api_key_var = tk.StringVar()
        self.api_entry = ttk.Entry(container, textvariable=self.api_key_var, show="*")
        self.api_entry.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(0, pady))

        self.show_key_var = tk.BooleanVar(value=False)
        show_key = ttk.Checkbutton(
            container,
            text="Show key",
            variable=self.show_key_var,
            command=self._toggle_key_visibility,
        )
        show_key.grid(row=1, column=2, sticky="w", padx=(10, 0))

        # Model
        ttk.Label(container, text="Model").grid(row=2, column=0, sticky="w")
        self.model_var = tk.StringVar(value=DEFAULT_MODEL)
        self.model_combo = ttk.Combobox(
            container,
            textvariable=self.model_var,
            values=[DEFAULT_MODEL],
            state="normal",
        )
        self.model_combo.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(0, pady))

        # Folder
        ttk.Label(container, text="Input folder").grid(row=4, column=0, sticky="w")
        self.path_var = tk.StringVar()
        path_entry = ttk.Entry(container, textvariable=self.path_var)
        path_entry.grid(row=5, column=0, columnspan=2, sticky="ew", pady=(0, pady))
        browse_btn = ttk.Button(container, text="Browse", command=self._browse_folder)
        browse_btn.grid(row=5, column=2, sticky="ew", padx=(10, 0))

        # Options
        options_frame = ttk.LabelFrame(container, text="Options")
        options_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(4, pady))
        options_frame.columnconfigure(1, weight=1)
        options_frame.columnconfigure(3, weight=1)

        ttk.Label(options_frame, text="Max pages").grid(row=0, column=0, sticky="w", padx=8, pady=4)
        self.max_pages_var = tk.StringVar(value="12")
        ttk.Entry(options_frame, textvariable=self.max_pages_var, width=8).grid(row=0, column=1, sticky="w")

        ttk.Label(options_frame, text="RPM").grid(row=0, column=2, sticky="w", padx=8)
        self.rpm_var = tk.StringVar(value="10")
        ttk.Entry(options_frame, textvariable=self.rpm_var, width=8).grid(row=0, column=3, sticky="w")

        self.refine_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(options_frame, text="Refine amounts", variable=self.refine_var).grid(
            row=1, column=0, sticky="w", padx=8, pady=4
        )

        self.recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(options_frame, text="Recursive", variable=self.recursive_var).grid(
            row=1, column=1, sticky="w", padx=8
        )

        # Run controls
        controls_frame = ttk.Frame(container)
        controls_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(4, pady))
        controls_frame.columnconfigure(0, weight=1)
        controls_frame.columnconfigure(1, weight=1)

        self.run_btn = ttk.Button(controls_frame, text="Run OCR", command=self._run_ocr)
        self.run_btn.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        self.stop_btn = ttk.Button(controls_frame, text="Stop", command=self._stop_ocr, state="disabled")
        self.stop_btn.grid(row=0, column=1, sticky="ew", padx=(6, 0))

        # Log
        ttk.Label(container, text="Log").grid(row=8, column=0, sticky="w")
        self.log = scrolledtext.ScrolledText(container, height=18, wrap="word", state="disabled")
        self.log.grid(row=9, column=0, columnspan=3, sticky="nsew")

        container.columnconfigure(0, weight=1)
        container.columnconfigure(1, weight=1)
        container.columnconfigure(2, weight=0)
        container.rowconfigure(9, weight=1)

    def _toggle_key_visibility(self):
        self.api_entry.configure(show="" if self.show_key_var.get() else "*")

    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.path_var.set(folder)

    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _run_ocr(self):
        api_key = self.api_key_var.get().strip()
        model = self.model_var.get().strip()
        input_path = self.path_var.get().strip()

        if not api_key:
            messagebox.showerror("Missing key", "Please paste your OpenRouter API key.")
            return
        if not model:
            messagebox.showerror("Missing model", "Please enter a model name.")
            return
        if not input_path or not os.path.isdir(input_path):
            messagebox.showerror("Invalid folder", "Please select a valid input folder.")
            return

        try:
            max_pages = int(self.max_pages_var.get().strip())
            rpm = int(self.rpm_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid options", "Max pages and RPM must be integers.")
            return

        script_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "ocr_tool.py")
        if not os.path.isfile(script_path):
            messagebox.showerror("Missing script", f"Could not find ocr_tool.py at {script_path}")
            return

        cmd = [
            sys.executable,
            script_path,
            "--input",
            input_path,
            "--model",
            model,
            "--max-pages",
            str(max_pages),
            "--rpm",
            str(rpm),
        ]

        if not self.recursive_var.get():
            cmd.append("--no-recursive")
        if not self.refine_var.get():
            cmd.append("--no-refine-amounts")

        env = os.environ.copy()
        env["OPENROUTER_API_KEY"] = api_key

        self.run_btn.configure(state="disabled")
        self.stop_btn.configure(state="normal")
        self._append_log("Starting OCR...")

        def worker():
            try:
                self.process = subprocess.Popen(
                    cmd,
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    text=True,
                    env=env,
                    bufsize=1,
                )
                for line in self.process.stdout:
                    self.after(0, self._append_log, line.rstrip())
                self.process.wait()
                code = self.process.returncode
                self.after(0, self._append_log, f"Done. Exit code: {code}")
            except Exception as exc:
                self.after(0, self._append_log, f"ERROR: {exc}")
            finally:
                self.process = None
                self.after(0, lambda: self.run_btn.configure(state="normal"))
                self.after(0, lambda: self.stop_btn.configure(state="disabled"))

        threading.Thread(target=worker, daemon=True).start()

    def _stop_ocr(self):
        if self.process and self.process.poll() is None:
            self.process.terminate()
            self._append_log("Stopped by user.")
        self.run_btn.configure(state="normal")
        self.stop_btn.configure(state="disabled")


if __name__ == "__main__":
    app = OcrRunnerApp()
    app.mainloop()
