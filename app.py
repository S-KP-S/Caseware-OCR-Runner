import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, simpledialog, ttk

from config import (
    DEFAULT_PROFILE,
    get_profile,
    load_accounts_map,
    load_config,
    load_vendor_map,
    save_accounts_map,
    save_config,
    save_profile,
    save_vendor_map,
)
from ocr_tool import (
    DUPLICATE_FLAG_PREFIX,
    EXPORT_CASEWARE,
    EXPORT_FORMATS,
    EXPORT_FULL,
    FULL_FIELDNAMES,
    LOW_CONFIDENCE_THRESHOLD,
    process_directory,
    set_progress_callback,
    write_csv,
    write_summary,
)


BG = "#0b1220"
PANEL_BG = "#121a2a"
FG = "#e8edf5"
MUTED = "#8da2bf"
ACCENT = "#2f7de1"
RUN_GREEN = "#1f9d55"
STOP_RED = "#d64545"
ENTRY_BG = "#0f1726"
YELLOW = "#f0b429"
RED = "#e05d5d"

MODEL_OPTIONS = [
    "nvidia/nemotron-nano-12b-v2-vl:free",
    "qwen/qwen2.5-vl-72b-instruct:free",
    "google/gemini-2.0-flash-exp:free",
]


class OcrRunnerApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Caseware OCR Runner")
        self.geometry("1280x820")
        self.minsize(1100, 700)

        self.cfg = load_config()
        self.vendor_map = load_vendor_map()
        self.accounts_map = load_accounts_map()

        self._rows = []
        self._filtered_indices = []
        self._sort_column = None
        self._sort_reverse = False
        self._stop_requested = False
        self._worker = None
        self._last_csv_path = ""
        self._last_summary_path = ""
        self._editor = None

        self._setup_style()
        self._build_ui()
        self._load_profile(self.cfg.get("active_profile", "default"))
        self._populate_mappings()

    # SECTION: theme and layout
    def _setup_style(self):
        self.configure(bg=BG)
        style = ttk.Style(self)
        style.theme_use("clam")

        style.configure("Root.TFrame", background=BG)
        style.configure("Panel.TFrame", background=PANEL_BG)
        style.configure("TLabel", background=PANEL_BG, foreground=FG)
        style.configure("Muted.TLabel", background=PANEL_BG, foreground=MUTED)
        style.configure("TNotebook", background=BG, borderwidth=0)
        style.configure("TNotebook.Tab", background="#182239", foreground=FG, padding=(14, 8))
        style.map("TNotebook.Tab", background=[("selected", ACCENT)])
        style.configure("TEntry", fieldbackground=ENTRY_BG, foreground=FG)
        style.configure("TCombobox", fieldbackground=ENTRY_BG, foreground=FG)
        style.configure("Dark.Treeview", background=ENTRY_BG, fieldbackground=ENTRY_BG, foreground=FG, rowheight=24)
        style.configure("Dark.Treeview.Heading", background="#1f2a44", foreground=FG)
        style.configure("Run.TButton", background=RUN_GREEN, foreground="white", padding=8)
        style.map("Run.TButton", background=[("active", "#168447")])
        style.configure("Stop.TButton", background=STOP_RED, foreground="white", padding=8)
        style.map("Stop.TButton", background=[("active", "#c33a3a")])
        style.configure("Accent.TButton", background=ACCENT, foreground="white")

    def _build_ui(self):
        root = ttk.Frame(self, style="Root.TFrame")
        root.pack(fill="both", expand=True, padx=10, pady=10)

        self.tabs = ttk.Notebook(root)
        self.tabs.pack(fill="both", expand=True)

        self.tab_ocr = ttk.Frame(self.tabs, style="Panel.TFrame")
        self.tab_review = ttk.Frame(self.tabs, style="Panel.TFrame")
        self.tab_mappings = ttk.Frame(self.tabs, style="Panel.TFrame")
        self.tabs.add(self.tab_ocr, text="OCR")
        self.tabs.add(self.tab_review, text="Review")
        self.tabs.add(self.tab_mappings, text="Mappings")

        self._build_tab_ocr()
        self._build_tab_review()
        self._build_tab_mappings()

    def _build_tab_ocr(self):
        frame = self.tab_ocr
        for col in range(6):
            frame.columnconfigure(col, weight=1 if col in (1, 3, 5) else 0)
        frame.rowconfigure(8, weight=1)

        row = 0
        ttk.Label(frame, text="Profile").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(frame, textvariable=self.profile_var, state="readonly")
        self.profile_combo.grid(row=row, column=1, sticky="ew", padx=8, pady=6)
        self.profile_combo.bind("<<ComboboxSelected>>", lambda _e: self._load_profile(self.profile_var.get()))
        ttk.Button(frame, text="Save Profile", command=self._save_current_profile).grid(row=row, column=2, sticky="ew", padx=8, pady=6)
        ttk.Button(frame, text="New Profile", command=self._new_profile).grid(row=row, column=3, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(frame, text="OpenRouter API Key").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.api_key_var = tk.StringVar()
        self.api_entry = ttk.Entry(frame, textvariable=self.api_key_var, show="*")
        self.api_entry.grid(row=row, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
        self.show_key_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Show", variable=self.show_key_var, command=self._toggle_key_visibility).grid(
            row=row, column=5, sticky="w", padx=8, pady=6
        )

        row += 1
        ttk.Label(frame, text="Model").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.model_var = tk.StringVar()
        self.model_combo = ttk.Combobox(frame, textvariable=self.model_var, values=MODEL_OPTIONS, state="normal")
        self.model_combo.grid(row=row, column=1, columnspan=5, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(frame, text="Input Folder").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.path_var).grid(row=row, column=1, columnspan=4, sticky="ew", padx=8, pady=6)
        ttk.Button(frame, text="Browse", command=self._browse_folder).grid(row=row, column=5, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(frame, text="Max Pages").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.max_pages_var = tk.StringVar(value="12")
        ttk.Entry(frame, textvariable=self.max_pages_var, width=8).grid(row=row, column=1, sticky="w", padx=8, pady=6)
        ttk.Label(frame, text="RPM").grid(row=row, column=2, sticky="w", padx=8, pady=6)
        self.rpm_var = tk.StringVar(value="12")
        ttk.Entry(frame, textvariable=self.rpm_var, width=8).grid(row=row, column=3, sticky="w", padx=8, pady=6)
        ttk.Label(frame, text="Currency").grid(row=row, column=4, sticky="w", padx=8, pady=6)
        self.currency_var = tk.StringVar(value="CAD")
        ttk.Combobox(frame, textvariable=self.currency_var, values=["CAD", "USD", "GBP", "EUR"], state="readonly").grid(
            row=row, column=5, sticky="ew", padx=8, pady=6
        )

        row += 1
        ttk.Label(frame, text="Export Format").grid(row=row, column=0, sticky="w", padx=8, pady=6)
        self.export_format_var = tk.StringVar(value=EXPORT_CASEWARE)
        ttk.Combobox(frame, textvariable=self.export_format_var, values=sorted(EXPORT_FORMATS), state="readonly").grid(
            row=row, column=1, sticky="ew", padx=8, pady=6
        )
        self.refine_var = tk.BooleanVar(value=True)
        self.recursive_var = tk.BooleanVar(value=True)
        self.use_cache_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Refine Amounts", variable=self.refine_var).grid(row=row, column=2, sticky="w", padx=8, pady=6)
        ttk.Checkbutton(frame, text="Recursive", variable=self.recursive_var).grid(row=row, column=3, sticky="w", padx=8, pady=6)
        ttk.Checkbutton(frame, text="Use Cache", variable=self.use_cache_var).grid(row=row, column=4, sticky="w", padx=8, pady=6)

        row += 1
        self.progress_label = ttk.Label(frame, text="Idle", style="Muted.TLabel")
        self.progress_label.grid(row=row, column=0, columnspan=4, sticky="w", padx=8, pady=6)
        self.progress_var = tk.DoubleVar(value=0)
        self.progress = ttk.Progressbar(frame, variable=self.progress_var, maximum=100)
        self.progress.grid(row=row, column=4, columnspan=2, sticky="ew", padx=8, pady=6)

        row += 1
        self.run_btn = ttk.Button(frame, text="Run OCR", style="Run.TButton", command=self._run_ocr)
        self.run_btn.grid(row=row, column=0, columnspan=3, sticky="ew", padx=8, pady=6)
        self.stop_btn = ttk.Button(frame, text="Stop", style="Stop.TButton", command=self._stop_ocr, state="disabled")
        self.stop_btn.grid(row=row, column=3, columnspan=3, sticky="ew", padx=8, pady=6)

        row += 1
        ttk.Label(frame, text="Log").grid(row=row, column=0, sticky="w", padx=8, pady=(8, 4))
        row += 1
        self.log = scrolledtext.ScrolledText(
            frame,
            wrap="word",
            font=("Consolas", 10),
            bg=ENTRY_BG,
            fg=FG,
            insertbackground=FG,
            relief="flat",
        )
        self.log.grid(row=row, column=0, columnspan=6, sticky="nsew", padx=8, pady=(0, 8))
        self.log.configure(state="disabled")

    # SECTION: review tab
    def _build_tab_review(self):
        frame = self.tab_review
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(1, weight=1)

        toolbar = ttk.Frame(frame, style="Panel.TFrame")
        toolbar.grid(row=0, column=0, sticky="ew", padx=8, pady=8)
        for i in range(8):
            toolbar.columnconfigure(i, weight=0)
        toolbar.columnconfigure(7, weight=1)

        ttk.Button(toolbar, text="Export CSV", command=self._export_csv_from_review).grid(row=0, column=0, padx=4)
        ttk.Button(toolbar, text="Export Summary", command=self._export_summary_from_review).grid(row=0, column=1, padx=4)
        ttk.Button(toolbar, text="Delete Selected", command=self._delete_selected_rows).grid(row=0, column=2, padx=4)

        ttk.Label(toolbar, text="Filter").grid(row=0, column=3, padx=(12, 4))
        self.filter_var = tk.StringVar(value="All")
        self.filter_combo = ttk.Combobox(
            toolbar,
            textvariable=self.filter_var,
            values=["All", "Low Confidence", "Flagged", "Duplicates"],
            state="readonly",
            width=18,
        )
        self.filter_combo.grid(row=0, column=4, padx=4)
        self.filter_combo.bind("<<ComboboxSelected>>", lambda _e: self._refresh_review_table())

        self.count_label = ttk.Label(toolbar, text="0 transactions", style="Muted.TLabel")
        self.count_label.grid(row=0, column=7, sticky="e", padx=4)

        table_wrap = ttk.Frame(frame, style="Panel.TFrame")
        table_wrap.grid(row=1, column=0, sticky="nsew", padx=8, pady=(0, 8))
        table_wrap.columnconfigure(0, weight=1)
        table_wrap.rowconfigure(0, weight=1)

        self.review_tree = ttk.Treeview(
            table_wrap,
            columns=FULL_FIELDNAMES,
            show="headings",
            style="Dark.Treeview",
            selectmode="extended",
        )
        for col in FULL_FIELDNAMES:
            self.review_tree.heading(col, text=col, command=lambda c=col: self._sort_by_column(c))
            self.review_tree.column(col, width=140, anchor="w")
        self.review_tree.grid(row=0, column=0, sticky="nsew")

        yscroll = ttk.Scrollbar(table_wrap, orient="vertical", command=self.review_tree.yview)
        yscroll.grid(row=0, column=1, sticky="ns")
        self.review_tree.configure(yscrollcommand=yscroll.set)

        self.review_tree.tag_configure("low70", background="#3a331c", foreground=FG)
        self.review_tree.tag_configure("low50", background="#442326", foreground=FG)
        self.review_tree.bind("<Double-1>", self._start_cell_edit)

    # SECTION: mappings tab
    def _build_tab_mappings(self):
        frame = self.tab_mappings
        frame.columnconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)
        frame.rowconfigure(0, weight=1)

        vendor_panel = ttk.LabelFrame(frame, text="Vendor Normalization")
        vendor_panel.grid(row=0, column=0, sticky="nsew", padx=8, pady=8)
        vendor_panel.columnconfigure(0, weight=1)
        vendor_panel.rowconfigure(0, weight=1)

        self.vendor_tree = ttk.Treeview(vendor_panel, columns=["Pattern", "Vendor"], show="headings", style="Dark.Treeview")
        self.vendor_tree.heading("Pattern", text="Pattern")
        self.vendor_tree.heading("Vendor", text="Vendor")
        self.vendor_tree.column("Pattern", width=220)
        self.vendor_tree.column("Vendor", width=220)
        self.vendor_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Button(vendor_panel, text="Add", command=self._add_vendor_mapping).grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Button(vendor_panel, text="Delete", command=self._delete_vendor_mapping).grid(row=1, column=0, padx=70, pady=6, sticky="w")
        ttk.Button(vendor_panel, text="Save", style="Accent.TButton", command=self._save_vendor_mappings).grid(
            row=1, column=0, sticky="e", padx=6, pady=6
        )

        account_panel = ttk.LabelFrame(frame, text="Chart of Accounts")
        account_panel.grid(row=0, column=1, sticky="nsew", padx=8, pady=8)
        account_panel.columnconfigure(0, weight=1)
        account_panel.rowconfigure(0, weight=1)

        self.account_tree = ttk.Treeview(account_panel, columns=["Vendor", "Account"], show="headings", style="Dark.Treeview")
        self.account_tree.heading("Vendor", text="Vendor")
        self.account_tree.heading("Account", text="Account")
        self.account_tree.column("Vendor", width=220)
        self.account_tree.column("Account", width=220)
        self.account_tree.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Button(account_panel, text="Add", command=self._add_account_mapping).grid(row=1, column=0, sticky="w", padx=6, pady=6)
        ttk.Button(account_panel, text="Delete", command=self._delete_account_mapping).grid(row=1, column=0, padx=70, pady=6, sticky="w")
        ttk.Button(account_panel, text="Save", style="Accent.TButton", command=self._save_account_mappings).grid(
            row=1, column=0, sticky="e", padx=6, pady=6
        )

    # SECTION: profiles and settings
    def _toggle_key_visibility(self):
        self.api_entry.configure(show="" if self.show_key_var.get() else "*")

    def _browse_folder(self):
        folder = filedialog.askdirectory()
        if folder:
            self.path_var.set(folder)

    def _profile_names(self):
        profiles = self.cfg.get("profiles", {})
        names = sorted(profiles.keys())
        if not names:
            names = ["default"]
        return names

    def _refresh_profile_combo(self):
        names = self._profile_names()
        self.profile_combo["values"] = names
        if self.profile_var.get() not in names:
            self.profile_var.set(names[0])

    def _collect_profile_data(self):
        return {
            "api_key": self.api_key_var.get().strip(),
            "model": self.model_var.get().strip(),
            "currency": self.currency_var.get().strip() or "CAD",
            "max_pages": int(self.max_pages_var.get().strip()),
            "rpm": int(self.rpm_var.get().strip()),
            "zoom": float(DEFAULT_PROFILE.get("zoom", 2.0)),
            "max_retries": int(DEFAULT_PROFILE.get("max_retries", 5)),
            "retry_backoff": int(DEFAULT_PROFILE.get("retry_backoff", 5)),
            "retry_max_sleep": int(DEFAULT_PROFILE.get("retry_max_sleep", 60)),
            "refine_amounts": bool(self.refine_var.get()),
            "recursive": bool(self.recursive_var.get()),
            "use_cache": bool(self.use_cache_var.get()),
            "export_format": self.export_format_var.get().strip() or EXPORT_CASEWARE,
        }

    def _load_profile(self, name):
        self.cfg = load_config()
        profile = get_profile(self.cfg, name)
        self._refresh_profile_combo()
        self.profile_var.set(name if name in self._profile_names() else self._profile_names()[0])
        self.api_key_var.set(profile.get("api_key", ""))
        self.model_var.set(profile.get("model", DEFAULT_PROFILE.get("model", "")))
        self.currency_var.set(profile.get("currency", "CAD"))
        self.max_pages_var.set(str(profile.get("max_pages", 12)))
        self.rpm_var.set(str(profile.get("rpm", 12)))
        self.refine_var.set(bool(profile.get("refine_amounts", True)))
        self.recursive_var.set(bool(profile.get("recursive", True)))
        self.use_cache_var.set(bool(profile.get("use_cache", True)))
        self.export_format_var.set(profile.get("export_format", EXPORT_CASEWARE))

    def _save_current_profile(self):
        name = self.profile_var.get().strip() or "default"
        try:
            data = self._collect_profile_data()
        except ValueError:
            messagebox.showerror("Invalid profile values", "Max Pages and RPM must be numeric.")
            return
        self.cfg = save_profile(load_config(), name, data)
        self.cfg["active_profile"] = name
        self.cfg = save_config(self.cfg)
        self._refresh_profile_combo()
        self.profile_var.set(name)
        self._append_log(f"Saved profile: {name}")

    def _new_profile(self):
        name = simpledialog.askstring("New Profile", "Profile name:")
        if not name:
            return
        name = name.strip()
        if not name:
            return
        self.cfg = save_profile(load_config(), name, DEFAULT_PROFILE.copy())
        self.cfg["active_profile"] = name
        self.cfg = save_config(self.cfg)
        self._refresh_profile_combo()
        self._load_profile(name)

    # SECTION: run workflow
    def _append_log(self, text):
        self.log.configure(state="normal")
        self.log.insert("end", safe_text(text) + "\n")
        self.log.see("end")
        self.log.configure(state="disabled")

    def _set_running(self, running):
        self.run_btn.configure(state="disabled" if running else "normal")
        self.stop_btn.configure(state="normal" if running else "disabled")

    def _progress_callback(self, event):
        self.after(0, self._handle_progress_event, event)

    def _handle_progress_event(self, event):
        message = safe_text(event.get("message", ""))
        current = event.get("current")
        total = event.get("total")
        if current is not None and total:
            pct = max(0.0, min(100.0, (float(current) / float(total)) * 100.0))
            self.progress_var.set(pct)
            self.progress_label.configure(text=f"Processing {int(current)}/{int(total)} files...")
        elif message:
            self.progress_label.configure(text=message)
        if message:
            self._append_log(message)

    def _run_ocr(self):
        api_key = self.api_key_var.get().strip()
        model = self.model_var.get().strip()
        input_path = self.path_var.get().strip()
        if not api_key:
            messagebox.showerror("Missing key", "Please enter your OpenRouter API key.")
            return
        if not model:
            messagebox.showerror("Missing model", "Please enter a model.")
            return
        if not input_path or not os.path.isdir(input_path):
            messagebox.showerror("Invalid folder", "Please select a valid input folder.")
            return
        try:
            max_pages = int(self.max_pages_var.get().strip())
            rpm = int(self.rpm_var.get().strip())
        except ValueError:
            messagebox.showerror("Invalid options", "Max Pages and RPM must be integers.")
            return

        self._save_current_profile()
        self._stop_requested = False
        self.progress_var.set(0)
        self.progress_label.configure(text="Starting...")
        self._set_running(True)

        vendor_map = self._read_vendor_tree()
        accounts_map = self._read_account_tree()

        def worker():
            set_progress_callback(self._progress_callback)
            try:
                rows, csv_path, summary_path = process_directory(
                    input_path=input_path,
                    api_key=api_key,
                    model=model,
                    max_pages=max_pages,
                    rpm=rpm,
                    currency=self.currency_var.get().strip() or "CAD",
                    recursive=self.recursive_var.get(),
                    refine_amounts=self.refine_var.get(),
                    use_cache=self.use_cache_var.get(),
                    export_format=self.export_format_var.get().strip() or EXPORT_CASEWARE,
                    vendor_map=vendor_map,
                    accounts_map=accounts_map,
                    stop_check=lambda: self._stop_requested,
                )
                self._rows = rows
                self._last_csv_path = csv_path
                self._last_summary_path = summary_path
                self.after(0, self._on_run_success)
            except Exception as exc:
                self.after(0, lambda: self._append_log(f"ERROR: {exc}"))
            finally:
                set_progress_callback(None)
                self.after(0, lambda: self._set_running(False))

        self._worker = threading.Thread(target=worker, daemon=True)
        self._worker.start()

    def _stop_ocr(self):
        self._stop_requested = True
        self._append_log("Stop requested...")

    def _on_run_success(self):
        self.progress_var.set(100)
        self.progress_label.configure(text="Complete")
        self._refresh_review_table()
        self.tabs.select(self.tab_review)
        self._append_log("OCR run complete.")

    # SECTION: review behavior
    def _passes_filter(self, row):
        mode = self.filter_var.get()
        if mode == "Low Confidence":
            try:
                return int(float(row.get("Confidence", 0))) < LOW_CONFIDENCE_THRESHOLD
            except (TypeError, ValueError):
                return True
        if mode == "Flagged":
            return bool(str(row.get("Flags", "")).strip())
        if mode == "Duplicates":
            return DUPLICATE_FLAG_PREFIX.lower() in str(row.get("Flags", "")).lower()
        return True

    def _sort_key(self, row, column):
        value = row.get(column, "")
        if value is None:
            return (1, "")
        text = str(value).strip()
        try:
            return (0, float(text.replace(",", "")))
        except ValueError:
            return (1, text.lower())

    def _refresh_review_table(self):
        tree = self.review_tree
        for item in tree.get_children():
            tree.delete(item)

        indices = [i for i, row in enumerate(self._rows) if self._passes_filter(row)]
        if self._sort_column:
            indices.sort(key=lambda idx: self._sort_key(self._rows[idx], self._sort_column), reverse=self._sort_reverse)
        self._filtered_indices = indices

        for idx in indices:
            row = self._rows[idx]
            values = [row.get(col, "") for col in FULL_FIELDNAMES]
            tag = self._confidence_tag(row)
            tree.insert("", "end", iid=str(idx), values=values, tags=(tag,) if tag else ())

        self.count_label.configure(text=f"{len(indices)} transactions")

    def _confidence_tag(self, row):
        try:
            conf = int(float(row.get("Confidence", 0)))
        except (TypeError, ValueError):
            return "low50"
        if conf < 50:
            return "low50"
        if conf < LOW_CONFIDENCE_THRESHOLD:
            return "low70"
        return ""

    def _sort_by_column(self, column):
        if self._sort_column == column:
            self._sort_reverse = not self._sort_reverse
        else:
            self._sort_column = column
            self._sort_reverse = False
        self._refresh_review_table()

    def _start_cell_edit(self, event):
        row_id = self.review_tree.identify_row(event.y)
        col_id = self.review_tree.identify_column(event.x)
        if not row_id or not col_id:
            return
        col_idx = int(col_id.replace("#", "")) - 1
        if col_idx < 0 or col_idx >= len(FULL_FIELDNAMES):
            return
        col_name = FULL_FIELDNAMES[col_idx]
        bbox = self.review_tree.bbox(row_id, col_id)
        if not bbox:
            return
        x, y, w, h = bbox
        current = self.review_tree.set(row_id, col_name)

        if self._editor is not None:
            self._editor.destroy()
            self._editor = None

        entry = tk.Entry(self.review_tree, bg=ENTRY_BG, fg=FG, insertbackground=FG, relief="flat")
        entry.insert(0, current)
        entry.place(x=x, y=y, width=w, height=h)
        entry.focus_set()
        self._editor = entry

        def commit(_event=None):
            if self._editor is None or not entry.winfo_exists():
                return
            value = entry.get()
            entry.destroy()
            self._editor = None
            self.review_tree.set(row_id, col_name, value)
            self._rows[int(row_id)][col_name] = value
            tag = self._confidence_tag(self._rows[int(row_id)])
            self.review_tree.item(row_id, tags=(tag,) if tag else ())

        def cancel(_event=None):
            if self._editor is None or not entry.winfo_exists():
                return
            entry.destroy()
            self._editor = None

        entry.bind("<Return>", commit)
        entry.bind("<Escape>", cancel)
        entry.bind("<FocusOut>", commit)

    def _delete_selected_rows(self):
        selected = self.review_tree.selection()
        if not selected:
            return
        indices = sorted({int(iid) for iid in selected}, reverse=True)
        for idx in indices:
            if 0 <= idx < len(self._rows):
                del self._rows[idx]
        self._refresh_review_table()

    def _export_csv_from_review(self):
        if not self._rows:
            messagebox.showinfo("No data", "No rows to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Export CSV",
            defaultextension=".csv",
            filetypes=[("CSV Files", "*.csv")],
            initialfile=os.path.basename(self._last_csv_path) if self._last_csv_path else "transactions.csv",
        )
        if not path:
            return
        write_csv(path, self._rows, export_format=self.export_format_var.get().strip() or EXPORT_CASEWARE)
        self._append_log(f"Exported CSV: {path}")

    def _export_summary_from_review(self):
        if not self._rows:
            messagebox.showinfo("No data", "No rows to export.")
            return
        path = filedialog.asksaveasfilename(
            title="Export Summary",
            defaultextension=".xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
            initialfile=os.path.basename(self._last_summary_path) if self._last_summary_path else "summary.xlsx",
        )
        if not path:
            return
        write_summary(path, self._rows)
        self._append_log(f"Exported summary: {path}")

    # SECTION: mappings behavior
    def _populate_mappings(self):
        for item in self.vendor_tree.get_children():
            self.vendor_tree.delete(item)
        for pattern, vendor in sorted(self.vendor_map.items()):
            self.vendor_tree.insert("", "end", values=[pattern, vendor])

        for item in self.account_tree.get_children():
            self.account_tree.delete(item)
        for vendor, account in sorted(self.accounts_map.items()):
            self.account_tree.insert("", "end", values=[vendor, account])

    def _read_vendor_tree(self):
        mappings = {}
        for iid in self.vendor_tree.get_children():
            pattern, vendor = self.vendor_tree.item(iid, "values")
            pattern = str(pattern).strip()
            vendor = str(vendor).strip()
            if pattern and vendor:
                mappings[pattern] = vendor
        return mappings

    def _read_account_tree(self):
        mappings = {}
        for iid in self.account_tree.get_children():
            vendor, account = self.account_tree.item(iid, "values")
            vendor = str(vendor).strip()
            account = str(account).strip()
            if vendor and account:
                mappings[vendor] = account
        return mappings

    def _add_vendor_mapping(self):
        pattern = simpledialog.askstring("Pattern", "Match pattern (substring):")
        if pattern is None:
            return
        vendor = simpledialog.askstring("Vendor", "Normalized vendor name:")
        if vendor is None:
            return
        pattern = pattern.strip()
        vendor = vendor.strip()
        if pattern and vendor:
            self.vendor_tree.insert("", "end", values=[pattern, vendor])

    def _delete_vendor_mapping(self):
        for iid in self.vendor_tree.selection():
            self.vendor_tree.delete(iid)

    def _save_vendor_mappings(self):
        self.vendor_map = self._read_vendor_tree()
        save_vendor_map(self.vendor_map)
        self._append_log("Saved vendor mappings.")

    def _add_account_mapping(self):
        vendor = simpledialog.askstring("Vendor", "Vendor name:")
        if vendor is None:
            return
        account = simpledialog.askstring("Account", "Account category:")
        if account is None:
            return
        vendor = vendor.strip()
        account = account.strip()
        if vendor and account:
            self.account_tree.insert("", "end", values=[vendor, account])

    def _delete_account_mapping(self):
        for iid in self.account_tree.selection():
            self.account_tree.delete(iid)

    def _save_account_mappings(self):
        self.accounts_map = self._read_account_tree()
        save_accounts_map(self.accounts_map)
        self._append_log("Saved account mappings.")


def safe_text(value):
    return str(value) if value is not None else ""


if __name__ == "__main__":
    app = OcrRunnerApp()
    app.mainloop()
