"""Main view for the Excel Formatter module."""

import os
import subprocess
from concurrent.futures import Future, ThreadPoolExecutor
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app.core import theme
from app.modules.excel_formatter.engine.analyzer import analyze_file
from app.modules.excel_formatter.engine.format_queue import FormattingQueue
from app.modules.excel_formatter.gui.config_dialog import ConfigDialog
from app.modules.excel_formatter.gui.file_list_panel import FileListPanel
from app.modules.excel_formatter.gui.progress_panel import ProgressPanel
from app.modules.excel_formatter.models.file_config import FileConfig
from app.utils.file_utils import get_default_output_folder


class FormatterView(ctk.CTkFrame):
    """Full-page view for the Excel Formatter module."""

    def __init__(self, master, **kwargs):
        super().__init__(master, fg_color="transparent", **kwargs)
        self._files: dict[str, FileConfig] = {}  # path -> config
        self._output_folder: str = ""
        self._file_relative_dirs: dict[str, str] = {}  # path -> relative subdir

        # Analysis pool (unchanged — works well)
        self._analysis_pool = ThreadPoolExecutor(max_workers=4)
        self._analysis_futures: dict[str, Future] = {}
        self._poll_active = False

        # Formatting queue (replaces old _processing flag + thread pool)
        self._format_queue = FormattingQueue()
        self._progress_poll_active = False
        # Track whether we already showed the "idle" summary for this batch
        self._idle_notified = True

        self._build_ui()

    # ------------------------------------------------------------------
    # UI construction
    # ------------------------------------------------------------------

    def _build_ui(self):
        # Top section: drag-and-drop zone
        self._drop_zone = ctk.CTkFrame(
            self, fg_color=theme.LIGHT_GRAY, corner_radius=theme.CORNER_RADIUS,
            border_width=2, border_color=theme.BORDER_GRAY, height=160,
        )
        self._drop_zone.pack(fill="x", padx=theme.PADDING_LARGE,
                             pady=(theme.PADDING_LARGE, theme.PADDING_NORMAL))
        self._drop_zone.pack_propagate(False)

        drop_inner = ctk.CTkFrame(self._drop_zone, fg_color="transparent")
        drop_inner.place(relx=0.5, rely=0.5, anchor="center")

        ctk.CTkLabel(
            drop_inner, text=".xlsx",
            font=(theme.FONT_FAMILY, 22, "bold"), text_color=theme.SLATE_GRAY,
        ).pack(pady=(0, 4))

        ctk.CTkLabel(
            drop_inner, text="Select Excel Files or Folder",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE),
            text_color=theme.TEXT_SECONDARY,
        ).pack(pady=(0, 2))

        ctk.CTkLabel(
            drop_inner, text="Click Browse to select files or a folder to format",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            text_color=theme.TEXT_MUTED,
        ).pack(pady=(0, 8))

        btn_frame = ctk.CTkFrame(drop_inner, fg_color="transparent")
        btn_frame.pack()

        ctk.CTkButton(
            btn_frame, text="Browse Files",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            text_color=theme.WHITE, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            command=self._browse_files,
        ).pack(side="left", padx=(0, 8))

        ctk.CTkButton(
            btn_frame, text="Browse Folder",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.WHITE, hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_PRIMARY, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
            command=self._browse_folder,
        ).pack(side="left")

        # Drag-and-drop disabled for now (windnd thread-safety issues)
        # self._setup_dnd()

        # Middle: file list + settings side by side
        mid = ctk.CTkFrame(self, fg_color="transparent")
        mid.pack(fill="both", expand=True, padx=theme.PADDING_LARGE)

        # File list (left, expanding)
        self._file_list = FileListPanel(
            mid, on_configure_click=self._open_config,
            on_remove_click=self._remove_file,
            on_format_click=self._format_single,
        )
        self._file_list.pack(side="left", fill="both", expand=True,
                             padx=(0, theme.PADDING_NORMAL))

        # Right column: settings + progress
        right = ctk.CTkFrame(mid, fg_color="transparent", width=320)
        right.pack(side="right", fill="y")
        right.pack_propagate(False)

        # Settings card
        settings_card = ctk.CTkFrame(
            right, fg_color=theme.WHITE, corner_radius=theme.CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
        )
        settings_card.pack(fill="x", pady=(0, theme.PADDING_NORMAL))

        ctk.CTkLabel(
            settings_card, text="Settings",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE, "bold"),
            text_color=theme.TEXT_PRIMARY, anchor="w",
        ).pack(fill="x", padx=theme.PADDING_LARGE, pady=(theme.PADDING_NORMAL, 4))

        # Freeze pane
        fp_row = ctk.CTkFrame(settings_card, fg_color="transparent")
        fp_row.pack(fill="x", padx=theme.PADDING_LARGE, pady=3)
        ctk.CTkLabel(
            fp_row, text="Freeze Pane:", width=100, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_PRIMARY,
        ).pack(side="left")
        self._freeze_var = ctk.BooleanVar(value=True)
        ctk.CTkSwitch(
            fp_row, text="On", variable=self._freeze_var,
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_PRIMARY,
            progress_color=theme.ACCENT_BLUE,
        ).pack(side="left", padx=(4, 0))

        # Output folder
        of_row = ctk.CTkFrame(settings_card, fg_color="transparent")
        of_row.pack(fill="x", padx=theme.PADDING_LARGE, pady=(3, theme.PADDING_NORMAL))
        ctk.CTkLabel(
            of_row, text="Output:", width=100, anchor="w",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            text_color=theme.TEXT_PRIMARY,
        ).pack(side="left")
        self._output_var = ctk.StringVar(value="(auto)")
        ctk.CTkEntry(
            of_row, textvariable=self._output_var, width=120, state="readonly",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_SMALL),
            fg_color=theme.LIGHT_GRAY, border_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_PRIMARY,
        ).pack(side="left", padx=(4, 4))
        ctk.CTkButton(
            of_row, text="...", width=32, height=28,
            font=(theme.FONT_FAMILY, 14),
            fg_color=theme.LIGHT_GRAY, hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_SECONDARY, border_width=1,
            border_color=theme.BORDER_GRAY,
            command=self._choose_output_folder,
        ).pack(side="left")

        # Progress panel
        self._progress_panel = ProgressPanel(right)
        self._progress_panel.pack(fill="x", pady=(0, theme.PADDING_NORMAL))

        # Bottom action bar
        action_bar = ctk.CTkFrame(self, fg_color="transparent", height=50)
        action_bar.pack(fill="x", padx=theme.PADDING_LARGE,
                        pady=(0, theme.PADDING_LARGE))
        action_bar.pack_propagate(False)

        self._format_btn = ctk.CTkButton(
            action_bar, text="Format All Files",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL, "bold"),
            fg_color=theme.ACCENT_BLUE, hover_color=theme.HOVER_BLUE,
            text_color=theme.WHITE, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            command=self._start_formatting,
        )
        self._format_btn.pack(side="left", padx=(0, theme.PADDING_NORMAL))

        ctk.CTkButton(
            action_bar, text="Open Output Folder",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color=theme.LIGHT_GRAY, hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_PRIMARY, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            border_width=1, border_color=theme.BORDER_GRAY,
            command=self._open_output_folder,
        ).pack(side="left", padx=(0, theme.PADDING_NORMAL))

        ctk.CTkButton(
            action_bar, text="Clear All",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_NORMAL),
            fg_color="transparent", hover_color=theme.BORDER_GRAY,
            text_color=theme.TEXT_SECONDARY, height=theme.BUTTON_HEIGHT,
            corner_radius=theme.BUTTON_CORNER_RADIUS,
            command=self._clear_all,
        ).pack(side="left")

    # ------------------------------------------------------------------
    # Drag and drop
    # ------------------------------------------------------------------

    def _setup_dnd(self):
        try:
            import windnd

            def _on_drop(file_list):
                # windnd callback runs on a background thread — decode paths
                # here (lightweight) and defer all I/O to the main thread.
                raw_paths: list[str] = []
                try:
                    for f in file_list:
                        raw = f.decode("utf-8") if isinstance(f, bytes) else str(f)
                        raw_paths.append(raw)
                except Exception:
                    pass
                if raw_paths:
                    self.after(0, lambda: self._process_dropped_paths(raw_paths))

            windnd.hook_dropfiles(self.winfo_toplevel(), func=_on_drop)
        except ImportError:
            pass

    def _process_dropped_paths(self, raw_paths: list[str]):
        """Resolve dropped paths on the main thread (safe for I/O + UI)."""
        paths: list[str] = []
        roots: dict[str, str] = {}
        for raw in raw_paths:
            p = Path(raw)
            if p.is_dir():
                root = str(p)
                for x in p.glob("**/*.xlsx"):
                    fp = str(x)
                    paths.append(fp)
                    roots[fp] = root
            else:
                paths.append(raw)
        if paths:
            self._add_files(paths, roots=roots)

    def _browse_files(self):
        paths = filedialog.askopenfilenames(
            title="Select Excel Files",
            filetypes=[("Excel Files", "*.xlsx"), ("All Files", "*.*")],
        )
        if paths:
            self._add_files(list(paths))

    def _browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder Containing Excel Files")
        if not folder:
            return
        xlsx_files = list(Path(folder).glob("**/*.xlsx"))
        if not xlsx_files:
            messagebox.showinfo(
                "No Excel Files",
                "No .xlsx files were found in the selected folder.",
            )
            return
        self._add_files([str(f) for f in xlsx_files], root_folder=folder)

    # ------------------------------------------------------------------
    # File management — batched UI updates to prevent freeze
    # ------------------------------------------------------------------

    def _add_files(
        self,
        paths: list[str],
        root_folder: str = "",
        roots: dict[str, str] | None = None,
    ):
        """Validate extension (instant), then add files in batches to avoid freeze.

        Args:
            paths: List of file paths.
            root_folder: Common root folder (from Browse Folder).
            roots: Per-file root folders (from drag-drop with mixed sources).
        """
        valid = []
        rejected = []
        for p in paths:
            if p.lower().endswith(".xlsx") and p not in self._files:
                valid.append(p)
            elif not p.lower().endswith(".xlsx"):
                rejected.append(os.path.basename(p))

        if rejected:
            self.after(0, lambda: messagebox.showwarning(
                "Unsupported Files",
                "The following files are not .xlsx files and were skipped:\n\n"
                + "\n".join(rejected[:20])
                + (f"\n...and {len(rejected) - 20} more" if len(rejected) > 20 else ""),
            ))

        if not valid:
            return

        # Compute relative subdirectory for each file (preserves folder structure)
        for p in valid:
            root = (roots or {}).get(p, root_folder)
            rel = ""
            if root:
                rel = os.path.relpath(os.path.dirname(p), root)
                if rel == ".":
                    rel = ""
            self._file_relative_dirs[p] = rel

        if not self._output_folder:
            # Use the root folder (if available) for a cleaner output base
            first_root = (roots or {}).get(valid[0], root_folder)
            if first_root:
                from app.utils.constants import DEFAULT_OUTPUT_FOLDER
                self._output_folder = str(Path(first_root) / DEFAULT_OUTPUT_FOLDER)
            else:
                self._output_folder = get_default_output_folder(valid[0])
            self._output_var.set(self._output_folder)

        # Add files in batches of 5 via after() to keep UI responsive
        self._pending_paths = list(valid)
        self._add_next_batch()

    def _add_next_batch(self):
        """Add up to 5 files to the UI, then yield to the event loop."""
        batch = self._pending_paths[:5]
        self._pending_paths = self._pending_paths[5:]

        for path in batch:
            placeholder = FileConfig(
                file_path=path,
                file_name=os.path.basename(path),
                file_size="...",
                status="Analyzing...",
                relative_dir=self._file_relative_dirs.get(path, ""),
            )
            self._files[path] = placeholder
            self._file_list.add_file(placeholder)
            self._file_list.set_file_buttons_state(path, "analyzing")
            future = self._analysis_pool.submit(analyze_file, path)
            self._analysis_futures[path] = future

        # Start analysis polling if not already active
        if not self._poll_active and self._analysis_futures:
            self._poll_active = True
            self.after(150, self._poll_analysis_results)

        # Schedule next batch if there are more files
        if self._pending_paths:
            self.after(10, self._add_next_batch)

    def _poll_analysis_results(self):
        """Check analysis futures every 150ms and update UI."""
        done_paths = []
        for path, future in self._analysis_futures.items():
            if future.done():
                done_paths.append(path)
                try:
                    config = future.result()
                    self._on_analysis_done(path, config)
                except Exception as exc:
                    self._on_analysis_error(path, str(exc))

        for p in done_paths:
            del self._analysis_futures[p]

        if self._analysis_futures:
            self.after(150, self._poll_analysis_results)
        else:
            self._poll_active = False

    def _on_analysis_done(self, path: str, config: FileConfig):
        # Preserve relative_dir computed during _add_files
        old = self._files.get(path)
        if old and old.relative_dir:
            config.relative_dir = old.relative_dir
        self._files[path] = config
        config.status = "Ready"
        self._file_list.update_file_status(path, "Ready")
        self._file_list.update_file_details(config)
        self._file_list.set_file_buttons_state(path, "ready")

    def _on_analysis_error(self, path: str, error: str):
        cfg = self._files.get(path)
        if cfg:
            cfg.status = "Error"
            cfg.error_message = error
        self._file_list.update_file_status(path, f"Error: {error}")
        self._file_list.set_file_buttons_state(path, "error")

    def _remove_file(self, path: str):
        # If queued, cancel it first
        if self._format_queue.is_job_active(path):
            if self._format_queue.is_job_processing(path):
                messagebox.showwarning(
                    "Cannot Remove",
                    "This file is currently being formatted and cannot be removed.",
                )
                return
            self._format_queue.cancel(path)
            self._progress_panel.remove_file(os.path.basename(path))

        self._files.pop(path, None)
        self._file_list.remove_file(path)

    def _clear_all(self):
        if not self._format_queue.is_idle():
            if not messagebox.askyesno(
                "Formatting in Progress",
                "Formatting is in progress. Clear all files and cancel "
                "queued items?\n\n(The file currently being formatted will "
                "finish processing.)",
            ):
                return

        self._format_queue.cancel_all_queued()
        for future in self._analysis_futures.values():
            future.cancel()
        self._analysis_futures.clear()
        self._pending_paths = []
        self._files.clear()
        self._file_relative_dirs.clear()
        self._file_list.clear_all()
        self._progress_panel.hide()

    def destroy(self):
        self._format_queue.shutdown()
        self._analysis_pool.shutdown(wait=False, cancel_futures=True)
        super().destroy()

    # ------------------------------------------------------------------
    # Configuration dialog
    # ------------------------------------------------------------------

    def _open_config(self, file_path: str):
        config = self._files.get(file_path)
        if not config or not config.analyzed:
            messagebox.showinfo("Please Wait", "File is still being analyzed.")
            return
        dialog = ConfigDialog(
            self.winfo_toplevel(), config, date_format=config.date_format,
        )
        self.wait_window(dialog)
        self._file_list.update_file_details(config)

    # ------------------------------------------------------------------
    # Settings
    # ------------------------------------------------------------------

    def _choose_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self._output_folder = folder
            self._output_var.set(folder)

    # ------------------------------------------------------------------
    # Formatting — queue-based
    # ------------------------------------------------------------------

    def _format_single(self, file_path: str):
        """Enqueue a single file for formatting."""
        config = self._files.get(file_path)
        if not config or not config.analyzed:
            messagebox.showinfo("Please Wait", "File is still being analyzed.")
            return
        if config.status not in ("Ready", "Done", "Error"):
            return

        if not self._output_folder:
            self._output_folder = get_default_output_folder(file_path)
            self._output_var.set(self._output_folder)

        out = Path(self._output_folder)
        if out.exists() and (out / config.file_name).exists():
            if not messagebox.askyesno(
                "File Exists",
                f"{config.file_name} already exists in the output folder. Overwrite?",
            ):
                return

        config.freeze_pane = self._freeze_var.get()
        self._enqueue_file(file_path, config)

    def _start_formatting(self):
        """Enqueue all ready files for formatting."""
        ready_files = {
            p: c for p, c in self._files.items()
            if c.analyzed and c.status in ("Ready",)
        }
        if not ready_files:
            messagebox.showinfo("Nothing to Process",
                                "No ready files to format. Add some Excel files first.")
            return

        if not self._output_folder:
            first_path = next(iter(ready_files))
            self._output_folder = get_default_output_folder(first_path)
            self._output_var.set(self._output_folder)

        out = Path(self._output_folder)
        if out.exists():
            existing = [
                c.file_name for c in ready_files.values()
                if (out / c.file_name).exists()
            ]
            if existing:
                if not messagebox.askyesno(
                    "Files Already Exist",
                    "The following files already exist in the output folder:\n\n"
                    + "\n".join(existing) + "\n\nOverwrite?",
                ):
                    return

        freeze = self._freeze_var.get()
        for path, config in ready_files.items():
            config.freeze_pane = freeze
            self._enqueue_file(path, config)

    def _enqueue_file(self, file_path: str, config: FileConfig):
        """Add one file to the formatting queue and update UI."""
        self._format_queue.enqueue(config, self._output_folder)
        self._file_list.update_file_status(file_path, "Queued")
        self._file_list.set_file_buttons_state(file_path, "queued")
        self._progress_panel.add_file(config.file_name)
        self._idle_notified = False
        self._start_progress_polling()

    # ------------------------------------------------------------------
    # Throttled progress UI polling — reads shared dict every 200ms
    # ------------------------------------------------------------------

    def _start_progress_polling(self):
        """Start a 200ms polling loop that reads progress and job state."""
        if not self._progress_poll_active:
            self._progress_poll_active = True
            self.after(200, self._poll_progress)

    def _poll_progress(self):
        """Read latest progress from the queue worker and update UI."""
        # 1. Read progress dict (same pattern as before)
        if self._format_queue.progress_state:
            snapshot = dict(self._format_queue.progress_state)
            self._format_queue.progress_state.clear()

            for file_name, (pct, text) in snapshot.items():
                self._progress_panel.update_file(file_name, pct, text)
                # Also update file list status
                for p, c in self._files.items():
                    if c.file_name == file_name:
                        self._file_list.update_file_status(p, text)
                        break

        # 2. Sync per-file button states from job statuses
        for job in self._format_queue.get_all_jobs():
            path = job.job_id
            if job.status == "processing":
                self._file_list.set_file_buttons_state(path, "processing")
            elif job.status == "queued":
                self._file_list.set_file_buttons_state(path, "queued")
            elif job.status in ("done", "error"):
                self._file_list.set_file_buttons_state(path, job.status)

        # 3. Check if queue just became idle
        if self._format_queue.is_idle() and not self._idle_notified:
            self._idle_notified = True
            self.after(0, self._on_queue_idle)
            self._progress_poll_active = False
            return

        # Keep polling while work remains
        if not self._format_queue.is_idle():
            self.after(200, self._poll_progress)
        else:
            self._progress_poll_active = False

    def _on_queue_idle(self):
        """Called when the formatting queue has drained."""
        completed = self._format_queue.pop_completed()
        if not completed:
            return

        # Update final statuses from completed jobs
        for job in completed:
            path = job.job_id
            config = job.config
            status = "Done" if job.status == "done" else "Error"
            self._file_list.update_file_status(path, status)
            self._file_list.set_file_buttons_state(path, job.status)
            self._progress_panel.update_file(
                config.file_name,
                1.0 if job.status == "done" else 0.0,
                status,
            )

        done = sum(1 for j in completed if j.status == "done")
        errors = sum(1 for j in completed if j.status == "error")
        msg = f"Formatting complete: {done} succeeded"
        if errors:
            msg += f", {errors} failed"
        messagebox.showinfo("Done", msg)

    def _open_output_folder(self):
        folder = self._output_folder
        if not folder or not Path(folder).exists():
            messagebox.showinfo("No Output", "No output folder to open yet.")
            return
        subprocess.Popen(["explorer", os.path.normpath(folder)])
