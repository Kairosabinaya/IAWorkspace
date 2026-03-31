"""Main view for the Excel Formatter module."""

import os
import subprocess
import threading
import time
from concurrent.futures import Future, ThreadPoolExecutor, as_completed
from pathlib import Path
from tkinter import filedialog, messagebox

import customtkinter as ctk

from app.core import theme
from app.modules.excel_formatter.engine.analyzer import analyze_file
from app.modules.excel_formatter.engine.processor import process_file
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
        self._processing = False
        self._analysis_pool = ThreadPoolExecutor(max_workers=4)
        self._analysis_futures: dict[str, Future] = {}
        self._poll_active = False

        # Throttled progress state — prevents UI event queue flooding
        self._progress_state: dict[str, tuple[float, str]] = {}  # file_name -> (pct, text)
        self._progress_poll_active = False

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
            drop_inner, text="Drag & Drop Excel Files Here",
            font=(theme.FONT_FAMILY, theme.FONT_SIZE_LARGE),
            text_color=theme.TEXT_SECONDARY,
        ).pack(pady=(0, 2))

        ctk.CTkLabel(
            drop_inner, text="or click Browse to select files or a folder",
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

        self._setup_dnd()

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
                paths = []
                for f in file_list:
                    raw = f.decode("utf-8") if isinstance(f, bytes) else str(f)
                    p = Path(raw)
                    if p.is_dir():
                        paths.extend(str(x) for x in p.glob("**/*.xlsx"))
                    else:
                        paths.append(raw)
                self.after(0, lambda: self._add_files(paths))

            windnd.hook_dropfiles(self.winfo_toplevel(), func=_on_drop)
        except ImportError:
            pass

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
        self._add_files([str(f) for f in xlsx_files])

    # ------------------------------------------------------------------
    # File management — batched UI updates to prevent freeze
    # ------------------------------------------------------------------

    def _add_files(self, paths: list[str]):
        """Validate extension (instant), then add files in batches to avoid freeze."""
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

        if not self._output_folder:
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
            )
            self._files[path] = placeholder
            self._file_list.add_file(placeholder)
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
        self._files[path] = config
        config.status = "Ready"
        self._file_list.update_file_status(path, "Ready")
        self._file_list.update_file_details(config)

    def _on_analysis_error(self, path: str, error: str):
        cfg = self._files.get(path)
        if cfg:
            cfg.status = "Error"
            cfg.error_message = error
        self._file_list.update_file_status(path, f"Error: {error}")

    def _remove_file(self, path: str):
        self._files.pop(path, None)
        self._file_list.remove_file(path)

    def _clear_all(self):
        for future in self._analysis_futures.values():
            future.cancel()
        self._analysis_futures.clear()
        self._pending_paths = []
        self._files.clear()
        self._file_list.clear_all()
        self._progress_panel.hide()

    def destroy(self):
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
    # Processing — throttled progress to prevent UI flooding
    # ------------------------------------------------------------------

    def _format_single(self, file_path: str):
        """Format a single file."""
        if self._processing:
            return
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
        config.status = "Processing..."
        self._processing = True
        self._format_btn.configure(state="disabled")
        self._file_list.set_buttons_enabled(False)
        self._progress_panel.show([config.file_name])

        threading.Thread(
            target=self._process_single_bg,
            args=(file_path, config, self._output_folder),
            daemon=True,
        ).start()
        self._start_progress_polling()

    def _process_single_bg(self, file_path: str, config: FileConfig,
                           output_folder: str):
        def progress_cb(file_name, pct, text):
            self._progress_state[file_name] = (pct, text)

        try:
            process_file(config, output_folder, progress_cb)
        except Exception as exc:
            config.status = "Error"
            config.error_message = str(exc)

        self.after(0, self._on_all_done)

    def _start_formatting(self):
        """Format all ready files."""
        if self._processing:
            return

        ready_files = {
            p: c for p, c in self._files.items()
            if c.analyzed and c.status == "Ready"
        }
        if not ready_files:
            messagebox.showinfo("Nothing to Process", "Add some Excel files first.")
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
        for c in ready_files.values():
            c.freeze_pane = freeze
            c.status = "Processing..."

        self._processing = True
        self._format_btn.configure(state="disabled")
        self._file_list.set_buttons_enabled(False)

        file_names = [c.file_name for c in ready_files.values()]
        self._progress_panel.show(file_names)

        threading.Thread(
            target=self._process_all,
            args=(ready_files, self._output_folder),
            daemon=True,
        ).start()
        self._start_progress_polling()

    def _process_all(self, files: dict[str, FileConfig], output_folder: str):
        """Run in background thread. Progress is written to _progress_state dict
        (thread-safe) and read by the UI polling loop — never calls self.after()
        directly, which prevents event queue flooding."""

        def progress_cb(file_name, pct, text):
            # Just write to shared dict — UI polls this periodically
            self._progress_state[file_name] = (pct, text)

        with ThreadPoolExecutor(max_workers=4) as pool:
            futures = {}
            for path, config in files.items():
                f = pool.submit(process_file, config, output_folder, progress_cb)
                futures[f] = config

            for f in as_completed(futures):
                config = futures[f]
                try:
                    f.result()
                except Exception as exc:
                    config.status = "Error"
                    config.error_message = str(exc)
                    self._progress_state[config.file_name] = (0.0, "Error")

        self.after(0, self._on_all_done)

    # ------------------------------------------------------------------
    # Throttled progress UI polling — reads shared dict every 200ms
    # ------------------------------------------------------------------

    def _start_progress_polling(self):
        """Start a 200ms polling loop that reads _progress_state and updates UI."""
        if not self._progress_poll_active:
            self._progress_poll_active = True
            self._progress_state.clear()
            self.after(200, self._poll_progress)

    def _poll_progress(self):
        """Read latest progress from background threads and update UI once."""
        if not self._progress_state:
            if self._processing:
                self.after(200, self._poll_progress)
            return

        # Snapshot and clear
        snapshot = dict(self._progress_state)
        self._progress_state.clear()

        for file_name, (pct, text) in snapshot.items():
            self._progress_panel.update_file(file_name, pct, text)
            for p, c in self._files.items():
                if c.file_name == file_name:
                    self._file_list.update_file_status(p, text)
                    break

        if self._processing:
            self.after(200, self._poll_progress)
        else:
            self._progress_poll_active = False

    def _on_all_done(self):
        # Flush any remaining progress writes
        self._progress_state.clear()
        self._progress_poll_active = False
        self._processing = False

        # Set final status from config (not from progress dict — avoids race)
        for p, c in self._files.items():
            if c.status in ("Done", "Error"):
                self._file_list.update_file_status(p, c.status)
                self._progress_panel.update_file(c.file_name, c.progress, c.status)

        self._format_btn.configure(state="normal")
        self._file_list.set_buttons_enabled(True)
        done = sum(1 for c in self._files.values() if c.status == "Done")
        errors = sum(1 for c in self._files.values() if c.status == "Error")
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
