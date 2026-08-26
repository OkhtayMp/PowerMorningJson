from __future__ import annotations

import json
import threading
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import ttk
from tkinterdnd2 import DND_FILES, TkinterDnD
from platformdirs import user_data_dir


# ============================================================
# Configuration
# ============================================================

APP_TITLE = "Morning XLSM Extractor"

WINDOW_WIDTH = 620
WINDOW_HEIGHT = 400

INPUT_FILE_NAME = "Morning.xlsm"
SOURCE_SHEET = "HO To MS Sites"

AUTHOR = "OkhtayMp"
APP_NAME = "PowerTTman"

# Application data directory
DATA_DIR = Path(
    user_data_dir(APP_NAME, AUTHOR)
)

# Final JSON file
OUTPUT_FILE = DATA_DIR / "MorninJson.json"


# ============================================================
# Colors
# ============================================================

BG = "#0D1117"
CARD = "#161B22"
CARD_2 = "#1C2128"
BORDER = "#30363D"

TEXT = "#F0F6FC"
TEXT_SECONDARY = "#8B949E"

ACCENT = "#58A6FF"
ACCENT_HOVER = "#79C0FF"

SUCCESS = "#3FB950"
ERROR = "#F85149"


# ============================================================
# Application
# ============================================================

class MorningExtractor(TkinterDnD.Tk):

    def __init__(self):
        super().__init__()

        self.title(APP_TITLE)
        self.geometry(
            f"{WINDOW_WIDTH}x{WINDOW_HEIGHT}"
        )
        self.resizable(False, False)
        self.configure(bg=BG)

        self.processing = False
        self.animation_job = None
        self.animation_index = 0

        self.setup_styles()
        self.build_initial_ui()

    # ========================================================
    # Styles
    # ========================================================

    def setup_styles(self):

        style = ttk.Style(self)

        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Modern.Horizontal.TProgressbar",
            troughcolor=CARD_2,
            background=ACCENT,
            bordercolor=CARD_2,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            thickness=11
        )

    # ========================================================
    # Initial UI
    # ========================================================

    def build_initial_ui(self):

        self.clear_window()

        self.root = tk.Frame(
            self,
            bg=BG
        )

        self.root.pack(
            fill="both",
            expand=True,
            padx=24,
            pady=20
        )

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        tk.Label(
            self.root,
            text="Morning XLSM Extractor",
            bg=BG,
            fg=TEXT,
            font=("Arial", 21, "bold")
        ).pack(
            anchor="w"
        )

        tk.Label(
            self.root,
            text="Extract Site ID and Priority → JSON",
            bg=BG,
            fg=TEXT_SECONDARY,
            font=("Arial", 9)
        ).pack(
            anchor="w",
            pady=(3, 18)
        )

        # ----------------------------------------------------
        # Drop Card
        # ----------------------------------------------------

        self.drop_card = tk.Frame(
            self.root,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        self.drop_card.pack(
            fill="x"
        )

        # Accent line
        tk.Frame(
            self.drop_card,
            bg=ACCENT,
            height=3
        ).pack(
            fill="x"
        )

        inner = tk.Frame(
            self.drop_card,
            bg=CARD
        )

        inner.pack(
            fill="x",
            padx=25,
            pady=25
        )

        tk.Label(
            inner,
            text="↓",
            bg=CARD,
            fg=ACCENT,
            font=("Arial", 28, "bold")
        ).pack()

        tk.Label(
            inner,
            text="DROP MORNING.XLSM HERE",
            bg=CARD,
            fg=TEXT,
            font=("Arial", 16, "bold")
        ).pack(
            pady=(3, 4)
        )

        tk.Label(
            inner,
            text="Drag & Drop",
            bg=CARD,
            fg=TEXT_SECONDARY,
            font=("Arial", 9)
        ).pack()

        # ----------------------------------------------------
        # Required File
        # ----------------------------------------------------

        info = tk.Frame(
            inner,
            bg=CARD_2
        )

        info.pack(
            pady=(15, 0)
        )

        tk.Label(
            info,
            text="REQUIRED",
            bg=CARD_2,
            fg=TEXT_SECONDARY,
            font=("Arial", 7, "bold")
        ).pack(
            side="left",
            padx=(10, 6),
            pady=6
        )

        tk.Label(
            info,
            text=INPUT_FILE_NAME,
            bg=CARD_2,
            fg=ACCENT_HOVER,
            font=("Arial", 8, "bold")
        ).pack(
            side="left",
            padx=(0, 10),
            pady=6
        )

        # ----------------------------------------------------
        # Output info
        # ----------------------------------------------------

        output_card = tk.Frame(
            self.root,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        output_card.pack(
            fill="x",
            pady=(10, 0)
        )

        output_inner = tk.Frame(
            output_card,
            bg=CARD
        )

        output_inner.pack(
            fill="x",
            padx=16,
            pady=10
        )

        tk.Label(
            output_inner,
            text="OUTPUT",
            bg=CARD,
            fg=TEXT_SECONDARY,
            font=("Arial", 7, "bold")
        ).pack(
            anchor="w"
        )

        tk.Label(
            output_inner,
            text="MorninJson.json",
            bg=CARD,
            fg=TEXT,
            font=("Arial", 9, "bold")
        ).pack(
            anchor="w",
            pady=(2, 0)
        )

        # ----------------------------------------------------
        # Source sheet
        # ----------------------------------------------------

        tk.Label(
            self.root,
            text=f"Source sheet: {SOURCE_SHEET}",
            bg=BG,
            fg=TEXT_SECONDARY,
            font=("Arial", 8)
        ).pack(
            anchor="w",
            pady=(11, 0)
        )

        # ----------------------------------------------------
        # Drag & Drop
        # ----------------------------------------------------

        self.drop_card.drop_target_register(
            DND_FILES
        )

        self.drop_card.dnd_bind(
            "<<Drop>>",
            self.on_file_drop
        )

    # ========================================================
    # Processing UI
    # ========================================================

    def build_processing_ui(self):

        self.clear_window()

        self.root = tk.Frame(
            self,
            bg=BG
        )

        self.root.pack(
            fill="both",
            expand=True,
            padx=24,
            pady=20
        )

        # ----------------------------------------------------
        # Header
        # ----------------------------------------------------

        tk.Label(
            self.root,
            text="Building JSON",
            bg=BG,
            fg=TEXT,
            font=("Arial", 21, "bold")
        ).pack(
            anchor="w"
        )

        tk.Label(
            self.root,
            text="Processing Morning.xlsm",
            bg=BG,
            fg=TEXT_SECONDARY,
            font=("Arial", 9)
        ).pack(
            anchor="w",
            pady=(3, 18)
        )

        # ----------------------------------------------------
        # Progress Card
        # ----------------------------------------------------

        progress_card = tk.Frame(
            self.root,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        progress_card.pack(
            fill="x"
        )

        tk.Frame(
            progress_card,
            bg=ACCENT,
            height=3
        ).pack(
            fill="x"
        )

        progress_inner = tk.Frame(
            progress_card,
            bg=CARD
        )

        progress_inner.pack(
            fill="x",
            padx=18,
            pady=17
        )

        status_row = tk.Frame(
            progress_inner,
            bg=CARD
        )

        status_row.pack(
            fill="x"
        )

        self.status_title = tk.Label(
            status_row,
            text="Preparing...",
            bg=CARD,
            fg=TEXT,
            font=("Arial", 9, "bold")
        )

        self.status_title.pack(
            side="left"
        )

        self.percent_label = tk.Label(
            status_row,
            text="0%",
            bg=CARD,
            fg=ACCENT,
            font=("Arial", 17, "bold")
        )

        self.percent_label.pack(
            side="right"
        )

        self.progress = ttk.Progressbar(
            progress_inner,
            style="Modern.Horizontal.TProgressbar",
            orient="horizontal",
            mode="determinate",
            maximum=100
        )

        self.progress.pack(
            fill="x",
            pady=(10, 8)
        )

        self.counter_label = tk.Label(
            progress_inner,
            text="0 / 0 records",
            bg=CARD,
            fg=TEXT_SECONDARY,
            font=("Arial", 8)
        )

        self.counter_label.pack(
            anchor="w"
        )

        # ----------------------------------------------------
        # Status Card
        # ----------------------------------------------------

        status_card = tk.Frame(
            self.root,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        status_card.pack(
            fill="x",
            pady=(10, 0)
        )

        status_inner = tk.Frame(
            status_card,
            bg=CARD
        )

        status_inner.pack(
            fill="x",
            padx=18,
            pady=13
        )

        tk.Label(
            status_inner,
            text="STATUS",
            bg=CARD,
            fg=TEXT_SECONDARY,
            font=("Arial", 7, "bold")
        ).pack(
            anchor="w"
        )

        self.activity_label = tk.Label(
            status_inner,
            text="Processing...",
            bg=CARD,
            fg=TEXT,
            font=("Arial", 9)
        )

        self.activity_label.pack(
            anchor="w",
            pady=(3, 0)
        )

        self.dots_label = tk.Label(
            status_inner,
            text="",
            bg=CARD,
            fg=ACCENT,
            font=("Arial", 9, "bold")
        )

        self.dots_label.place(
            relx=0.16,
            rely=0.50
        )

        tk.Label(
            self.root,
            text=f"Output directory: {DATA_DIR}",
            bg=BG,
            fg=TEXT_SECONDARY,
            font=("Arial", 7)
        ).pack(
            anchor="w",
            pady=(11, 0)
        )

        self.animate_loading()

    # ========================================================
    # Completion UI
    # ========================================================

    def build_completion_ui(self):

        # Everything from previous screen is removed.
        self.clear_window()

        root = tk.Frame(
            self,
            bg=BG
        )

        root.pack(
            fill="both",
            expand=True,
            padx=24,
            pady=20
        )

        # ----------------------------------------------------
        # Completion Card
        # ----------------------------------------------------

        card = tk.Frame(
            root,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        card.pack(
            fill="x",
            pady=(38, 0)
        )

        # Success accent
        tk.Frame(
            card,
            bg=SUCCESS,
            height=4
        ).pack(
            fill="x"
        )

        inner = tk.Frame(
            card,
            bg=CARD
        )

        inner.pack(
            fill="x",
            padx=25,
            pady=30
        )

        # Success icon
        tk.Label(
            inner,
            text="✓",
            bg=CARD,
            fg=SUCCESS,
            font=("Arial", 38, "bold")
        ).pack()

        # Success message
        tk.Label(
            inner,
            text="Completed successfully",
            bg=CARD,
            fg=TEXT,
            font=("Arial", 20, "bold")
        ).pack(
            pady=(8, 7)
        )

        tk.Label(
            inner,
            text="MorninJson.json has been created successfully.",
            bg=CARD,
            fg=TEXT_SECONDARY,
            font=("Arial", 10)
        ).pack()

        # ----------------------------------------------------
        # Large Exit button
        # ----------------------------------------------------

        tk.Button(
            root,
            text="EXIT",
            command=self.destroy,
            bg=ACCENT,
            fg=BG,
            activebackground=ACCENT_HOVER,
            activeforeground=BG,
            relief="flat",
            bd=0,
            font=("Arial", 12, "bold"),
            cursor="hand2",
            height=2
        ).pack(
            fill="x",
            pady=(20, 0)
        )

    # ========================================================
    # Error UI
    # ========================================================

    def build_error_ui(self, error):

        self.clear_window()

        root = tk.Frame(
            self,
            bg=BG
        )

        root.pack(
            fill="both",
            expand=True,
            padx=24,
            pady=20
        )

        card = tk.Frame(
            root,
            bg=CARD,
            highlightbackground=BORDER,
            highlightthickness=1
        )

        card.pack(
            fill="x",
            pady=(45, 0)
        )

        tk.Frame(
            card,
            bg=ERROR,
            height=4
        ).pack(
            fill="x"
        )

        inner = tk.Frame(
            card,
            bg=CARD
        )

        inner.pack(
            fill="x",
            padx=25,
            pady=30
        )

        tk.Label(
            inner,
            text="!",
            bg=CARD,
            fg=ERROR,
            font=("Arial", 36, "bold")
        ).pack()

        tk.Label(
            inner,
            text="Processing failed",
            bg=CARD,
            fg=TEXT,
            font=("Arial", 20, "bold")
        ).pack(
            pady=(8, 8)
        )

        tk.Label(
            inner,
            text=error,
            bg=CARD,
            fg=TEXT_SECONDARY,
            font=("Arial", 10),
            wraplength=500,
            justify="center"
        ).pack()

        tk.Button(
            root,
            text="EXIT",
            command=self.destroy,
            bg=ERROR,
            fg="#FFFFFF",
            activebackground="#FF6B63",
            activeforeground="#FFFFFF",
            relief="flat",
            bd=0,
            font=("Arial", 12, "bold"),
            cursor="hand2",
            height=2
        ).pack(
            fill="x",
            pady=(20, 0)
        )

    # ========================================================
    # Clear Window
    # ========================================================

    def clear_window(self):

        for widget in self.winfo_children():
            widget.destroy()

    # ========================================================
    # File Drop
    # ========================================================

    def on_file_drop(self, event):

        if self.processing:
            return

        files = self.tk.splitlist(
            event.data
        )

        if not files:
            return

        input_file = Path(
            files[0]
        )

        # Exact filename validation
        if (
            input_file.name.lower()
            != INPUT_FILE_NAME.lower()
        ):

            self.build_error_ui(
                f"Please drop {INPUT_FILE_NAME}."
            )

            return

        # Start
        self.processing = True

        self.build_processing_ui()

        worker = threading.Thread(
            target=self.process_file,
            args=(input_file,),
            daemon=True
        )

        worker.start()

    # ========================================================
    # Process Excel
    # ========================================================

    def process_file(
        self,
        input_file: Path
    ):

        try:

            # ------------------------------------------------
            # Open workbook
            # ------------------------------------------------

            self.set_activity(
                "Opening Morning.xlsm..."
            )

            workbook = pd.ExcelFile(
                input_file,
                engine="openpyxl"
            )

            # ------------------------------------------------
            # Find source sheet
            # ------------------------------------------------

            target_sheet = None

            for sheet in workbook.sheet_names:

                if (
                    sheet.strip().lower()
                    == SOURCE_SHEET.lower()
                ):
                    target_sheet = sheet
                    break

            if target_sheet is None:

                raise ValueError(
                    f"Sheet '{SOURCE_SHEET}' was not found."
                )

            # ------------------------------------------------
            # Read sheet
            # ------------------------------------------------

            self.set_activity(
                "Reading HO To MS Sites..."
            )

            df = pd.read_excel(
                input_file,
                sheet_name=target_sheet,
                engine="openpyxl"
            )

            # ------------------------------------------------
            # Validate columns
            # ------------------------------------------------

            if "Site ID" not in df.columns:

                raise ValueError(
                    "Column 'Site ID' was not found."
                )

            if "Priority" not in df.columns:

                raise ValueError(
                    "Column 'Priority' was not found."
                )

            # ------------------------------------------------
            # Extract required columns
            # ------------------------------------------------

            df = df[
                ["Site ID", "Priority"]
            ].copy()

            # Remove empty Site ID rows
            df = df.dropna(
                subset=["Site ID"]
            )

            total = len(df)

            if total == 0:

                raise ValueError(
                    "No valid Site ID records were found."
                )

            self.after(
                0,
                self.set_total_records,
                total
            )

            self.set_activity(
                "Separating Site ID and Priority..."
            )

            data = []

            # ------------------------------------------------
            # Build JSON
            # ------------------------------------------------

            for index, row in enumerate(
                df.itertuples(index=False),
                start=1
            ):

                site_id = str(
                    row[0]
                ).strip()

                priority = (
                    ""
                    if pd.isna(row[1])
                    else str(row[1]).strip()
                )

                data.append(
                    {
                        "site_id": site_id,
                        "priority": priority
                    }
                )

                percent = (
                    index / total
                ) * 100

                self.after(
                    0,
                    self.update_progress,
                    percent,
                    index,
                    total
                )

            # ------------------------------------------------
            # Create application data directory
            # ------------------------------------------------

            self.set_activity(
                "Preparing output directory..."
            )

            DATA_DIR.mkdir(
                parents=True,
                exist_ok=True
            )

            # ------------------------------------------------
            # Save JSON
            # ------------------------------------------------

            self.set_activity(
                "Writing MorninJson.json..."
            )

            with OUTPUT_FILE.open(
                "w",
                encoding="utf-8"
            ) as json_file:

                json.dump(
                    data,
                    json_file,
                    ensure_ascii=False,
                    indent=2
                )

            # ------------------------------------------------
            # Finished
            # ------------------------------------------------

            self.after(
                0,
                self.processing_finished
            )

        except Exception as error:

            self.after(
                0,
                self.processing_error,
                str(error)
            )

    # ========================================================
    # UI Updates
    # ========================================================

    def set_activity(self, text):

        self.after(
            0,
            lambda: self.activity_label.config(
                text=text
            )
        )

    def set_total_records(self, total):

        self.counter_label.config(
            text=f"0 / {total:,} records"
        )

    def update_progress(
        self,
        percent,
        current,
        total
    ):

        self.progress["value"] = percent

        self.percent_label.config(
            text=f"{percent:.1f}%"
        )

        self.counter_label.config(
            text=f"{current:,} / {total:,} records"
        )

    # ========================================================
    # Loading Animation
    # ========================================================

    def animate_loading(self):

        if not self.processing:
            return

        self.animation_index = (
            self.animation_index + 1
        ) % 4

        self.dots_label.config(
            text="." * self.animation_index
        )

        self.animation_job = self.after(
            350,
            self.animate_loading
        )

    # ========================================================
    # Finished
    # ========================================================

    def processing_finished(self):

        self.processing = False

        if self.animation_job:

            try:
                self.after_cancel(
                    self.animation_job
                )
            except tk.TclError:
                pass

            self.animation_job = None

        # Completely replace the old UI
        self.build_completion_ui()

    # ========================================================
    # Error
    # ========================================================

    def processing_error(self, error):

        self.processing = False

        if self.animation_job:

            try:
                self.after_cancel(
                    self.animation_job
                )
            except tk.TclError:
                pass

            self.animation_job = None

        # Replace UI with error screen
        self.build_error_ui(
            error
        )


# ============================================================
# Main
# ============================================================

if __name__ == "__main__":

    app = MorningExtractor()
    app.mainloop()