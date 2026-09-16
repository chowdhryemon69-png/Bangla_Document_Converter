import os
import queue
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

import converter as converter_module
from converter import convert_docx_file


APP_DIR = Path(__file__).resolve().parent
TARGET_FONT = "Nikosh"


def _packaged_runtime_dir():
    if getattr(sys, "frozen", False):
        return Path(sys._MEIPASS) / "runtime"
    return APP_DIR


def _packaged_log_dir():
    local_app_data = os.environ.get("LOCALAPPDATA")
    base_dir = Path(local_app_data) if local_app_data else Path.home() / "AppData" / "Local"
    return base_dir / "BanglaDocumentConverter" / "logs"


def _configure_packaged_runtime():
    if not getattr(sys, "frozen", False):
        return

    runtime_dir = _packaged_runtime_dir()
    os.environ["PATH"] = str(runtime_dir) + os.pathsep + os.environ.get("PATH", "")

    original_logger = converter_module.log_conversion_result

    def user_log(source_file, output_file, success, validation_status, error_message=None, log_dir=None):
        return original_logger(
            source_file,
            output_file,
            success,
            validation_status,
            error_message,
            log_dir=_packaged_log_dir(),
        )

    converter_module.log_conversion_result = user_log


_configure_packaged_runtime()


class ConverterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Bangla Document Converter")
        self.root.geometry("650x650")
        self.root.minsize(590, 600)
        self.root.configure(bg="#f4f6f8")

        self.source_path = None
        self.output_path = None
        self.events = queue.Queue()
        self.worker = None

        self.source_text = tk.StringVar(value="No file selected")
        self.output_mode = tk.StringVar(value="Same folder")
        self.output_text = tk.StringVar(value="Choose a DOCX file first")
        self.progress_text = tk.StringVar(value="Ready")
        self.status_text = tk.StringVar(value="Select a DOCX file to begin.")
        self.result_text = tk.StringVar(value="Ready")
        self.result_detail = tk.StringVar(value="")
        self.file_controls = []

        self._build_style()
        self._build_ui()
        self.root.after(100, self._drain_events)

    def _build_style(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")
        style.configure("App.TFrame", background="#f4f6f8")
        style.configure("Card.TFrame", background="#ffffff")
        style.configure("Title.TLabel", background="#ffffff", foreground="#17324d", font=("Segoe UI Semibold", 23))
        style.configure("Subtitle.TLabel", background="#ffffff", foreground="#637587", font=("Segoe UI", 11))
        style.configure("Section.TLabel", background="#ffffff", foreground="#17324d", font=("Segoe UI Semibold", 11))
        style.configure("Body.TLabel", background="#ffffff", foreground="#34495e", font=("Segoe UI", 10))
        style.configure("Muted.TLabel", background="#ffffff", foreground="#738496", font=("Segoe UI", 10))
        style.configure("Primary.TButton", font=("Segoe UI Semibold", 11), padding=(20, 10))
        style.configure("Secondary.TButton", font=("Segoe UI", 10), padding=(10, 7))
        style.configure("TCombobox", padding=5)
        style.configure("Horizontal.TProgressbar", troughcolor="#e5ebf0", background="#1976a8", thickness=12)

    def _build_ui(self):
        outer = ttk.Frame(self.root, style="App.TFrame", padding=24)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Developer: Chowdhury Emon", style="Muted.TLabel").pack(side="bottom", anchor="w")
        card = ttk.Frame(outer, style="Card.TFrame", padding=28)
        card.pack(fill="both", expand=True)

        ttk.Label(card, text="Bangla Document Converter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(card, text="Bijoy / SutonnyMJ → Unicode Bangla", style="Subtitle.TLabel").pack(anchor="w", pady=(3, 24))

        select_button = ttk.Button(card, text="Select DOCX File", command=self.select_file, style="Primary.TButton")
        select_button.pack(anchor="w")
        self.file_controls.append(select_button)
        ttk.Label(card, text="Selected file:", style="Section.TLabel").pack(anchor="w", pady=(20, 2))
        ttk.Label(card, textvariable=self.source_text, style="Body.TLabel").pack(anchor="w")

        ttk.Label(card, text="Output:", style="Section.TLabel").pack(anchor="w", pady=(18, 5))
        output_row = ttk.Frame(card, style="Card.TFrame")
        output_row.pack(fill="x")
        self.output_combo = ttk.Combobox(output_row, textvariable=self.output_mode, values=("Same folder", "Choose location..."), state="readonly", width=18)
        self.output_combo.pack(side="left")
        self.output_combo.bind("<<ComboboxSelected>>", self._output_mode_changed)
        choose_button = ttk.Button(output_row, text="Choose...", command=self.choose_output, style="Secondary.TButton")
        choose_button.pack(side="left", padx=(10, 0))
        self.file_controls.extend((self.output_combo, choose_button))
        ttk.Label(card, textvariable=self.output_text, style="Muted.TLabel").pack(anchor="w", pady=(5, 0))

        ttk.Label(card, text="Font:", style="Section.TLabel").pack(anchor="w", pady=(16, 2))
        ttk.Label(card, text=TARGET_FONT, style="Body.TLabel").pack(anchor="w")

        self.convert_button = ttk.Button(card, text="CONVERT", command=self.start_conversion, style="Primary.TButton")
        self.convert_button.pack(pady=(23, 24))
        self.file_controls.append(self.convert_button)

        ttk.Separator(card).pack(fill="x")
        ttk.Label(card, text="Progress", style="Section.TLabel").pack(anchor="w", pady=(18, 8))
        self.progress = ttk.Progressbar(card, orient="horizontal", mode="determinate", maximum=100, value=0, style="Horizontal.TProgressbar")
        self.progress.pack(fill="x")
        ttk.Label(card, textvariable=self.progress_text, style="Muted.TLabel").pack(anchor="e", pady=(4, 0))
        ttk.Label(card, text="Status:", style="Section.TLabel").pack(anchor="w", pady=(8, 2))
        ttk.Label(card, textvariable=self.status_text, style="Body.TLabel").pack(anchor="w")

        ttk.Label(card, text="Result:", style="Section.TLabel").pack(anchor="w", pady=(20, 2))
        ttk.Label(card, textvariable=self.result_text, style="Body.TLabel").pack(anchor="w")
        ttk.Label(card, textvariable=self.result_detail, style="Muted.TLabel", justify="left").pack(anchor="w", pady=(4, 0))
        actions = ttk.Frame(card, style="Card.TFrame")
        actions.pack(fill="x", pady=(10, 0))
        self.open_button = ttk.Button(actions, text="Open Output", command=self.open_output, state="disabled", style="Secondary.TButton")
        self.open_button.pack(side="left")
        self.folder_button = ttk.Button(actions, text="Open Folder", command=self.open_folder, state="disabled", style="Secondary.TButton")
        self.folder_button.pack(side="left", padx=(8, 0))
        self.another_button = ttk.Button(actions, text="Convert Another", command=self.reset, state="disabled", style="Secondary.TButton")
        self.another_button.pack(side="right")

    def select_file(self):
        selected = filedialog.askopenfilename(title="Select a DOCX file", filetypes=(("Word documents", "*.docx"),))
        if not selected:
            return
        path = Path(selected)
        if path.suffix.lower() != ".docx":
            messagebox.showerror("Unsupported file", "Please select a .docx document.")
            return
        self.source_path = path
        self.output_path = None
        self.source_text.set(path.name)
        self.output_mode.set("Same folder")
        self._refresh_output_text()
        self.result_text.set("Ready")
        self.result_detail.set("")
        self.status_text.set("Ready to convert.")

    def _default_output(self):
        return self.source_path.with_name(f"{self.source_path.stem}_Unicode_Nikosh.docx")

    def _refresh_output_text(self):
        if not self.source_path:
            self.output_text.set("Choose a DOCX file first")
        elif self.output_mode.get() == "Same folder":
            self.output_path = self._default_output()
            self.output_text.set(self.output_path.name)
        elif self.output_path:
            self.output_text.set(str(self.output_path))
        else:
            self.output_text.set("Choose an output location")

    def _output_mode_changed(self, _event=None):
        if self.output_mode.get() == "Same folder":
            self._refresh_output_text()
        else:
            self.choose_output()

    def choose_output(self):
        if not self.source_path:
            messagebox.showinfo("Select a file", "Select a DOCX file first.")
            return
        selected = filedialog.asksaveasfilename(title="Choose output DOCX", initialfile=self._default_output().name, defaultextension=".docx", filetypes=(("Word documents", "*.docx"),))
        if selected:
            self.output_path = Path(selected)
            self.output_mode.set("Choose location...")
            self._refresh_output_text()

    def start_conversion(self):
        if not self.source_path:
            messagebox.showinfo("Select a file", "Select a DOCX file first.")
            return
        self._refresh_output_text()
        if not self.output_path:
            return
        if self.output_path.exists() and not messagebox.askyesno("Replace existing file?", f"{self.output_path.name} already exists. Replace it?"):
            return

        self._set_busy(True)
        self.progress.configure(value=0)
        self.progress_text.set("0%")
        self.status_text.set("Starting conversion...")
        self.result_text.set("Converting")
        self.worker = threading.Thread(target=self._convert_worker, args=(self.source_path, self.output_path), daemon=True)
        self.worker.start()

    def _convert_worker(self, source, output):
        try:
            bridge_file = _packaged_runtime_dir() / "bijoy_bridge.mjs"
            result = convert_docx_file(
                source,
                output,
                bridge_file=bridge_file,
                progress_callback=lambda status, percent: self.events.put(("progress", status, percent)),
            )
            self.events.put(("success", result))
        except Exception as exc:
            self.events.put(("failure", str(exc)))

    def _drain_events(self):
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, status, percent = event
                    self.progress.configure(value=percent)
                    self.progress_text.set(f"{percent}%")
                    self.status_text.set(status)
                elif event[0] == "success":
                    result = event[1]
                    self._set_busy(False)
                    self.progress.configure(value=100)
                    self.progress_text.set("100%")
                    self.status_text.set("Complete")
                    self.result_text.set("Conversion completed successfully")
                    self.result_detail.set(
                        f"Source: {Path(result['source_file']).name}\n"
                        f"Output: {Path(result['output_file']).name}\n"
                        f"Validation: {result['validation']['details']}"
                    )
                    self.open_button.configure(state="normal")
                    self.folder_button.configure(state="normal")
                    self.another_button.configure(state="normal")
                    messagebox.showinfo("Conversion complete", f"Converted {Path(result['source_file']).name} successfully.")
                elif event[0] == "failure":
                    self._set_busy(False)
                    self.result_text.set("Conversion failed")
                    self.result_detail.set("")
                    self.status_text.set("The document could not be converted.")
                    messagebox.showerror("Conversion failed", event[1].splitlines()[0])
        except queue.Empty:
            pass
        self.root.after(100, self._drain_events)

    def _set_busy(self, busy):
        state = "disabled" if busy else "normal"
        for control in self.file_controls:
            control.configure(state=state)
        if not busy:
            self.output_combo.configure(state="readonly")

    def open_output(self):
        if self.output_path and self.output_path.exists():
            os.startfile(self.output_path)

    def open_folder(self):
        if self.output_path:
            os.startfile(self.output_path.parent)

    def reset(self):
        self.source_path = None
        self.output_path = None
        self.source_text.set("No file selected")
        self.output_mode.set("Same folder")
        self._refresh_output_text()
        self.progress.configure(value=0)
        self.progress_text.set("0%")
        self.status_text.set("Select a DOCX file to begin.")
        self.result_text.set("Ready")
        self.result_detail.set("")
        self.open_button.configure(state="disabled")
        self.folder_button.configure(state="disabled")
        self.another_button.configure(state="disabled")


def main():
    root = tk.Tk()
    ConverterApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()