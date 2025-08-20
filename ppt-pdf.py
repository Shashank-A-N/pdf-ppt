import os
import sys
import threading
import queue
import subprocess
import shlex
import time
import traceback
import shutil
from pathlib import Path

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ------------------------------
# Utilities: engine detection
# ------------------------------

def which_libreoffice():
    """
    Return a command string for LibreOffice/soffice if found, else None.
    Searches PATH and common install locations.
    """
    # Try PATH
    for cand in ("soffice", "libreoffice"):
        p = shutil.which(cand)
        if p:
            return p

    # Windows common paths
    if os.name == "nt":
        candidates = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        ]
        for c in candidates:
            if os.path.exists(c):
                return c

    # macOS default app path
    if sys.platform == "darwin":
        c = "/Applications/LibreOffice.app/Contents/MacOS/soffice"
        if os.path.exists(c):
            return c

    return None


def powerpoint_available():
    """
    Return True if PowerPoint COM automation is likely available (Windows + pywin32 + installed PowerPoint).
    We can't be 100% sure until we try to create the COM object.
    """
    if os.name != "nt":
        return False
    try:
        import win32com.client  # noqa: F401
    except Exception:
        return False
    return True


# ------------------------------
# Engines
# ------------------------------

class LibreOfficeEngine:
    def __init__(self, soffice_cmd):
        self.soffice = soffice_cmd

    def convert(self, in_path: str, out_dir: str):
        """
        Convert a single .ppt/.pptx to PDF using LibreOffice CLI.
        Raises subprocess.CalledProcessError on failure.
        """
        in_path = str(Path(in_path).resolve())
        out_dir = str(Path(out_dir).resolve())

        # LibreOffice sometimes prefers quotes; use shlex to be safe across OS
        cmd = [
            self.soffice,
            "--headless",
            "--nologo",
            "--nodefault",
            "--invisible",
            "--nofirststartwizard",
            "--convert-to", "pdf",
            in_path,
            "--outdir", out_dir,
        ]
        subprocess.run(cmd, check=True)

        # Result file path
        stem = Path(in_path).with_suffix("").name
        # LibreOffice keeps original base name, ensures .pdf
        expected_pdf = Path(out_dir) / f"{stem}.pdf"
        if not expected_pdf.exists():
            # Occasionally LO returns a differently-cased extension or minor changes;
            # fallback: pick newest pdf in out_dir with matching stem
            candidates = list(Path(out_dir).glob(f"{stem}*.pdf"))
            if candidates:
                return str(sorted(candidates, key=lambda p: p.stat().st_mtime)[-1])
            raise FileNotFoundError(f"Expected PDF not found for '{in_path}'")
        return str(expected_pdf)


class PowerPointEngine:
    def __init__(self):
        import win32com.client  # type: ignore
        self.client = win32com.client

    def convert(self, in_path: str, out_dir: str):
        """
        Convert using PowerPoint COM SaveAs (format 32 = PDF).
        """
        in_path = str(Path(in_path).resolve())
        out_dir = Path(out_dir).resolve()
        out_dir.mkdir(parents=True, exist_ok=True)
        stem = Path(in_path).with_suffix("").name
        out_pdf = out_dir / f"{stem}.pdf"

        powerpoint = self.client.Dispatch("PowerPoint.Application")
        # Use WithWindow=False to avoid flashing windows
        presentation = None
        try:
            presentation = powerpoint.Presentations.Open(in_path, WithWindow=False)
            # 32 = PDF
            presentation.SaveAs(str(out_pdf), 32)
        finally:
            if presentation is not None:
                presentation.Close()
            powerpoint.Quit()

        if not out_pdf.exists():
            raise FileNotFoundError(f"PowerPoint did not create: {out_pdf}")
        return str(out_pdf)


# ------------------------------
# Converter and worker thread
# ------------------------------

SUPPORTED_EXTS = {".ppt", ".pptx"}

def discover_presentations(paths, recursive=True):
    """
    Expand files/folders into a list of PPT/PPTX file paths.
    """
    files = []
    for p in paths:
        p = Path(p)
        if p.is_file() and p.suffix.lower() in SUPPORTED_EXTS:
            files.append(str(p.resolve()))
        elif p.is_dir():
            if recursive:
                for fp in p.rglob("*"):
                    if fp.suffix.lower() in SUPPORTED_EXTS:
                        files.append(str(fp.resolve()))
            else:
                for fp in p.glob("*"):
                    if fp.suffix.lower() in SUPPORTED_EXTS:
                        files.append(str(fp.resolve()))
    return sorted(set(files))


class ConverterWorker(threading.Thread):
    def __init__(self, tasks, engine_mode, out_dir, log_queue, progress_callback, stop_event):
        """
        tasks: list of input file paths
        engine_mode: "AUTO" | "LIBREOFFICE" | "POWERPOINT"
        out_dir: output directory
        log_queue: queue for UI logs
        progress_callback: callable(done, total)
        stop_event: threading.Event to cancel
        """
        super().__init__(daemon=True)
        self.tasks = tasks
        self.engine_mode = engine_mode
        self.out_dir = out_dir
        self.log_queue = log_queue
        self.progress_callback = progress_callback
        self.stop_event = stop_event

    def _pick_engine(self):
        lo_cmd = which_libreoffice()
        pp_ok = powerpoint_available()

        if self.engine_mode == "LIBREOFFICE":
            if not lo_cmd:
                raise RuntimeError("LibreOffice not found. Install LibreOffice or switch engine.")
            return LibreOfficeEngine(lo_cmd), "LibreOffice"

        if self.engine_mode == "POWERPOINT":
            if not pp_ok:
                raise RuntimeError("PowerPoint COM not available (pywin32/PowerPoint missing).")
            return PowerPointEngine(), "PowerPoint"

        # AUTO
        if pp_ok:
            return PowerPointEngine(), "PowerPoint"
        if lo_cmd:
            return LibreOfficeEngine(lo_cmd), "LibreOffice"
        raise RuntimeError("No conversion engine found. Install LibreOffice or (on Windows) PowerPoint+pywin32.")

    def run(self):
        try:
            engine, name = self._pick_engine()
            self.log_queue.put(f"Using engine: {name}")
            total = len(self.tasks)
            done = 0
            errors = 0

            for src in self.tasks:
                if self.stop_event.is_set():
                    self.log_queue.put("⚠️ Conversion cancelled by user.")
                    break
                self.log_queue.put(f"Converting: {src}")
                try:
                    pdf_path = engine.convert(src, self.out_dir)
                    self.log_queue.put(f"✅ Done: {pdf_path}")
                except Exception as e:
                    errors += 1
                    details = "".join(traceback.format_exception_only(type(e), e)).strip()
                    self.log_queue.put(f"❌ Failed: {src}\n    {details}")
                finally:
                    done += 1
                    self.progress_callback(done, total)

            if errors == 0 and done == total:
                self.log_queue.put("🎉 All files converted successfully.")
            elif errors > 0:
                self.log_queue.put(f"⚠️ Completed with {errors} error(s). Check logs above.")
        except Exception as e:
            details = "".join(traceback.format_exception_only(type(e), e)).strip()
            self.log_queue.put(f"❌ Fatal error: {details}")


# ------------------------------
# GUI
# ------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PPT → PDF Converter (Offline)")
        self.geometry("760x560")
        self.minsize(740, 520)

        # State
        self.input_items = []  # list of file/folder paths
        self.output_dir = os.path.expanduser("~/Desktop/PPT2PDF_Output")
        self.engine_mode = tk.StringVar(value="AUTO")
        self.recursive = tk.BooleanVar(value=True)
        self.worker = None
        self.stop_event = threading.Event()
        self.log_queue = queue.Queue()

        # Top frame: Input controls
        top = ttk.Frame(self, padding=12)
        top.pack(fill="x")

        ttk.Label(top, text="Engine:").pack(side="left")
        ttk.OptionMenu(top, self.engine_mode, "AUTO", "AUTO", "LIBREOFFICE", "POWERPOINT").pack(side="left", padx=(6,12))

        ttk.Checkbutton(top, text="Scan folders recursively", variable=self.recursive).pack(side="left")

        ttk.Button(top, text="Add Files", command=self.add_files).pack(side="left", padx=6)
        ttk.Button(top, text="Add Folder", command=self.add_folder).pack(side="left", padx=6)
        ttk.Button(top, text="Clear List", command=self.clear_list).pack(side="left", padx=6)

        ttk.Button(top, text="Detect Engines", command=self.detect_engines).pack(side="right")

        # Middle frame: Listbox + output dir
        mid = ttk.Frame(self, padding=(12, 6))
        mid.pack(fill="both", expand=True)

        # Input list
        lf = ttk.LabelFrame(mid, text="Selected files & folders")
        lf.pack(side="left", fill="both", expand=True, padx=(0,6))

        self.listbox = tk.Listbox(lf, selectmode="extended")
        self.listbox.pack(fill="both", expand=True, padx=8, pady=8)

        btns = ttk.Frame(lf)
        btns.pack(fill="x", padx=8, pady=(0,8))
        ttk.Button(btns, text="Remove Selected", command=self.remove_selected).pack(side="left")
        ttk.Button(btns, text="Remove Non-existent", command=self.remove_missing).pack(side="left", padx=6)

        # Output dir
        od = ttk.LabelFrame(mid, text="Output Folder")
        od.pack(side="right", fill="both", expand=True, padx=(6,0))

        path_row = ttk.Frame(od)
        path_row.pack(fill="x", padx=8, pady=(8,4))
        self.out_var = tk.StringVar(value=self.output_dir)
        ttk.Entry(path_row, textvariable=self.out_var).pack(side="left", fill="x", expand=True)
        ttk.Button(path_row, text="Browse…", command=self.browse_outdir).pack(side="left", padx=6)
        ttk.Button(path_row, text="Open", command=self.open_outdir).pack(side="left")

        info = ttk.Label(od, text="Converted PDFs will be written here.\nExisting files with the same name will be overwritten.")
        info.pack(anchor="w", padx=8, pady=8)

        # Progress + controls
        bottom = ttk.Frame(self, padding=12)
        bottom.pack(fill="x")

        self.progress = ttk.Progressbar(bottom, length=300, mode="determinate")
        self.progress.pack(side="left", padx=(0, 12))

        self.status_var = tk.StringVar(value="Ready.")
        ttk.Label(bottom, textvariable=self.status_var).pack(side="left", fill="x", expand=True)

        self.start_btn = ttk.Button(bottom, text="Start Conversion", command=self.start_conversion)
        self.start_btn.pack(side="right")
        self.cancel_btn = ttk.Button(bottom, text="Cancel", command=self.cancel_conversion, state="disabled")
        self.cancel_btn.pack(side="right", padx=(0, 8))

        # Log area
        logf = ttk.LabelFrame(self, text="Log")
        logf.pack(fill="both", expand=True, padx=12, pady=(0,12))
        self.log_text = tk.Text(logf, height=12, wrap="word")
        self.log_text.pack(fill="both", expand=True, padx=8, pady=8)
        self.log_text.configure(state="disabled")

        # Periodic UI updates
        self.after(100, self._drain_logs)

        # Style tweaks
        try:
            self.style = ttk.Style(self)
            if self.style.theme_use() in ("vista", "xpnative", "clam", "alt", "default"):
                pass
        except Exception:
            pass

    # ------------- UI actions -------------

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Select PowerPoint files",
            filetypes=[("PowerPoint", "*.ppt *.pptx"), ("All files", "*.*")]
        )
        if not paths:
            return
        self.input_items.extend(list(paths))
        self._refresh_list()

    def add_folder(self):
        folder = filedialog.askdirectory(title="Select a folder")
        if not folder:
            return
        self.input_items.append(folder)
        self._refresh_list()

    def clear_list(self):
        self.input_items.clear()
        self._refresh_list()

    def remove_selected(self):
        sel = list(self.listbox.curselection())
        sel.reverse()
        for idx in sel:
            del self.input_items[idx]
        self._refresh_list()

    def remove_missing(self):
        self.input_items = [p for p in self.input_items if os.path.exists(p)]
        self._refresh_list()

    def browse_outdir(self):
        d = filedialog.askdirectory(title="Choose output folder")
        if d:
            self.output_dir = d
            self.out_var.set(d)

    def open_outdir(self):
        d = self.out_var.get().strip()
        if not d:
            return
        Path(d).mkdir(parents=True, exist_ok=True)
        if sys.platform == "win32":
            os.startfile(d)
        elif sys.platform == "darwin":
            subprocess.run(["open", d])
        else:
            subprocess.run(["xdg-open", d])

    def detect_engines(self):
        lo = which_libreoffice()
        pp = powerpoint_available()
        msg = []
        msg.append(f"LibreOffice: {'FOUND' if lo else 'NOT FOUND'}" + (f"\n  Path: {lo}" if lo else ""))
        msg.append(f"PowerPoint COM (Windows): {'AVAILABLE' if pp else 'NOT AVAILABLE'}")
        messagebox.showinfo("Engine Detection", "\n\n".join(msg))

    def start_conversion(self):
        if self.worker and self.worker.is_alive():
            messagebox.showwarning("Busy", "Conversion is already running.")
            return

        items = list(self.input_items)
        if not items:
            messagebox.showwarning("No input", "Add at least one file or folder.")
            return

        outd = self.out_var.get().strip()
        if not outd:
            messagebox.showwarning("No output folder", "Choose an output folder.")
            return
        Path(outd).mkdir(parents=True, exist_ok=True)

        # Expand items into files to convert
        files = discover_presentations(items, recursive=self.recursive.get())
        if not files:
            messagebox.showwarning("No PPT/PPTX files", "No .ppt/.pptx found in the selection.")
            return

        self.progress.configure(value=0, maximum=len(files))
        self.status_var.set(f"Converting 0 / {len(files)} …")
        self._log(f"Found {len(files)} presentation(s).")

        self.stop_event.clear()
        self.start_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")

        self.worker = ConverterWorker(
            tasks=files,
            engine_mode=self.engine_mode.get(),
            out_dir=outd,
            log_queue=self.log_queue,
            progress_callback=self._on_progress,
            stop_event=self.stop_event
        )
        self.worker.start()

    def cancel_conversion(self):
        if self.worker and self.worker.is_alive():
            self.stop_event.set()
            self._log("Cancelling…")

    # ------------- helpers -------------

    def _on_progress(self, done, total):
        def update():
            self.progress.configure(value=done, maximum=total)
            self.status_var.set(f"Converting {done} / {total} …" if done < total else "Done.")
            if done >= total:
                self.start_btn.configure(state="normal")
                self.cancel_btn.configure(state="disabled")
        self.after(0, update)

    def _refresh_list(self):
        self.listbox.delete(0, "end")
        for p in self.input_items:
            self.listbox.insert("end", p)

    def _log(self, text):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", text + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")

    def _drain_logs(self):
        try:
            while True:
                msg = self.log_queue.get_nowait()
                self._log(msg)
        except queue.Empty:
            pass
        # If worker finished, ensure buttons are reset
        if self.worker and not self.worker.is_alive():
            self.start_btn.configure(state="normal")
            self.cancel_btn.configure(state="disabled")
        self.after(120, self._drain_logs)


if __name__ == "__main__":
    app = App()
    app.mainloop()
