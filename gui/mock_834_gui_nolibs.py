#!/usr/bin/env python3
"""
mock_834_gui_nolibs.py

Stdlib-only GUI wrapper for mock_834_generator_nolibs.py using Tkinter.

Requirements:
- Python 3.x installed
- NO third-party libraries
- Put this file in the same folder as:
    - mock_834_generator_nolibs.py

How to use:
1) Double-click this file (or run: python mock_834_gui_nolibs.py)
2) Choose an input .xlsx workbook (the template or your filled copy)
3) Choose where to save the output .txt
4) Click "Generate 834"

Notes:
- Runs the generator in a background thread so the UI stays responsive.
- Uses the same Python interpreter as the GUI (sys.executable).
"""

from __future__ import annotations
import subprocess
import sys
from pathlib import Path
import threading
import tkinter as tk
from tkinter import filedialog, messagebox

APP_TITLE = "Mock 834 Generator (No-Libs)"

def run_generator(py_exe: str, generator_py: Path, in_xlsx: Path, out_txt: Path) -> tuple[int, str]:
    cmd = [py_exe, str(generator_py), "--in", str(in_xlsx), "--out", str(out_txt)]
    p = subprocess.run(cmd, capture_output=True, text=True)
    combined = ""
    if p.stdout:
        combined += p.stdout
    if p.stderr:
        combined += ("\n" if combined else "") + p.stderr
    return p.returncode, combined.strip()

def main():
    root = tk.Tk()
    root.title(APP_TITLE)
    root.geometry("780x310")
    root.resizable(False, False)

    generator_py = Path(__file__).with_name("mock_834_generator_nolibs.py")
    if not generator_py.exists():
        messagebox.showerror(
            APP_TITLE,
            f"Couldn't find {generator_py.name} in the same folder.\n\n"
            "Place this GUI file next to mock_834_generator_nolibs.py and try again."
        )
        return

    in_var = tk.StringVar(value="")
    out_var = tk.StringVar(value="")
    status_var = tk.StringVar(value="Ready.")

    def choose_in():
        path = filedialog.askopenfilename(
            title="Select input Excel workbook",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            in_var.set(path)
            default_out = str(Path(path).with_suffix("").with_name(Path(path).stem + "_834.txt"))
            out_var.set(default_out)

    def choose_out():
        path = filedialog.asksaveasfilename(
            title="Save output 834 as",
            defaultextension=".txt",
            filetypes=[("Text File", "*.txt")],
        )
        if path:
            out_var.set(path)

    def set_running(running: bool):
        btn_generate.config(state=("disabled" if running else "normal"))
        btn_browse_in.config(state=("disabled" if running else "normal"))
        btn_save_as.config(state=("disabled" if running else "normal"))

    def generate():
        in_path = Path(in_var.get()).expanduser()
        out_path = Path(out_var.get()).expanduser()

        if not in_path.exists():
            messagebox.showwarning(APP_TITLE, "Pick a valid input .xlsx file.")
            return
        if in_path.suffix.lower() != ".xlsx":
            messagebox.showwarning(APP_TITLE, "No-libs mode supports .xlsx only (not .xls, .xlsm, etc.).")
            return
        if not out_path.parent.exists():
            messagebox.showwarning(APP_TITLE, "Output folder does not exist.")
            return

        set_running(True)
        status_var.set("Generating... (UI stays responsive)")

        py_exe = sys.executable

        def worker():
            rc, output = run_generator(py_exe, generator_py, in_path, out_path)

            def finish():
                set_running(False)
                if rc == 0:
                    status_var.set("Done.")
                    messagebox.showinfo(APP_TITLE, f"Created:\n{out_path}")
                else:
                    status_var.set("Error.")
                    details = output if output else "(no output)"
                    messagebox.showerror(APP_TITLE, "The generator returned an error.\n\nDetails:\n" + details)

            root.after(0, finish)

        threading.Thread(target=worker, daemon=True).start()

    pad = {"padx": 10, "pady": 6}

    tk.Label(root, text="Input Excel (.xlsx):").grid(row=0, column=0, sticky="w", **pad)
    tk.Entry(root, textvariable=in_var, width=76).grid(row=0, column=1, sticky="w", **pad)
    btn_browse_in = tk.Button(root, text="Browse…", command=choose_in, width=12)
    btn_browse_in.grid(row=0, column=2, sticky="e", **pad)

    tk.Label(root, text="Output 834 (.txt):").grid(row=1, column=0, sticky="w", **pad)
    tk.Entry(root, textvariable=out_var, width=76).grid(row=1, column=1, sticky="w", **pad)
    btn_save_as = tk.Button(root, text="Save As…", command=choose_out, width=12)
    btn_save_as.grid(row=1, column=2, sticky="e", **pad)

    btn_generate = tk.Button(root, text="Generate 834", command=generate, width=22, height=2)
    btn_generate.grid(row=2, column=1, sticky="w", padx=10, pady=12)

    tk.Label(root, textvariable=status_var, anchor="w").grid(row=3, column=0, columnspan=3, sticky="we", padx=10)

    help_text = (
        "No-libs mode: reads .xlsx OpenXML (zip+XML). Keep your workbook simple.\n"
        "Put mock_834_gui_nolibs.py + mock_834_generator_nolibs.py in the same folder."
    )
    tk.Label(root, text=help_text, justify="left").grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
