#!/usr/bin/env python3
"""
mock_834_gui.py

Tiny GUI wrapper for mock_834_generator.py using Tkinter (built into Python on Windows).

How to use:
1) Put this file in the same folder as mock_834_generator.py
2) Double-click mock_834_gui.py (or run: python mock_834_gui.py)
3) Pick an input Excel workbook (.xlsx)
4) Choose where to save the output .txt
5) Click Generate

Notes:
- This calls the underlying CLI generator so you only maintain one "real" generator.
"""

from __future__ import annotations
import subprocess
import sys
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

APP_TITLE = "Mock 834 Generator"

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
    root.geometry("720x260")
    root.resizable(False, False)

    generator_py = Path(__file__).with_name("mock_834_generator.py")
    if not generator_py.exists():
        messagebox.showerror(
            APP_TITLE,
            f"Couldn't find {generator_py.name} in the same folder.\n\n"
            "Place this GUI file next to mock_834_generator.py and try again."
        )
        return

    in_var = tk.StringVar(value="")
    out_var = tk.StringVar(value="")

    def choose_in():
        path = filedialog.askopenfilename(
            title="Select input Excel workbook",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if path:
            in_var.set(path)
            # Default output next to input
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

    def generate():
        in_path = Path(in_var.get()).expanduser()
        out_path = Path(out_var.get()).expanduser()

        if not in_path.exists():
            messagebox.showwarning(APP_TITLE, "Pick a valid input .xlsx file.")
            return
        if not out_path.parent.exists():
            messagebox.showwarning(APP_TITLE, "Output folder does not exist.")
            return

        rc, output = run_generator(sys.executable, generator_py, in_path, out_path)

        if rc == 0:
            messagebox.showinfo(APP_TITLE, f"Created:\n{out_path}")
        else:
            details = output if output else "(no output)"
            messagebox.showerror(APP_TITLE, "The generator returned an error.\n\nDetails:\n" + details)

    pad = {"padx": 10, "pady": 6}

    tk.Label(root, text="Input Excel (.xlsx):").grid(row=0, column=0, sticky="w", **pad)
    tk.Entry(root, textvariable=in_var, width=70).grid(row=0, column=1, sticky="w", **pad)
    tk.Button(root, text="Browse…", command=choose_in, width=12).grid(row=0, column=2, sticky="e", **pad)

    tk.Label(root, text="Output 834 (.txt):").grid(row=1, column=0, sticky="w", **pad)
    tk.Entry(root, textvariable=out_var, width=70).grid(row=1, column=1, sticky="w", **pad)
    tk.Button(root, text="Save As…", command=choose_out, width=12).grid(row=1, column=2, sticky="e", **pad)

    tk.Label(root, text="").grid(row=2, column=0, **pad)

    tk.Button(root, text="Generate 834", command=generate, width=20, height=2).grid(row=3, column=1, sticky="w", **pad)

    help_text = (
        "Tip: Put mock_834_gui.py + mock_834_generator.py in the same folder.\n"
        "Then fill in your workbook, open the GUI, and click Generate."
    )
    tk.Label(root, text=help_text, justify="left").grid(row=4, column=0, columnspan=3, sticky="w", padx=10, pady=10)

    root.mainloop()

if __name__ == "__main__":
    main()
