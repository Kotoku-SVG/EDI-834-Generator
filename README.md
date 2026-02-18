
# Mock 834 Generator – All‑in‑One Package

This folder contains everything needed to generate a **mock EDI 834 enrollment file** from Excel.

---

## 📁 Folder Structure

excel_template/
- mock_834_generator_template.xlsx

generators/
- mock_834_generator.py  (requires pandas + openpyxl)
- mock_834_generator_nolibs.py  (NO external libraries)

gui/
- mock_834_gui.py  (GUI for library version)
- mock_834_gui_nolibs.py  (GUI for no‑libs version)

run_834.bat  (simple double‑click runner for library version)

output_examples/
- sample_output_834.txt

single_exe/
- .exe builder, for power users to build for end-users.

---

## 🚀 Quick Start (Easiest)

1. Open:
   excel_template/mock_834_generator_template.xlsx

2. Fill in:
   - Settings
   - Plans
   - Members
   - Dependents (optional)

3. Run ONE of these:

### Option A – GUI (recommended)
Double‑click:

gui/mock_834_gui_nolibs.py (mock_834_gui.py is preferred if you have pandas + openpyxl installed)

Choose the Excel file → choose save location → click Generate.

(No installs required.)

---

### Option B – Batch file (library version)
If you installed pandas/openpyxl:

Double‑click:
run_834.bat

Output:
out_834.txt

---

### Option C – Command line

Library version:
python generators/mock_834_generator.py --in yourfile.xlsx --out out_834.txt

No‑libs version:
python generators/mock_834_generator_nolibs.py --in yourfile.xlsx --out out_834.txt

---

## 🧾 Output Naming

GUI default:
yourfile_834.txt

Batch default:
out_834.txt

---

## ⚠️ No‑Libraries Mode Notes

The no‑libs generator:

✔ Works with standard Python only  
✔ Reads .xlsx directly  
✔ Good for portability  

Limitations:
- Only supports .xlsx
- Cells should contain typed values (avoid formulas)
- Uses header row 1 on Plans/Members/Dependents
- Settings read from Column A (key) / Column B (value)

---

## 🧠 What This Generates

A simplified EDI 834 including:

- ISA / GS / ST envelope
- BGN + as‑of DTP
- Sponsor + Payer N1
- Subscriber loops
- Dependent loops
- HD coverage segments
- Proper SE/GE/IEA closing

Designed for:
- Integration testing
- Demo files
- Mock enrollment loads

---

## 💬 Tip

If sharing with others, just send this entire folder.
They can run the no‑libs GUI immediately with zero setup.

---
