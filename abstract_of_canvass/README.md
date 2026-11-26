# Abstract of Canvass Document Generator

A local web application built with **Python**, **Streamlit**, and **docxtpl** for generating filled-out official **Abstract of Canvass** Microsoft Word (`.docx`) documents based on `abstract_template.docx` for the Municipality of Maragondon Committee on Awards.

---

## 🚀 Quick Start

### 1. Launch with One Click (Windows)
Double-click `run_app.bat` in this folder.

### 2. Or Run from Command Line
```powershell
python -m streamlit run app.py
```
Then open your web browser at `http://localhost:8501`.

---

## 🛠 Features

1. **Clean Web Form UI**:
   - **Abstract Number**: Text input (e.g. `AC –0433H-10-26`).
   - **Opened Date**: Intuitive Date Picker with automatic formatting (`Month DD, YYYY`) or custom text input override.
   - **Purpose / Description**: Multi-line text area with smart boilerplate trimming to prevent accidental duplicated prefixes/suffixes already present in the official document layout. Includes a live document line preview.
   - **Bidders Section (3 Bidders)**:
     - Bidder Name (e.g. `RBK PRINTING SERVICES`)
     - Location / Address (e.g. `Bucal 1, Maragondon, Cavite`)
     - Amount of Bid (accepts numbers or formatted amounts like `₱ 1,000.00`).
   - **Award Section**:
     - **Automatic Lowest Bidder Detection**: Compares numerical amounts and pre-selects the winning bidder.
     - Manual override options to select any bidder or custom awardee.
     - Automatically updates winning bidder name, location, and amount.
     - Live preview of the bottom award clause.

2. **Accurate Document Generation**:
   - Uses `docxtpl` to populate all placeholders in `abstract_template.docx`:
     - `{{purpose}}`
     - `{{opened_date}}`
     - `{{bidder_1_name}}`, `{{bidder_1_location}}`, `{{bidder_1_amount}}`
     - `{{bidder_2_name}}`, `{{bidder_2_location}}`, `{{bidder_2_amount}}`
     - `{{bidder_3_name}}`, `{{bidder_3_location}}`, `{{bidder_3_amount}}`
     - `{{winning_bidder_name}}`, `{{winning_bidder_location}}`
     - `{{ac_number}}`
   - Preserves all municipal seals, BAC signature committee lines, alignments, and styling.

3. **Summary Preview & Export**:
   - Detailed visual comparison table showing all 3 bidders and winner badge.
   - One-click **Download Word Document** button with auto-named file (`Abstract_of_Canvass_AC_0433H_10_26.docx`).

---

## 📦 Requirements
- Python 3.10+
- `streamlit`
- `docxtpl`
- `python-docx`
