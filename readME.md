# 🖥️ PC Activity Visualizer (Monthly)

A dynamic, interactive tool built with **Tkinter** and **Matplotlib** to visualize PC login activity across multiple sites/branches — with scrollable UI, threshold filtering, summary exports, and Power BI-friendly output. 📊

---

## 🚀 Features

✅ **Fixed-size plot display**  
- 1200x600 pixels — consistent layout when switching branches  
- Full-screen GUI with vertical scroll for details and summaries

🎚 **Threshold-based filtering**  
- Slider to filter times when a minimum percentage of PCs were active  
- Visualizes login overlaps per day by time-of-day blocks

📈 **Interactive timeline scatterplot**  
- See when PC activity met selected thresholds  
- Time vs. date plots colored for clarity

📤 **Export to Excel**  
- One-click export of all qualifying times and summary usage  
- Output includes: `Branch`, `Threshold (%)`, `Timestamp`, and `PCs Used` (e.g. `"6 of 7"` — **Power BI safe!** ✅)

🧾 **Summary Text Blocks**  
- Displays how many PCs contributed at threshold-qualified times  
- Example: `"🖥 5 of 6 PCs contributed during qualified time blocks (≥60%)"`  
- Lists sample time blocks like:



📂 **Load any monthly Excel login report**  
- Auto-detects unique PCs, login/logout times, and site names  
- Filters data to February 2025 (can be adjusted)

---

## 📁 File Format Requirements

Your Excel file should contain the following columns:

| Column Name   | Type           | Example                  |
|---------------|----------------|--------------------------|
| `Site`        | Text           | `Branch A`, `Main Floor` |
| `Resource`    | Text/PC name   | `HHPC01`, `PC-102`       |
| `Login Time`  | Datetime       | `2/12/2025 09:10 AM`     |
| `Logout Time` | Datetime       | `2/12/2025 12:00 PM`     |

Date format: **`MM/DD/YYYY hh:mm AM/PM`**

---

## 🛠 Tech Stack

- 🐍 Python 3.x  
- 🧰 Tkinter (GUI)  
- 📊 Matplotlib (plotting)  
- 🧮 Pandas + NumPy (data manipulation)  
- 💾 `openpyxl` for Excel exports  

---

## 📦 Installation

Make sure you have Python 3 installed. Then:

```bash
pip install pandas matplotlib openpyxl

LAUNCH WITH: 

python pc_activity_visualizer.py

🧠 Notes
Designed for February 2025 dataset by default.
Change this line in the code if needed:

python
Copy
Edit
df = df[(df['Login Time'] >= '2025-02-01') & (df['Login Time'] < '2025-03-01')].copy()
Exported Excel files are fully Power BI compatible — no fraction/date issues!
"6 of 7" format avoids the dreaded 1/4 → Jan 4 problem. ✅

