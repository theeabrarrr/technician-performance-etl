# Automated Technician Performance & KPI Reporting Engine

## 📌 Business Problem
In appliance after-sales service management, tracking technician efficiency requires merging multiple disparate raw reports (`CANCEL_NIL_TRANSFER` excel files and quality feedback CSVs). Manual compilation is slow, prone to mismatch errors, and often fails to filter out out-of-zone transferred jobs, distorting true performance metrics.

---

## 💡 The Solution & Architecture
Developed a robust **Python-based Automated ETL & Excel Reporting Pipeline** using `pandas`, `numpy`, and `openpyxl` that:
* **Auto-Fetch & Ingest:** Automatically detects and ingests the latest ERP report files from a temporary directory, moving them into structured raw storage.
* **Data Harmonization & Cleaning:** Standardizes date formats, handles pandas dtype casting safely, and merges datasets using multi-priority sorting (`_PRIORITY` and `_SORT_DATE`).
* **Accurate KPI Filtering:** Excludes out-of-zone transferred jobs (`TRANSFERED`) to ensure performance metrics reflect true technician productivity.
* **Executive Excel Dashboard Generation:** Automatically builds a styled multi-tab Excel workbook containing:
  * Top-tier **KPI Summary Cards** (Assigned, Completed, Efficiency Rate).
  * A structured **Technician Performance Pivot Table** with dynamic Excel formulas (`SUM`, completion percentages).
  * An **Embedded Stacked Horizontal Bar Chart** visualizing workload status breakdown.
  * A formatted **Master Consolidated Data Table** with table styling.

---

## 🛠️ Tech Stack & Skills Demonstrated
* **Language:** Python
* **Libraries:** Pandas, NumPy, OpenPyXL, Pathlib, Shutil
* **Concepts:** Automated ETL Pipelines, Data Harmonization, Programmatic Excel Styling, Chart Embedding, KPI Engine.

---

## 📂 Repository Structure
* `technician_performance_pipeline.py` - Core automation and report generation script.
* `README.md` - Project documentation.

---

## 🚀 How to Run
1. Clone this repository.
2. Ensure your raw ERP reports are dropped into `C:\temp`.
3. Run the script:
   ```bash
   python technician_performance_pipeline.py