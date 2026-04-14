# Excel-automation-script

# 🔷 1. Big Picture (What You’re Building)

You are implementing a **data ingestion pipeline**:

```
Site Excel Files → Python Script → Master Excel (UPSERT logic)
```

Aligned with your document’s recommendation :

* Automated ingestion ✔
* No manual merge ✔
* Site isolation ✔
* Master reporting dataset ✔

---

# 🔷 2. Folder Structure (LOCAL VERSION)

You can later extend to SharePoint, but start locally:

```
C:\ProjectAutomation\
│
├── Sites\
│   ├── Site_India\
│   │   └── ProjectPlan.xlsx
│   │
│   ├── Site_US\
│   │   └── ProjectPlan.xlsx
│   │
│   └── Site_Germany\
│       └── ProjectPlan.xlsx
│
├── Master\
│   └── Master.xlsx
│
└── logs\
    └── ingestion.log
```

---

# 🔷 3. Master File Requirements

Your **Master.xlsx must already exist** with:

* Same columns as template
* Extra column:

  ```
  RecStatus
  SiteName   (IMPORTANT for deletion logic)
  ```

👉 No new tables — only row updates ✔

---

# 🔷 4. Core Logic (UPSERT Engine)

We will implement:

| Scenario          | Action                       |
| ----------------- | ---------------------------- |
| New record        | Add → RecStatus = "Uploaded" |
| Existing changed  | Update → "Updated"           |
| Missing in source | Mark → "Deleted in Source"   |
| Same              | Keep / "Unchanged"           |

---

# 🔷 5. FULL PYTHON SCRIPT (Production Ready)

### ✅ Install dependencies first:

```bash
pip install pandas openpyxl
```

---

## 🔥 SCRIPT: `ingestion_pipeline.py`

```python
import pandas as pd
import os
from pathlib import Path
import logging

# =========================
# CONFIGURATION (EDIT HERE)
# =========================

ROOT_DIR = Path(r"C:\ProjectAutomation\Sites")
MASTER_FILE = Path(r"C:\ProjectAutomation\Master\Master.xlsx")
LOG_FILE = Path(r"C:\ProjectAutomation\logs\ingestion.log")

SHEET_NAME = "Action Tab"  # must match template

PRIMARY_KEY = ["Project Title", "Customer"]

TRACKING_COLUMNS = ["RecStatus", "SiteName"]

# =========================
# LOGGING SETUP
# =========================

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)

# =========================
# HELPER FUNCTIONS
# =========================

def load_excel(file_path):
    try:
        df = pd.read_excel(file_path, sheet_name=SHEET_NAME, engine='openpyxl')
        return df
    except Exception as e:
        logging.error(f"Failed to read {file_path}: {e}")
        return None


def create_composite_key(df):
    return df[PRIMARY_KEY].astype(str).agg('|'.join, axis=1)


def normalize_df(df):
    df.columns = df.columns.str.strip()
    return df


# =========================
# MAIN LOGIC
# =========================

def run_pipeline():

    logging.info("===== PIPELINE STARTED =====")

    # Load master
    if MASTER_FILE.exists():
        master_df = pd.read_excel(MASTER_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
    else:
        logging.error("Master file not found!")
        return

    master_df = normalize_df(master_df)

    if "SiteName" not in master_df.columns:
        master_df["SiteName"] = ""

    if "RecStatus" not in master_df.columns:
        master_df["RecStatus"] = ""

    master_df["__key__"] = create_composite_key(master_df)

    updated_master = master_df.copy()

    all_site_keys = set()

    # =========================
    # LOOP THROUGH SITES
    # =========================

    for site_folder in ROOT_DIR.iterdir():

        if not site_folder.is_dir():
            continue

        site_name = site_folder.name

        logging.info(f"Processing site: {site_name}")

        excel_files = list(site_folder.glob("*.xlsx"))

        if not excel_files:
            logging.warning(f"No Excel file found in {site_name}")
            continue

        file_path = excel_files[0]

        site_df = load_excel(file_path)

        if site_df is None:
            continue

        site_df = normalize_df(site_df)

        site_df["SiteName"] = site_name

        site_df["__key__"] = create_composite_key(site_df)

        all_site_keys.update(site_df["__key__"].tolist())

        # =========================
        # PROCESS EACH ROW
        # =========================

        for _, row in site_df.iterrows():

            key = row["__key__"]

            master_match = updated_master[updated_master["__key__"] == key]

            if master_match.empty:
                # NEW RECORD
                new_row = row.copy()
                new_row["RecStatus"] = "Uploaded"

                updated_master = pd.concat(
                    [updated_master, pd.DataFrame([new_row])],
                    ignore_index=True
                )

            else:
                # EXISTING RECORD
                idx = master_match.index[0]

                existing_row = updated_master.loc[idx]

                # Compare excluding tracking columns
                compare_cols = [col for col in site_df.columns if col not in ["RecStatus", "__key__"]]

                if not existing_row[compare_cols].equals(row[compare_cols]):
                    # UPDATE
                    for col in compare_cols:
                        updated_master.at[idx, col] = row[col]

                    updated_master.at[idx, "RecStatus"] = "Updated"
                else:
                    # UNCHANGED
                    if pd.isna(updated_master.at[idx, "RecStatus"]):
                        updated_master.at[idx, "RecStatus"] = "Unchanged"

    # =========================
    # SOFT DELETE LOGIC
    # =========================

    for idx, row in updated_master.iterrows():

        key = row["__key__"]

        if key not in all_site_keys:
            updated_master.at[idx, "RecStatus"] = "Deleted in Source"

    # =========================
    # CLEANUP
    # =========================

    updated_master.drop(columns=["__key__"], inplace=True)

    # =========================
    # WRITE BACK TO MASTER
    # =========================

    with pd.ExcelWriter(MASTER_FILE, engine="openpyxl", mode="w") as writer:
        updated_master.to_excel(writer, sheet_name=SHEET_NAME, index=False)

    logging.info("===== PIPELINE COMPLETED SUCCESSFULLY =====")


# =========================
# RUN SCRIPT
# =========================

if __name__ == "__main__":
    run_pipeline()
```

---

# 🔷 6. KEY DESIGN DECISIONS (IMPORTANT)

### ✅ Composite Key

```
Project Title + Customer
```

### ✅ Site Tracking

Needed for:

* Deletion detection
* Data ownership

### ✅ No Table Recreation

We overwrite sheet but:

* Keep same schema ✔
* No new tables ✔

---

# 🔷 7. ERROR HANDLING INCLUDED

✔ File read failure
✔ Missing files
✔ Logging per site

Log file:

```
logs/ingestion.log
```

---

# 🔷 8. AUTOMATION (Windows Task Scheduler)

## 🔹 Step-by-step

1. Open:

   ```
   Task Scheduler
   ```

2. Click:

   ```
   Create Basic Task
   ```

3. Name:

   ```
   Project Data Pipeline
   ```

4. Trigger:

   * Monthly / Quarterly

5. Action:

   ```
   Start a Program
   ```

6. Program:

   ```
   python
   ```

7. Arguments:

   ```
   "C:\ProjectAutomation\ingestion_pipeline.py"
   ```

8. Finish ✔

---

## 🔹 Alternative (Using .bat file)

Create `run_pipeline.bat`:

```bat
@echo off
cd C:\ProjectAutomation
python ingestion_pipeline.py
```

Then schedule the `.bat` file.

---

# 🔷 9. CRON JOB (LINUX / SERVER)

```bash
crontab -e
```

Example (monthly):

```bash
0 2 1 * * /usr/bin/python3 /path/to/ingestion_pipeline.py
```

---

# 🔷 10. BEST PRACTICES (VERY IMPORTANT)

### 🔹 1. Lock Template

As suggested in your doc :

* Prevent column changes
* Use dropdowns

### 🔹 2. Add These Columns (Optional but powerful)

| Column           | Purpose      |
| ---------------- | ------------ |
| IngestionDate    | Audit        |
| FileModifiedDate | Traceability |

---

### 🔹 3. Future Upgrade (Recommended)

Move Master to:

* SQL Server / Azure SQL

Why:

* Faster
* No corruption risk
* Better analytics

---

# 🔷 11. WHAT YOU NOW HAVE

You now have:

✅ Fully automated ingestion
✅ UPSERT logic
✅ Soft delete handling
✅ Logging
✅ Scheduler setup
✅ Scalable architecture
