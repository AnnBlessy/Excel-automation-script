Let’s build this step-by-step in a **clean, production-ready way**.

---

# 🔷 1. What “Delta Processing” Means

Instead of:

```text
Every run → Read ALL site files ❌
```

You do:

```text
Only read files that changed since last run ✅
```

---

# 🔷 2. How We Detect Changes

We use **file metadata from SharePoint**:

### ✅ Best signal:

```text
lastModifiedDateTime
```

Every file in SharePoint has:

```json
"lastModifiedDateTime": "2026-04-10T08:21:00Z"
```

---

# 🔷 3. Core Idea

We store a **checkpoint file locally**:

```text
last_run.json
```

Example:

```json
{
  "Site_India_ProjectPlan.xlsx": "2026-04-01T10:00:00Z",
  "Site_US_ProjectPlan.xlsx": "2026-04-05T12:00:00Z"
}
```

---

# 🔷 4. Pipeline with Delta Logic

```text
1. Load last_run.json
2. Fetch SharePoint files
3. Compare timestamps
4. Process ONLY changed files
5. Update last_run.json
```

---

# 🔷 5. Implementation (Step-by-Step)

---

## 🔥 Step 1: Add Config

```python
CHECKPOINT_FILE = "last_run.json"
```

---

## 🔥 Step 2: Load Checkpoint

```python
import json
from pathlib import Path

def load_checkpoint():
    if Path(CHECKPOINT_FILE).exists():
        with open(CHECKPOINT_FILE, "r") as f:
            return json.load(f)
    return {}
```

---

## 🔥 Step 3: Save Checkpoint

```python
def save_checkpoint(data):
    with open(CHECKPOINT_FILE, "w") as f:
        json.dump(data, f, indent=4)
```

---

## 🔥 Step 4: Modify File Processing Logic

Update your file loop:

```python
checkpoint = load_checkpoint()
new_checkpoint = {}

for folder in folders:

    if "folder" not in folder:
        continue

    site_name = folder['name']
    folder_id = folder['id']

    files = list_files_in_folder(token, folder_id)

    for file in files:

        if not file['name'].endswith(".xlsx"):
            continue

        file_id = file['id']
        file_name = f"{site_name}_{file['name']}"
        last_modified = file['lastModifiedDateTime']

        new_checkpoint[file_name] = last_modified

        # 🔥 DELTA CHECK
        if file_name in checkpoint:
            if checkpoint[file_name] == last_modified:
                print(f"⏭ Skipping unchanged file: {file_name}")
                continue

        print(f"🔄 Processing changed file: {file_name}")

        file_content = download_file(token, file_id)

        df = pd.read_excel(file_content, sheet_name=SHEET_NAME)

        # continue your logic...
```

---

## 🔥 Step 5: Save Checkpoint at End

```python
save_checkpoint(new_checkpoint)
```

---

# 🔷 6. Important Design Behavior

### ✅ First Run

* No checkpoint → ALL files processed

### ✅ Next Runs

* Only changed files processed

---

# 🔷 7. HUGE IMPACT 🚀

| Scenario | Without Delta | With Delta      |
| -------- | ------------- | --------------- |
| 50 sites | 50 files      | maybe 3–5 files |
| Time     | 5–7 min       | ~20–30 sec      |

---

# 🔷 8. Edge Case Handling (IMPORTANT)

### ⚠️ 1. File Deleted from SharePoint

👉 Still need soft delete logic
(Your master comparison already handles this ✔)

---

### ⚠️ 2. File Renamed

👉 Treated as new file
(acceptable in most systems)

---

### ⚠️ 3. Corrupted Checkpoint

👉 Add fallback:

```python
try:
    checkpoint = load_checkpoint()
except:
    checkpoint = {}
```

---

# 🔷 9. Advanced Upgrade (Optional but Powerful)

Instead of timestamp, use:

### 🔥 File Hash (More Accurate)

```text
hash = MD5(file content)
```

Detects:

* Even small changes ✔

But:

* Slightly slower

---

# 🔷 10. Combine with Parallel Processing

👉 Best combo:

```text
Parallel + Delta = 🔥 Maximum performance
```

Flow:

1. Filter changed files
2. Process them in parallel
3. Merge once

---

# 🔷 11. Final Optimized Pipeline

```text
1. Authenticate
2. Load checkpoint
3. Get SharePoint files
4. Filter changed files
5. Parallel process
6. UPSERT to master
7. Save checkpoint
```

---

# 🔥 12. Real-World Insight

This is exactly how:

* Data pipelines (ETL)
* Big data systems
* Cloud ingestion jobs

are designed.

👉 You just moved from:
**Script → Data Engineering System**
