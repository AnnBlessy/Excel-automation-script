What you built is already 80% of a **production-grade pipeline**.
Now we just replace the *local file discovery layer* with **SharePoint ingestion**.

---

# 🔷 1. Updated Architecture (SharePoint Version)

```id="z4s9dl"
SharePoint Site Folders
        ↓
Python (Download / Read Files)
        ↓
UPSERT Logic (same as before)
        ↓
Local Master.xlsx (or SQL later)
```

👉 Important:
**Your core logic DOES NOT change.**
Only the **data source layer changes**.

---

# 🔷 2. Two Ways to Connect Python → SharePoint

## ✅ Option A (Recommended): Microsoft Graph API (Modern & Scalable)

* Secure
* Works with enterprise setups
* Future-proof

## ✅ Option B: SharePoint REST (older, less flexible)

👉 We’ll go with **Graph API (best practice)**

---

# 🔷 3. Setup (One-Time – VERY IMPORTANT)

## 🔹 Step 1: Register App in Azure

Go to:
👉 Microsoft Azure Portal

Steps:

1. Azure Active Directory → App Registrations
2. Click **New Registration**
3. Name:

   ```
   Python-SharePoint-Ingestion
   ```
4. Register

---

## 🔹 Step 2: Add Permissions

Add:

```
Sites.Read.All
Files.Read.All
```

Then:
👉 Click **Grant Admin Consent**

---

## 🔹 Step 3: Create Secret

* Certificates & Secrets → New Client Secret
* Copy value (IMPORTANT)

---

## 🔹 Step 4: Collect These Values

```id="y9v0zs"
TENANT_ID =
CLIENT_ID =
CLIENT_SECRET =
SITE_ID =
DRIVE_ID =
```

---

# 🔷 4. Install Required Libraries

```bash id="4ybdq7"
pip install msal requests pandas openpyxl
```

---

# 🔷 5. SharePoint Folder Example

```id="l4o1k9"
ProjectPlans/
│
├── Site_India/
│   └── ProjectPlan.xlsx
│
├── Site_US/
│   └── ProjectPlan.xlsx
```

---

# 🔷 6. FULL UPDATED SCRIPT (SharePoint + Local Master)

🔥 This replaces your folder-walking logic

```python
import requests
import msal
import pandas as pd
from io import BytesIO
import logging

# =========================
# CONFIG
# =========================

TENANT_ID = "your-tenant-id"
CLIENT_ID = "your-client-id"
CLIENT_SECRET = "your-secret"

SITE_ID = "your-site-id"
DRIVE_ID = "your-drive-id"

MASTER_FILE = r"C:\ProjectAutomation\Master\Master.xlsx"
SHEET_NAME = "Action Tab"

PRIMARY_KEY = ["Project Title", "Customer"]

# =========================
# AUTHENTICATION
# =========================

def get_access_token():
    authority = f"https://login.microsoftonline.com/{TENANT_ID}"
    app = msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=authority,
        client_credential=CLIENT_SECRET
    )

    token = app.acquire_token_for_client(
        scopes=["https://graph.microsoft.com/.default"]
    )

    return token['access_token']


# =========================
# GET FILES FROM SHAREPOINT
# =========================

def list_site_folders(token):

    url = f"https://graph.microsoft.com/v1.0/sites/{SITE_ID}/drives/{DRIVE_ID}/root/children"

    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(url, headers=headers).json()

    return response['value']


def list_files_in_folder(token, folder_id):

    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{folder_id}/children"

    headers = {"Authorization": f"Bearer {token}"}

    return requests.get(url, headers=headers).json()['value']


def download_file(token, file_id):

    url = f"https://graph.microsoft.com/v1.0/drives/{DRIVE_ID}/items/{file_id}/content"

    headers = {"Authorization": f"Bearer {token}"}

    response = requests.get(url, headers=headers)

    return BytesIO(response.content)


# =========================
# CORE PIPELINE (REUSED)
# =========================

def create_key(df):
    return df[PRIMARY_KEY].astype(str).agg('|'.join, axis=1)


def run_pipeline():

    token = get_access_token()

    master_df = pd.read_excel(MASTER_FILE, sheet_name=SHEET_NAME, engine='openpyxl')
    master_df["__key__"] = create_key(master_df)

    all_keys = set()

    folders = list_site_folders(token)

    for folder in folders:

        if "folder" not in folder:
            continue

        site_name = folder['name']
        folder_id = folder['id']

        files = list_files_in_folder(token, folder_id)

        for file in files:

            if not file['name'].endswith(".xlsx"):
                continue

            file_content = download_file(token, file['id'])

            site_df = pd.read_excel(file_content, sheet_name=SHEET_NAME)

            site_df["SiteName"] = site_name
            site_df["__key__"] = create_key(site_df)

            all_keys.update(site_df["__key__"].tolist())

            for _, row in site_df.iterrows():

                key = row["__key__"]

                match = master_df[master_df["__key__"] == key]

                if match.empty:
                    row["RecStatus"] = "Uploaded"
                    master_df = pd.concat([master_df, pd.DataFrame([row])])

                else:
                    idx = match.index[0]

                    if not master_df.loc[idx].equals(row):
                        master_df.loc[idx] = row
                        master_df.at[idx, "RecStatus"] = "Updated"

    # SOFT DELETE
    for idx, row in master_df.iterrows():
        if row["__key__"] not in all_keys:
            master_df.at[idx, "RecStatus"] = "Deleted in Source"

    master_df.drop(columns=["__key__"], inplace=True)

    master_df.to_excel(MASTER_FILE, sheet_name=SHEET_NAME, index=False)


if __name__ == "__main__":
    run_pipeline()
```

---

# 🔷 7. What Changed vs Your Old Script

| Component    | Old           | New        |
| ------------ | ------------- | ---------- |
| File source  | Local folders | SharePoint |
| File reading | `os/pathlib`  | Graph API  |
| Core logic   | Same          | SAME ✅     |

---

# 🔷 8. Key Advantages

✔ No manual file collection
✔ Global access
✔ Secure per-site folders
✔ Fully automated ingestion
✔ Scales to 100+ sites

---

# 🔷 9. Scheduler (Same as Before)

No change — still use:

* Windows Task Scheduler
* Cron

---

# 🔷 10. Important Gotchas

### ⚠️ 1. Excel File Lock

If someone is editing:

* File may fail to read
  👉 Add retry logic if needed

---

### ⚠️ 2. Column Mismatch

Must match template exactly

---

### ⚠️ 3. Permissions

If you get:

```
403 Forbidden
```

👉 Missing admin consent

---

# 🔥 FINAL TAKE

Yes — not only *can* it be altered,
this is actually the **enterprise-grade version** of your solution.

You now have:

✅ Local pipeline
✅ SharePoint ingestion
✅ Fully automated system

---

# 🚀 If you want next level (high impact)

I can help you build:

* 🔁 Incremental sync (only changed files)
* 🧠 Data validation engine (schema + rules)
* 📊 Power BI auto refresh
* 🗄️ SQL version (real enterprise pipeline)
* ⚡ Parallel processing (faster for many sites)

Just tell me 👍
