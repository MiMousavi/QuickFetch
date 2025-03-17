import sys
import subprocess
import pkg_resources

# List of required packages
required_packages = ['requests', 'pandas', 'XlsxWriter']

# Function to check and install missing packages
def install_missing_packages(packages):
    installed = {pkg.key for pkg in pkg_resources.working_set}
    missing = [pkg for pkg in packages if pkg.lower() not in installed]
    if missing:
        print("Installing missing packages:", missing)
        python = sys.executable
        subprocess.check_call([python, '-m', 'pip', 'install', *missing])
        
install_missing_packages(required_packages)

import os
import re
import base64
import requests
import pandas as pd
from concurrent.futures import ThreadPoolExecutor, as_completed

# -------------------------------
# Configuration - update these!
# -------------------------------
realm = "YourDomain.quickbase.com"            # Your Quickbase realm domain (no protocol)
user_token = "Your API"      				  # Your Quickbase user token
table_id = "tbid123"                          # Table ID from your Quickbase URL
file_field_id = 123                           # File attachment field id
max_workers = 5                               # Number of concurrent threads (adjust as needed)
# -------------------------------

# Common headers for Quickbase API calls
headers = {
    "QB-Realm-Hostname": realm,
    "Authorization": f"QB-USER-TOKEN {user_token}",
    "User-Agent": "Quickbase API Python Script",
    "Accept": "application/json"
}

# Create a session to reuse connections
session = requests.Session()
session.headers.update(headers)

# Folder where attachments will be downloaded
download_folder = "downloads"
os.makedirs(download_folder, exist_ok=True)

# -------------------------------
# Step 1: Retrieve Field Metadata
# -------------------------------
fields_url = f"https://api.quickbase.com/v1/fields?tableId={table_id}"
fields_resp = session.get(fields_url)
if fields_resp.status_code != 200:
    raise Exception(f"Error fetching fields: {fields_resp.status_code} {fields_resp.text}")

fields_data = fields_resp.json()
# Handle if response is a dict or list
if isinstance(fields_data, dict):
    fields = fields_data.get("fields", [])
elif isinstance(fields_data, list):
    fields = fields_data
else:
    fields = []

# Build a mapping from field id (as string) to field label
field_mapping = {}
all_field_ids = []
for field in fields:
    fid = str(field["id"])
    field_mapping[fid] = field.get("label", fid)
    all_field_ids.append(int(fid))  # for query, field ids are integers

# -------------------------------
# Step 2: Query All Records from the Table
# -------------------------------
query_url = "https://api.quickbase.com/v1/records/query"
payload = {
    "from": table_id,
    "select": all_field_ids,
    "options": {"skip": 0, "top": 1000}  # Adjust if you have more records
}

query_resp = session.post(query_url, json=payload, headers={"Content-Type": "application/json"})
if query_resp.status_code != 200:
    raise Exception(f"Error querying records: {query_resp.status_code} {query_resp.text}")

records_data = query_resp.json()
records = records_data.get("data", [])

# -------------------------------
# Step 3: Download Attachments Concurrently
# -------------------------------
def download_attachment(record_id):
    """
    Downloads the file attachment for a given record_id from the specified file_field_id.
    Returns a tuple of (record_id, local_filename) if successful, or (record_id, "") on failure.
    """
    download_url = f"https://api.quickbase.com/v1/files/{table_id}/{record_id}/{file_field_id}/0"
    r = session.get(download_url)
    if r.status_code == 200:
        # Extract filename from Content-Disposition header if available
        content_disp = r.headers.get("Content-Disposition", "")
        default_filename = f"record_{record_id}_file"
        if "filename=" in content_disp:
            part = content_disp.split("filename=")[1]
            original_name = part.strip('\"')
            filename = f"{record_id}_{original_name}"
        else:
            filename = default_filename

        # Sanitize the filename to remove illegal characters
        filename = re.sub(r'[\\/*?:"<>|]', "_", filename)
        # If no extension is present, append .pdf
        if not os.path.splitext(filename)[1]:
            filename += ".pdf"

        local_path = os.path.join(download_folder, filename)
        try:
            file_data = base64.b64decode(r.content)
        except Exception as e:
            print(f"Error decoding file for record {record_id}: {e}")
            return record_id, ""
        with open(local_path, "wb") as f:
            f.write(file_data)
        print(f"Downloaded attachment for record {record_id} as '{filename}'")
        return record_id, filename
    else:
        print(f"Failed to download attachment for record {record_id}. HTTP {r.status_code}: {r.text}")
        return record_id, ""

# Dictionary to store mapping from record id to downloaded attachment filename
attachment_results = {}

with ThreadPoolExecutor(max_workers=max_workers) as executor:
    futures = {}
    # Loop through records and schedule download if file field is present and non-empty
    for rec in records:
        # Get record id, typically in field "3"
        record_id = rec.get("3", {}).get("value", None)
        if record_id is None:
            continue
        file_field_key = str(file_field_id)
        if file_field_key in rec and rec[file_field_key].get("value"):
            futures[executor.submit(download_attachment, record_id)] = record_id

    # Collect results as they complete
    for future in as_completed(futures):
        rec_id, local_file = future.result()
        attachment_results[rec_id] = local_file

# -------------------------------
# Step 4: Generate Final XLSX Report with Clickable Hyperlinks
# -------------------------------
final_records = []
for rec in records:
    record_dict = {}
    record_id = rec.get("3", {}).get("value", "unknown")
    for fid_str, field_obj in rec.items():
        label = field_mapping.get(fid_str, fid_str)
        record_dict[label] = field_obj.get("value", "")
    # Set LocalAttachment column to the downloaded file name (if any)
    record_dict["LocalAttachment"] = attachment_results.get(record_id, "")
    final_records.append(record_dict)

df = pd.DataFrame(final_records)
output_excel = "Quickbase_Table_Report.xlsx"

# Write the DataFrame to Excel with clickable hyperlinks in the LocalAttachment column
with pd.ExcelWriter(output_excel, engine='xlsxwriter') as writer:
    df.to_excel(writer, index=False, sheet_name='Report')
    workbook  = writer.book
    worksheet = writer.sheets['Report']

    # Get the column index of "LocalAttachment"
    col_idx = df.columns.get_loc("LocalAttachment")
    # Loop over the rows (starting at row 1, since row 0 has headers)
    for row_num, attachment in enumerate(df["LocalAttachment"], start=1):
        if attachment:
            # Build the relative file path.
            # Use backslashes (Windows-style) as required and the "file:" prefix.
            file_path = os.path.join("downloads", attachment)
            # Ensure backslashes are used in the hyperlink:
            file_path = file_path.replace("/", "\\")
            # Write a hyperlink in the cell using the "file:" prefix.
            # This creates a clickable link that opens the file from the current folder.
            worksheet.write_url(row_num, col_idx, f'file:{file_path}', string=attachment)

print("Final report generated with clickable attachment links:", output_excel)
