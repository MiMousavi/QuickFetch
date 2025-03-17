# QuickFetch
 Python utility that retrieves all records and fields from a QuickBase table, downloads file attachments concurrently, and generates an Excel (XLSX) report with clickable links to the locally downloaded files.

# QuickFetch: QuickBase Attachment Downloader

QuickFetch is an open source Python tool designed to streamline the process of retrieving data from QuickBase. It fetches all records and fields from a specified QuickBase table, downloads file attachments concurrently, and generates an Excel (XLSX) report with clickable links to the downloaded files.

## Features

- **Comprehensive Data Retrieval:**  
  Automatically retrieves field metadata and all records from a QuickBase table.

- **Concurrent Attachment Downloads:**  
  Utilizes multithreading to download file attachments quickly, saving them to a local `downloads` folder.

- **Excel Report Generation:**  
  Produces an XLSX report that includes all table data along with clickable hyperlinks pointing to each downloaded attachment.

- **Automatic Dependency Handling:**  
  The script checks for and installs required packages (`requests`, `pandas`, and `XlsxWriter`) automatically if they are missing.

## Alternative Solutions

There are several commercial solutions that provide QuickBase integration and file attachment management. 
For example, the **FileDown+ add-on** is a paid tool offering similar functionality. 
However, these commercial options require a subscription or one-time fee. 
QuickFetch is offered as a completely free and open source alternative that you can customize to fit your needs.

## QuickBase Table and Field IDs Explained

To use QuickFetch, you need to provide:
- **Table ID:**  
When you navigate to a QuickBase table in your browser, the URL typically looks like:  
https://YourDomain.quickbase.com/nav/app/<appID>/table/<tableID>/action/td?skip=0

In this URL, `<tableID>` is the identifier for the table (for example, `tbid123`). Simply copy this value into the configuration section of QuickFetch.

- **Field ID:**  
Each field in a QuickBase table is assigned a unique numerical ID. To find the field ID:
1. Open your QuickBase table.
2. Go to the table settings or click on the **Fields** section.
3. Locate the file attachment field you wish to download (e.g., the field where file attachments are stored). The field ID is usually displayed alongside the field label in the settings.
4. Use that numerical ID (for example, `7`) in the configuration section of QuickFetch.

## Requirements

- Python 3.x  
- Internet connectivity to access the QuickBase API  
- A valid QuickBase user token with permissions to access the target table and file attachments

## Installation

Simply clone the repository and run the script. QuickFetch will automatically install any missing dependencies.

## Configuration

Before running QuickFetch, edit the configuration section at the top of `quickfetch.py`:

- **realm:** Your QuickBase realm domain (e.g., `YourDomain.quickbase.com`)  
- **user_token:** Your QuickBase user token  
- **table_id:** The QuickBase table ID (e.g., `tbid123`) — see the **QuickBase Table and Field IDs Explained** section  
- **file_field_id:** The field ID for the file attachment (e.g., `7`) — see the instructions above  
- **max_workers:** (Optional) Number of threads for concurrent downloads (default is 5)

## Usage

Run QuickFetch from the command line:

```bash
python quickfetch.py

========================================

The script will:

Retrieve field metadata and records from the specified QuickBase table.
Download attachments concurrently into a local downloads folder.
Generate an Excel report (Quickbase_Table_Report.xlsx) 
containing all table data with a "LocalAttachment" column that provides clickable links to each downloaded file.
Note: The hyperlinks in the Excel report are relative links (e.g., downloads\record_1_file.pdf), so ensure that the downloads folder remains alongside the Excel report.

File Structure
After running the script, your project folder should look like this:


quickfetch.py                # The main script file
downloads/                   # Folder with downloaded attachment files
Quickbase_Table_Report.xlsx  # Generated Excel report with clickable links
README.md                    # This README file
LICENSE                      # License file (e.g., MIT License)

Contributing
Contributions, bug reports, and feature requests are welcome. Please feel free to open issues or submit pull requests.

License
This project is licensed under the MIT License. See the LICENSE file for details.