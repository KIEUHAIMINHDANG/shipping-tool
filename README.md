==============================================================================
                 SHIPMENT CONSOLIDATOR PRO - USER GUIDE
==============================================================================

[1] OVERVIEW
------------
This tool automates the extraction of packing list data from weekly Excel files.
It scans a selected folder, finds shipment details (Item Name, Code, Qty, 
Carton Ranges, Weights), and consolidates them into a single "Master Database" 
Excel file.

[2] INSTALLATION (First Time Only)
----------------------------------
1. Install Python: https://www.python.org/downloads/
   * IMPORTANT: Check the box "Add Python to PATH" during installation.

2. Install Required Libraries:
   Open Command Prompt (cmd) or Terminal and run:
   pip install pandas openpyxl

[3] HOW TO RUN
--------------
1. Place the script (shipment_tool.py) in a folder on your computer.
2. Open Command Prompt/Terminal in that folder.
3. Run the command:
   python shipment_tool.py

4. A window will appear:
   - Click "Select Folder".
   - Choose the folder containing your weekly files (e.g., "Tuần 50...").
   - Click "START PROCESSING".

[4] INPUT FILE REQUIREMENTS
---------------------------
To ensure data is detected correctly, your Excel files must follow these rules:

A. Filenames
   - Must start with the word "Tuần" (e.g., "Tuần 50.2025.xlsm").
   - The tool extracts the Week Number automatically from this name.

B. Sheet Names
   - The tool ignores sheets named: "Mail", "mau_mui_tuan", "Sheet1".
   - All other sheets are treated as Invoices/Shipments.

C. Header Keywords (Case Insensitive)
   The tool looks for these specific words in the first 30 rows to find data:
   - Item Name:   "Tên sản phẩm" OR "Item Description"
   - Total Qty:   "Lượng xuất" OR "Total Qty" (REQUIRED)
   - Item Code:   "Mã số" OR "Item Code"
   - Range:       "Dải số thùng" OR "Carton No"
   - Weights:     "N.W", "G.W", "Total N.W", "Total G.W"

D. Weight Column Format
   - Separate Columns: N.W and G.W in their own columns is fine.
   - Merged Column: If N.W and G.W are in one cell, they must be separated 
     by a slash (e.g., "96.79 / 104.52").

[5] OUTPUT FORMAT
-----------------
The tool generates a file named "Master_Shipment_DB.xlsx" in the scanned folder.
The "Shipment Details" column uses the following format for each item:

   Item Name (Code) - QTY: X pcs - Y cartons [Ctn: Start-End] - Total pcs - N.W - G.W

   * Multiple items in one shipment are separated by a Line Break + "||".

[6] TROUBLESHOOTING
-------------------
> "The file is currently open" Warning:
  - You cannot save the database if "Master_Shipment_DB.xlsx" is open in Excel.
  - Close the Excel file and click "Retry" on the popup message.

> "No data found":
  - Check if your filename starts with "Tuần".
  - Open the file and check if the header row contains "Lượng xuất" or "Total Qty".
  - Ensure the data is within the first 100 rows.

> "Stop Words":
  - The tool stops reading a sheet when it sees: "TOTAL", "MÃ CÂN", "LÁI XE", 
    "SIGNATURE", or "GIÁ GỖ". Ensure these words are NOT used in item descriptions.

==============================================================================
