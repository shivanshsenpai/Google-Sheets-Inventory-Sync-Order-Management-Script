# Google-Sheets-Inventory-Sync-Order-Management-Script
This Google Apps Script automates inventory management across multiple sheets. It syncs stock availability between an Inventory sheet, a Matching Table, and Order sheets. Features include automatic image URL rendering, cross-spreadsheet SKU lookups, and conditional formatting for low stock. Replace the placeholder IDs to safely deploy in your works

# Google Sheets Inventory Sync

A Google Apps Script designed to synchronize product inventory, pricing, and stock status across multiple tracking sheets. 

## Features
* **Cross-Sheet Syncing:** Automatically pulls data from a primary `Inventory` sheet and a `MatchingTable` to update stock levels on external sheets (`All Orders` and `New_Orders`).
* **Automated Image Rendering:** Converts image URLs pasted into the `IMAGE_URL` column directly into Google Sheets `=IMAGE()` formulas.
* **Intelligent Stock Alerts:** Highlights rows with "Out of stock" status and calculates exactly how many units are needed to fulfill an order.
* **Error Handling:** Identifies missing Vendor SKUs, Portal SKUs, and general lookup errors, placing them in a designated "What to fix" column.

## Setup Instructions

1. **Open your Google Sheet:** Navigate to `Extensions` > `Apps Script`.
2. **Paste the Code:** Replace any existing code with the script provided in this repository.
3. **Set your Spreadsheet ID:** Find the `var ORDER_SPREADSHEET_ID` variable at the top of the file. Replace `'YOUR_SPREADSHEET_ID_HERE'` with the actual ID of your target tracking spreadsheet (found in the URL between `/d/` and `/edit`).
4. **Save the File:** Click the floppy disk icon or press `Ctrl + S` / `Cmd + S`.
5. **Reload your Sheet:** Go back to your Google Sheet and refresh the page. A new menu item called **Inventory Sync** will appear.

## Usage

* **Install the Trigger:** Click `Inventory Sync` > `Install edit sync trigger`. This will prompt you to authorize the script. Once authorized, it establishes a background trigger that detects inventory updates and syncs them automatically.
* **Manual Refreshes:** If you ever need to forcefully sync the data without waiting for an edit, use the manual refresh options available in the custom `Inventory Sync` menu.
