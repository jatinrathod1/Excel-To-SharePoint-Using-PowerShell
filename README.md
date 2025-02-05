Excel to SharePoint List Sync using PowerShell 🚀
This PowerShell script automates the process of importing data from a local Excel file 📊 into a SharePoint Online List 🌐. It reads data using the Import-Excel module, resolves lookup columns, and adds or updates list items in SharePoint via PnP PowerShell.

Features ✨
✅ Reads data from an Excel file (.xlsx) stored locally.
✅ Maps lookup columns to SharePoint List IDs dynamically.
✅ Supports updating existing records and adding new ones.
✅ Uses PnP PowerShell for seamless SharePoint integration.
✅ Handles date formats and missing data errors.

Requirements ⚙️
📌 PnP PowerShell module (Install-Module PnP.PowerShell).
📌 ImportExcel PowerShell module (Install-Module ImportExcel).
📌 SharePoint Online access with necessary permissions.

Usage 🛠️
1️⃣ Update the Excel file path and SharePoint List name in the script.
2️⃣ Run the script in PowerShell after logging in (Connect-PnPOnline).
3️⃣ Data from Excel will be synced to SharePoint automatically.
