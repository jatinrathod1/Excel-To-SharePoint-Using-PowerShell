# Excel to SharePoint List Sync using PowerShell 🚀

This PowerShell script automates the process of importing data from a **local Excel file** 📊 into a **SharePoint Online List** 🌐. It reads data using the `Import-Excel` module, resolves **lookup columns**, and adds or updates list items in SharePoint via `PnP PowerShell`.

## Features ✨
✅ Reads data from an **Excel file (.xlsx)** stored locally.  
✅ Maps **lookup columns** to SharePoint List IDs dynamically.  
✅ Supports **updating existing records** and adding new ones.  
✅ Uses **PnP PowerShell** for seamless SharePoint integration.  
✅ Handles **date formats and missing data errors**.  

## Requirements ⚙️
📌 **PnP PowerShell** module (`Install-Module PnP.PowerShell`).  
📌 **ImportExcel** PowerShell module (`Install-Module ImportExcel`).  
📌 **SharePoint Online access** with necessary permissions.  

## Usage 🛠️
1. Update the **Excel file path** and **SharePoint List name** in the script.  
2. Run the script in **PowerShell** after logging in (`Connect-PnPOnline`).  
3. Data from Excel will be **synced to SharePoint** automatically.  

## Installation 📥
1. Install required PowerShell modules:
   ```powershell
   Install-Module PnP.PowerShell
   Install-Module ImportExcel
   ```
2. Connect to SharePoint Online:
   ```powershell
   Connect-PnPOnline -Url "https://yourtenant.sharepoint.com/sites/yoursite" -UseWebLogin
   ```
3. Run the script:
   ```powershell
   .\ExcelToSharePoint.ps1
   ```

## Example Code 📝
```powershell
# Import Excel data
$ExcelFilePath = "C:\Users\User\Documents\data.xlsx"
$ExcelData = Import-Excel -Path $ExcelFilePath

# Define SharePoint List
$ListName = "EmployeeList"

foreach ($row in $ExcelData) {
    $Name = $row.Name
    $JobRole = $row.Job_Role
    $Department = $row.Department
    $Payment = $row.Payment
    
    # Lookup column handling
    $DepartmentId = (Get-PnPListItem -List "DepartmentList" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Department</Value></Eq></Where></Query></View>").FieldValues["ID"]
    $PaymentId = (Get-PnPListItem -List "PaymentList" -Query "<View><Query><Where><Eq><FieldRef Name='Title'/><Value Type='Text'>$Payment</Value></Eq></Where></Query></View>").FieldValues["ID"]
    
    # Add or update SharePoint list item
    Add-PnPListItem -List $ListName -Values @{
        "Title" = $Name
        "Job_Role" = $JobRole
        "Department" = $DepartmentId
        "Payment" = $PaymentId
    }
}
```

## GitHub Tags🔍
- `PowerShell SharePoint Automation`  
- `Excel to SharePoint List Upload`  
- `PnP PowerShell Import Data`  
- `Sync Excel with SharePoint`  
- `SharePoint List Bulk Upload`  

---

