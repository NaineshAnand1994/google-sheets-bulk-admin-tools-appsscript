# 🛠️ Google Sheets Bulk Admin Tools (Google Apps Script)

A professional Google Apps Script add-on for Google Sheets that automates common admin tasks—such as clearing data, formatting headers, duplicating template sheets, and archiving data rows.  

All configuration is controlled via a Settings sheet, making the add-on easy to customize without editing any code.

This project demonstrates how to build real-world, business-focused Google Workspace automation with dynamic, user-configurable behavior.

---

## 🚀 Features
✅ Custom menu added automatically on open  
✅ Clear all data except header  
✅ Apply standard header formatting  
✅ Duplicate a template sheet with a new name  
✅ Archive current sheet data to an archive sheet  
✅ Protect header row to prevent edits  
✅ All settings are controlled from a Google Sheet

---

## ⚙️ How It Works
- On open, the script reads settings from a **Settings** tab in your Google Sheet.
- It dynamically builds a custom menu with these actions.
- Each menu item performs a useful admin task, using parameters from the Settings sheet.

✅ No need to touch the code to customize!  
✅ Change parameters in the Settings sheet anytime.

---

## 🗂️ 📌 Example Settings Sheet

Create a tab in your Google Sheet called **Settings** with this structure:

| Parameter            | Value                    |
|-----------------------|-------------------------|
| MenuName              | Bulk Admin Tools        |
| TemplateSheetName     | InvoiceTemplate         |
| ArchiveSheetName      | Archive                 |
| ProtectRangeRows      | 1                       |

✅ **What each setting does:**
- `MenuName`: The name of the custom menu shown in the sheet.  
- `TemplateSheetName`: Name of the sheet you want to duplicate when "Duplicate Template" is used.  
- `ArchiveSheetName`: Name of the sheet to store archived data.  
- `ProtectRangeRows`: Number of header rows to protect from editing.  

---

## 🛠️ Setup Instructions

1️⃣ Create a Google Sheet and add a tab named **Settings**.  
2️⃣ Fill in the table above with your desired values.  
3️⃣ Open **Apps Script Editor**:  
   - Paste the code from `Code.gs`.  
   - Link it to your Google Sheet.  
4️⃣ Save the project and close the editor.  
5️⃣ Refresh your Google Sheet to see the custom menu.  
6️⃣ Click any menu action to run it!

---

## 🧩 Included Menu Actions
- **Clear Data**  
  - Clears all data except the header row.  
- **Apply Header Formatting**  
  - Makes header row bold with background color.  
- **Duplicate Template**  
  - Creates a copy of the template sheet with timestamp in name.  
- **Archive Data**  
  - Appends all current sheet data to an Archive sheet.  
- **Protect Header Row**  
  - Adds protected range to header rows to prevent edits.

---

## 💻 Code
See [Code.gs](Code.gs) for the full Google Apps Script code.

---

## 💡 Ideas for Extensions
⭐ Add custom email notifications when archiving.  
⭐ Make formatting fully customizable via Settings.  
⭐ Add data validation before archiving.  
⭐ Support multiple templates or archives.

---

## 🪪 License
MIT License

---

## ❤️ Author
[Your Name or GitHub Profile Link]
