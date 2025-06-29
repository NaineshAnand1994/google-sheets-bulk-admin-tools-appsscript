# 🛠️ Google Sheets Bulk Admin Tools (Google Apps Script)

A professional Google Apps Script add-on for Google Sheets that automates common admin tasks—such as clearing data, formatting headers, duplicating template sheets, and archiving data rows.  

It reads configuration from a Settings sheet for repeatable automation, **but also supports interactive prompts when settings are blank**—giving you full flexibility with no code edits needed.

This project demonstrates how to build real-world, business-focused Google Workspace automation with dynamic, user-configurable behavior.

---

## 🚀 Features
✅ Custom menu added automatically on open  
✅ Clear all data except header  
✅ Apply standard header formatting  
✅ Duplicate a template sheet with timestamp in name  
✅ Archive current sheet data to an archive sheet  
✅ Protect header row to prevent edits  
✅ Hybrid mode:
- Reads from Settings sheet for automation
- Prompts user interactively if settings are blank

---

## ⚙️ How It Works
- On open, the script reads settings from a **Settings** tab in your Google Sheet.
- It dynamically builds a custom menu with these admin actions.
- Each menu item performs a task using parameters from the Settings sheet.
- **If any parameter is blank in Settings, the script will prompt the user interactively**—so you don't have to edit the sheet every time.

✅ No need to touch the code to customize!  
✅ Change parameters in the Settings sheet or choose at runtime.

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
- `MenuName`: Name of the custom menu shown in the sheet.  
- `TemplateSheetName`: Sheet to copy when "Duplicate Template" is used. Leave blank to choose interactively.  
- `ArchiveSheetName`: Sheet to store archived data. Leave blank to choose interactively.  
- `ProtectRangeRows`: Number of header rows to protect from editing. Leave blank to enter at runtime.

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
  - Wipes all data except header rows.  
- **Apply Header Formatting**  
  - Makes header row(s) bold with background color.  
- **Duplicate Template**  
  - Copies a named template sheet (or prompts for one) with timestamp in name.  
- **Archive Data**  
  - Appends active sheet data to an archive sheet (or prompts for one).  
- **Protect Header Row**  
  - Locks header rows against editing (reads count from Settings or prompts user).

---

## 💻 Code
See [Code.gs](Code.gs) for the full Google Apps Script code.

---

## 💡 Ideas for Extensions
⭐ Add email notifications on archiving.  
⭐ Customize formatting via Settings.  
⭐ Add data validation before archiving.  
⭐ Support multiple templates or archives.  
⭐ Add InteractiveMode toggle in Settings.

---

## 🪪 License
MIT License

---

## ❤️ Author
[Your Name or GitHub Profile Link]
