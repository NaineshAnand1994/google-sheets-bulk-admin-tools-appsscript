# 📂 Google Drive File Organizer and Sharer (Google Apps Script)

A professional Google Apps Script solution that automates organizing your Google Drive by moving or copying files between folders, renaming them with advanced rules, filtering by type, and automatically setting sharing permissions for different users.  

All configuration is controlled via a single Settings sheet in Google Sheets—making it easy to customize without editing any code.  

---

## 🚀 Features
✅ Move or copy files from Source to Destination folder  
✅ Supports renaming: prefix, suffix, timestamp, replace text  
✅ Optionally creates date-stamped subfolders for archiving  
✅ Filters files by MIME type or high-level categories (Documents, Images, Videos)  
✅ Optional Dry Run / Preview mode to see planned changes without moving files  
✅ Batch limit (MaxFiles) to control how many files are processed at once  
✅ Optional delay between moves to avoid API limits  
✅ Interactive prompts if settings are blank or missing  
✅ Detailed logging of all operations to a Google Sheet tab  
✅ Error logging with Admin email notification for failures  
✅ Fine-grained sharing rules per user and role (reader, writer, commenter)  

---

## ⚙️ How It Works
- Reads all configuration from a **Settings** tab in your linked Google Sheet.  
- Supports both fully automated and interactive usage:
  - If settings are filled in → runs automatically.  
  - If settings are blank → prompts user for input at runtime.  
- Moves or copies files from a source folder to a destination folder.  
- Optionally renames files according to user-defined rules.  
- Filters files by type, MIME category, or filename contains text.  
- Sets sharing permissions automatically for different users with different roles.  
- Logs all operations and errors to a Google Sheet log.  

---

## 🗂️ 📌 Example Settings Sheet

Create a tab in your Google Sheet called **Settings** with this structure:

### **General Parameters Section**

| Parameter               | Value                                         |
|---------------------------|-----------------------------------------------|
| SourceFolderID            | 123ABC456DEF                                |
| DestinationFolderID       | 789XYZ987ZYX                                |
| RenameMode                | prefix                                       |
| RenameText                | Processed_                                   |
| ReplaceFrom               | oldtext                                      |
| ReplaceTo                 | newtext                                      |
| MaxFiles                  | 10                                           |
| AllowedMimeTypes          | application/pdf,image/png                    |
| AllowedCategory           | Documents                                    |
| DryRunMode                | FALSE                                        |
| MoveMode                  | move                                         |
| UseDateSubfolder          | TRUE                                         |
| DelaySeconds              | 2                                            |
| AdminEmail                | admin@example.com                            |
| MoveIfContainsText        | Invoice                                      |

*(Optional blank row for clarity)*

### **Sharing Rules Section**

| ShareType     | Emails                                         |
|----------------|-----------------------------------------------|
| reader         | reader1@example.com,reader2@example.com       |
| writer         | writer1@example.com                           |
| commenter      | commenter1@example.com,commenter2@example.com |

✅ **How to use this:**
- Fill in any parameters you want fixed for automation.
- Leave fields blank to have the script prompt you at runtime.
- Add or remove sharing rules rows as needed.

---

## 🛠️ Setup Instructions

1️⃣ Create a Google Sheet with a tab named **Settings**.  
2️⃣ Fill in the example table above with your actual values.  
3️⃣ Open **Apps Script Editor**:  
   - Paste the code from `Code.gs`.  
   - Link it to your Google Sheet.  
4️⃣ Save and authorize the script.  
5️⃣ Run the **organizeFiles()** function manually or via a trigger.  
6️⃣ Check the Logs tab (if enabled) for details of what was processed.

---

## 🧩 Included Features in Detail

### ✅ 1. File Moving or Copying
- Choose **move** or **copy** mode in settings.
- Supports creating a date-stamped subfolder automatically.

### ✅ 2. Advanced Renaming Options
- Prefix
- Suffix
- Timestamp
- Replace text in filenames

### ✅ 3. File Filtering
- MIME type filtering (e.g. only PDFs, Images)
- High-level categories (Documents, Images, Videos)
- Filename contains specific text

### ✅ 4. Sharing Rules
- Define emails for each permission level (reader, writer, commenter).
- The script automatically assigns these roles to all processed files.

### ✅ 5. Dry Run / Preview Mode
- See which files would be processed without actually moving them.
- Safer for testing new settings.

### ✅ 6. Batch Limits and Delays
- Control how many files are processed per run.
- Add optional delays between operations to avoid API limits.

### ✅ 7. Logging
- Appends details of each run to a **Logs** tab in your Sheet.
- Tracks file name, new name, source/destination folder, status, errors.
- Logs errors with timestamp and details.
- Optionally emails error reports to AdminEmail.

---

## 💡 Example Use Cases
✅ Organize bulk uploads into year/month subfolders  
✅ Add "Processed_" prefix to filenames  
✅ Filter and move only invoices or contracts  
✅ Ensure all files shared with your team automatically  
✅ Preview changes before running in production  

---

## 💻 Code
See [Code.gs](Code.gs) for the full Google Apps Script implementation.

---

## 💡 Ideas for Extensions
⭐ Automatic subfolder creation based on metadata (e.g. client name)  
⭐ Integration with Google Forms for custom triggers  
⭐ Custom approval workflows before moving files  
⭐ Scheduled runs via time-based triggers  

---

## 🪪 License
MIT License

---

## ❤️ Author
[Your Name or GitHub Profile Link]
