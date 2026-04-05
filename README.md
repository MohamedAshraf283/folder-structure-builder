# 📁 Folder Structure Builder

A professional Python GUI application for building folder structures and optionally copying files/folders in batch mode.

Built with Tkinter and packaged using PyInstaller.

---

## 🚀 Overview

**Folder Structure Builder** is a powerful desktop tool that allows you to:

- Create nested folder structures in batch
- Optionally copy files or entire folder contents
- Prevent overwriting existing files
- Perform dry-run simulations before execution
- Automatically rollback partial folder copies on failure
- Enforce path validation and Windows reserved name protection
- Block symlink copying for safety

This tool is designed for developers, automation workflows, and structured project setups.

---

## ✨ Features

- 📂 Batch folder creation
- 📄 Optional file or folder copy
- 🔁 Rollback support for partial copy failures
- 🚫 No-overwrite protection
- 🧪 Dry Run mode (simulation without filesystem changes)
- 🔐 Symlink protection
- 🪟 Windows reserved names validation
- 🖥️ Clean GUI (Tkinter-based)
- 📦 One-file executable build support

---

## 🖥️ Download

You can download the latest compiled executable from the repository or the Releases section.

File: FolderStructureBuilder.exe


---

## 🛠️ Build From Source

### Requirements

- Python 3.11+ (3.12 recommended)
- PyInstaller

### Install Dependencies

```bash
pip install pyinstaller
pyinstaller --noconfirm --onefile --windowed --clean --noupx --name FolderStructureBuilder FolderStructureBuilder.py
```

🔒 Safety Design

The application enforces:

No overwriting of existing files or directories
Automatic rollback on partial copy failure
Path traversal prevention (..)
Windows reserved name validation
Optional symlink blocking

This ensures predictable and safe filesystem operations.

🧭 Main Interface Overview

The application interface is divided into several sections:

Destination (Output Root)
Created Tasks (Jobs List)
Settings
Action Buttons
Log Panel

Each section plays a specific role.

📂 1. Destination (Output Root)

This is the base directory where all folder structures will be created.

How to Use:
Click Browse
Select an existing folder
This folder becomes the root under which all jobs will be executed

Example:

If Output Root is:

D:\Projects

And your structure path is:

ClientA\Reports\2026

The final created path will be:

D:\Projects\ClientA\Reports\2026

⚠ Output Root must already exist.

🧱 2. Created Tasks (Jobs)

Each Job represents one folder structure creation task, optionally with copy behavior.

Add Job

Click Add Job

You will see two fields:

Structure Path (Required)

Enter a relative folder structure:

ClientA\Reports\2026

Rules:

Must not contain invalid Windows characters
Cannot contain reserved names (CON, NUL, COM1, etc.)
.. is blocked if traversal protection is enabled
Absolute paths are optional (based on settings)
Copy Source (Optional)

You may choose:

A file
A folder

If you choose a folder:

Only its CONTENTS are copied
The container folder itself is NOT copied

Example:

If Copy Source is:

D:\Templates\BaseProject

And inside it:

config.json
assets/

These will be copied into the target directory directly.

Edit Job
Select a job
Click Edit
Modify structure or copy source
Remove Job
Select job
Click Remove
Reorder Jobs

Use:

Move Up
Move Down

Jobs execute in order from top to bottom.

⚙️ 3. Settings Explained
🔹 Dry Run Mode

When enabled:

No folders are created
No files are copied
A full execution simulation is displayed

Use this to verify operations before executing.

🔹 Allow Absolute Structure Path

If enabled:
You can enter:

C:\xampp\htdocs\project

The drive letter will be stripped and applied under Output Root.

Example:

Output Root: D:\Deploy
Structure Path: C:\xampp\htdocs
Result: D:\Deploy\xampp\htdocs
🔹 Forbid '..' Traversal

Prevents using:

..\AnotherFolder

This protects against directory escape.

Recommended: Keep enabled.

🔹 Disallow Symlinks

Blocks copying:

Symlink files
Symlink folders
Symlinks inside folders (recursive detection)

This ensures safe and predictable file operations.

▶ 4. Action Buttons
🔹 Run All Jobs

Executes all jobs sequentially.

Behavior:

Creates structure
Copies content (if specified)
Prevents overwrite
Rolls back partial copies on failure
🔹 Preview Plan

Simulates execution and opens a detailed plan window.

Shows:

Job number
Type (DIR / FILE)
Full path
Notes

Recommended before real execution.

🔹 Clear Log

Clears execution logs in the lower panel.

📜 5. Log Panel

Displays:

Execution steps
Created directories
Copy operations
Errors
Rollback actions
Summary statistics

Always check this panel if something fails.

🔁 Execution Flow (What Happens Internally)

For each job:

Validate structure path
Compute full target path
Ensure directory exists
Validate copy source (if any)
Check overwrite rules
Perform copy
Rollback if failure occurs
🛑 Error Handling

Common errors:

Destination already exists
Permission denied
Invalid folder name
Reserved Windows name
Symlink detected
Copy source does not exist

All errors are logged clearly.

🧪 Recommended Workflow
Select Output Root
Add Jobs
Enable Dry Run
Click Preview Plan
Review paths
Disable Dry Run
Click Run All Jobs
💡 Best Practices
Always test with Dry Run first
Keep overwrite protection enabled
Avoid using system directories as Output Root
Use small test folders before large operations
Keep symlink protection enabled unless necessary
🏗 Example Scenario

Goal:
Create project structure for multiple clients.

Jobs:

ClientA\Docs
ClientA\Reports
ClientB\Docs
ClientB\Reports

Output Root:

D:\Clients

Click Run → All structures created automatically.

🧩 Advanced Usage

You can:

Chain multiple jobs
Use folder templates
Use as a deployment preprocessor
Combine with version-controlled templates   

👨‍💻 Author

Mohamed Ashraf
