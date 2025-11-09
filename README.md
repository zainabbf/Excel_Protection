# Excel Password Protection Automation

This tool adds "Password to Open" protection, meaning users must enter a password before they can even view the file contents.

## 🚀 Features

- **Batch Processing**: Protect multiple Excel files at once
- **Open Protection**: Adds encryption requiring password before file opens
- **Progress Tracking**: Real-time progress and error reporting
- **Windows Compatible**: Designed for Windows systems with Excel installed

## 📋 Prerequisites

### System Requirements
- **Windows OS** (required for win32com)
- **Microsoft Excel** installed on your computer
- **Python 3.6+**

### Python Dependencies
Install required packages:
```bash
pip install pywin32
```

## 🛠 Installation

1. **Clone or download** the script files to your computer
2. **Install Python dependencies**:
   ```bash
   pip install pywin32
   ```
3. **Ensure Excel is installed** on your Windows system

## 🎯 Usage

### Basic Usage
1. **Run the script**:
   ```bash
   python password_protect_excel.py
   ```

2. **Follow the prompts**:
   - Enter the folder path containing your Excel files
   - Enter the desired password

3. **Wait for completion** - the script will process all Excel files and provide a summary

### Example
```bash
Enter folder path: C:\Users\YourName\Documents\ExcelFiles
Enter password to open: MySecurePassword123

Found 15 Excel files to protect
Processing: file1.xlsx
✓ Password protected: file1.xlsx
Processing: file2.xlsx
✓ Password protected: file2.xlsx
...
SUMMARY: 15 files protected, 0 files failed
```

## 🔧 Technical Details

### About win32com
**win32com** is a Python library that enables communication with Windows COM objects, specifically Microsoft Excel in this case. It provides:

- **Direct Excel Integration**: Uses the actual Excel application
- **Full Feature Access**: Can access all Excel functionality
- **Native Encryption**: Uses Excel's built-in encryption engine
- **Windows Exclusive**: Only works on Windows systems

### Why win32com is Required
Unlike other Excel libraries (like OpenPyXL), win32com can:
- Set "Password to Open" protection
- Encrypt the entire Excel file
- Use Excel's native encryption algorithms
- Create truly password-protected files that prompt for password on open

### Protection Levels
This script adds **"Password to Open"** protection, which:
- ✅ Encrypts the entire file
- ✅ Requires password before file opens
- ✅ Provides high security level
- ✅ Cannot be bypassed without password

## ⚠️ Important Notes

### Security Warnings
- **Remember your password** - there is NO recovery option
- **Keep backups** of important files before running the script
- **Test with sample files** first to ensure the process works as expected

### Limitations
- **Windows only** - requires Windows OS and Excel installation
- **Excel files only** - works with .xlsx and .xls files
- **File access** - ensure files aren't open in Excel during processing

## 🐛 Troubleshooting

### Common Issues

1. **"Cannot access file" errors**
   - Close all Excel instances before running
   - Ensure files aren't open in other programs
   - Check file permissions

2. **win32com not found**
   ```bash
   pip install pywin32
   ```

3. **Excel not installed**
   - Install Microsoft Excel
   - Ensure it's properly activated

4. **Permission errors**
   - Run script as Administrator
   - Check file/folder permissions

### Error Messages
- **"Excel application not found"**: Excel not installed
- **"Cannot access file"**: File is open or locked
- **"Permission denied"**: Insufficient file permissions

## 🔄 Removing Protection

To remove password protection later, you can modify the script or use Excel manually:

1. Open the protected file in Excel (with password)
2. Go to **File → Save As → Tools → General Options**
3. Delete the password fields
4. Save the file

## 📝 Script Details

### Supported File Types
- ✅ .xlsx (Excel Workbook)
- ✅ .xls (Excel 97-2003 Workbook)
- ✅ .xlsm (Excel Macro-Enabled Workbook)

### Processing Details
- Processes all subfolders recursively
- Creates temporary backups during processing
- Provides detailed progress reporting
- Handles errors gracefully without stopping entire process

## ⚡ Quick Start

1. **Install**: `pip install pywin32`
2. **Run**: `python password_protect_excel.py`
3. **Enter**: Folder path and password
4. **Wait**: For processing to complete

---

**Need Help?** Check the troubleshooting section or create an issue in the project repository.

## 🎉 Success Message
When successful, you'll see:
```
SUMMARY: 15 files protected, 0 files failed
```

All your Excel files are now encrypted and will require a password to open! 🔒

---