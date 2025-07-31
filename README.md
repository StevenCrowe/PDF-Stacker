# PDF Merger Tool - User Guide

## WHAT THIS TOOL DOES
The PDF Merger Tool automatically combines multiple PDF files into one document based on part numbers listed in your Excel file. It searches your computer for PDFs that match the part numbers and merges them in the exact order specified in your Excel spreadsheet.

## PACKAGE CONTENTS
This package contains 3 files:
• SETUP_AND_RUN.bat - The installer and launcher (double-click this to start)
• pdf_merger.py - The main program (don't click this directly)

## SYSTEM REQUIREMENTS
• Windows PC (Windows 7 or newer)
• Excel file with a "Part Number" column
• PDF files named to match your part numbers (e.g., "12345.pdf" for part number "12345")
• Internet connection (for initial setup only)

## INSTALLATION - SUPER EASY!
1. Extract all files from the ZIP to a folder (like your Desktop)
2. Double-click "SETUP_AND_RUN.bat"
3. Press any key when prompted
4. If Python isn't installed, it will guide you through the installation
5. The tool will open automatically when setup is complete

That's it! Setup only happens once.

## HOW TO USE THE TOOL

### Step 1: Select Your Excel File
• Click "Browse Excel File"
• Choose your Excel file that contains part numbers
• The tool will automatically find the "Part Number" column

### Step 2: Add Search Folders
• Click "Add Folder" to specify where your PDFs are located
• You can add multiple folders
• Use "Add Entire Drive" to search an entire drive (C:, D:, etc.)
• The tool will search all subfolders automatically

### Step 3: Choose Output Location
• Click "Browse Output Folder"
• Select where you want the final merged PDF saved
• The merged file will be named after your Excel file

### Step 4: Start Merging
• Click "Start Merging PDFs"
• Watch the progress in the status window
• The tool will show you which files it finds and any missing ones

## EXCEL FILE REQUIREMENTS
• Must have a column named "Part Number" (exact spelling)
• Part numbers should be in the order you want them merged
• Empty cells will be skipped automatically

## PDF FILE NAMING
Your PDF files should be named exactly like the part numbers:
• Part number "12345" → PDF file "12345.pdf"
• Part number "ABC-123" → PDF file "ABC-123.pdf"
• Capitalization doesn't matter

## WHAT HAPPENS DURING MERGING
1. Tool reads your Excel file for part numbers
2. Searches specified folders for matching PDF files
3. Shows you what it finds (and what's missing)
4. Merges PDFs in the exact order from your Excel file
5. Saves final PDF with your Excel file's name
6. Shows completion message with file location

## TROUBLESHOOTING

### "Python not found" Error
• The installer will automatically open the Python download page
• Download Python and make sure to check "Add Python to PATH"
• Run SETUP_AND_RUN.bat again after installing Python

### PDFs Not Found
• Check that PDF filenames exactly match part numbers
• Make sure you've added the correct search folders
• Try adding more folders or entire drives
• Check for typos in part numbers or filenames

### Excel File Errors
• Make sure your column is named exactly "Part Number"
• Save Excel file before using the tool
• Close Excel file before running the merger

### Tool Won't Start
• Right-click SETUP_AND_RUN.bat and choose "Run as administrator"
• Make sure both files (SETUP_AND_RUN.bat and pdf_merger.py) are in the same folder
• Check that your antivirus isn't blocking the files

### Slow Performance
• Searching entire drives can take time
• Use specific folders instead of entire drives when possible
• Close other programs while the tool is running

## TIPS FOR BEST RESULTS
• Close your Excel file before running the merger
• Use specific folders instead of searching entire drives
• Make sure PDF filenames match part numbers exactly
• Test with a small Excel file first to verify everything works

## AFTER FIRST SETUP
Once setup is complete, you can:
• Just double-click SETUP_AND_RUN.bat anytime to use the tool
• No need to install anything again
• The tool remembers all settings

## EXAMPLE WORKFLOW
1. You have Excel file "Project_ABC_Parts.xlsx" with part numbers in column "Part Number"
2. Your PDFs are scattered across folders: C:\Documents\PDFs\, D:\Archive\, etc.
3. Run SETUP_AND_RUN.bat
4. Select "Project_ABC_Parts.xlsx"
5. Add both folders (C:\Documents\PDFs\ and D:\Archive\)
6. Choose Desktop as output location
7. Click "Start Merging PDFs"
8. Result: "Project_ABC_Parts.pdf" created on Desktop with all PDFs merged in order

## SUPPORT
If you continue to have problems:
• Try running as Administrator
• Check that your antivirus isn't blocking the files
• Contact your IT administrator
• Make sure you have internet connection for initial setup

## VERSION INFORMATION
This tool requires:
• Python 3.6 or newer (installed automatically)
• pandas library (for Excel files)
• PyPDF2 library (for PDF merging)
• openpyxl library (for Excel support)

All libraries are installed automatically during setup.

---
Created for easy PDF merging based on Excel part number lists.
No technical knowledge required - just follow the simple steps above!
