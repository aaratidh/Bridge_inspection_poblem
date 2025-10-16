# What this code does 

It take inspection data (from Excel) and automatically make a formatted report with photos in another Excel file.

## Step 1: Set Up Folder
- Create a folder anywhere (like Desktop) named: "excel_to_excel"


## Inside it, put these files:
```bash
excel_to_excel/
├─ excel_to_formatted_excel.py   ← (main program you run)
├─ exceltemplateWSP.py           ← (builds the inspection template)
├─ inputexcelfile.xlsx           ← input data file
├─ photo/                        ← folder containing the pictures
│  ├─ AA_113_2933.jpg
│  └─ BB_001_0002.jpg

```

## Step 2: Install Python (if not already)
- Go to https://www.python.org/downloads
- Click Download for Windows
- During install, “Add Python to PATH”
- After install, open Command Prompt and type:
- python --version
- If it shows a version (like Python 3.11.5), it’s ready.

## Step 3: Install the Required Tools
- In Command Prompt, go into your folder:
- cd Desktop\excel_to_excel
- Then install everything you need: "pip install openpyxl pandas pillow"

## Step 4: Prepare the Excel Files
- Your inputexcelfile.xlsx should have columns like:
- BIN | Inspection Date | ... | Photo Filename | Photo Path
- Example:
  1065318 | 2022-08-31 | ... | AA_113_2933 | C:\Users\Hp\Desktop\excel_to_excel\photo

Your photo folder must have the matching photos.
## Step 5: Run the Script
- Still in the same folder, run:
- python excel_templateWSP.py
- it will geberate Template2.xls file
- then run python excel_to_formatted_excel.py

You’ll see something like:

Working directory: C:\Users\Hp\Desktop\excel_to_excel
[PHOTO] Row strictly resolved -> ['C:\\Users\\Hp\\Desktop\\excel_to_excel\\photo\\AA_113_2933.jpg']
[PHOTO] Placed ... at E27
Report generated: inspection_reports.xlsx

## Step 6: Open the Result
- Open the new file:

"inspection_reports.xlsx"

- Each inspection record → its own sheet
- Text fields filled in
- Photos automatically added in the lower photo box

