# Workflow to generate bridge inspection record from scratch

## I developed a two-part system to generate structured inspection reports in Excel.

## Part 1 – Template Builder (exceltemplateWSP.py)
- I designed a clean, formatted Excel template for bridge inspection reports. The script defines consistent styles, borders, and column widths, and organizes sections for the project title, inspection details, notes, condition states, descriptions, and photographs. Each section includes labeled areas for data entry. I also created a hidden “_anchors” sheet that maps each field name (like BIN, Team Leader, Condition Note, Photo Filename) to a specific cell location, allowing the next script to populate the template automatically.

## Part 2 – Report Generator (excel_to_formatted_excel.py)
- I wrote this script to read inspection data from an input Excel file and fill the template row by row. It normalizes column names, maps them to the correct fields, and generates one completed report sheet per record by copying the template.The script also processes photo paths and filenames, locates the corresponding images (even when extensions differ), and places up to two photos in the report at fixed positions. Each image is resized proportionally to fit neatly within its section. After all reports are created, it removes the template and anchor sheets and saves everything in a single Excel file.

## In summary:
- I built an automated workflow that transforms raw inspection data and photo references into polished, professional Excel reports ,one sheet per record ,with all text and images placed precisely according to the defined layout.
