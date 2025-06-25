# Meeting Attendance Report Web Application

## Overview
This is a frontend-only web application for processing Microsoft Teams meeting attendance reports in CSV format. It generates a summary table showing each participant's total attendance time per meeting and allows exporting the results to Excel.

## Features
- Upload and process multiple Teams attendance CSV files at once
- Validates file structure and required columns
- Calculates total attendance time for each participant, constrained by the organizer's presence window
- Applies a default minimum attendance threshold of 2 minutes
- Displays a summary table: rows = participants, columns = meetings (by date), cells = attendance time (minutes)
- Export the summary table to Excel format
- User-friendly error messages for invalid files or missing data
- No backend or persistent storage required; all processing is done in the browser

## Usage
1. Open `index.html` in your web browser.
2. Click **Upload CSV Files** and select one or more Teams attendance report CSV files.
3. Click **Process Files** to generate the summary table.
4. Review the attendance summary table.
5. Click **Export to Excel** to download the summary as an Excel file.

## CSV File Requirements
- Must include the following columns:
  - `Full Name`
  - `Email Address`
  - `Join Time`
  - `Leave Time`
  - `Duration (minutes)`
  - `Role`
- Each file must contain at least one row with `Role` set to `Organizer` (to define the valid meeting window).
- Meeting is identified by the date in the `Join Time` field.

## Error Handling
- If a file is missing required columns, an error message will specify which fields are missing.
- If organizer data is missing or invalid, the app will prompt the user to check the report.
- Corrupted or empty files will be reported with a clear error message.

## Technology Stack
- HTML, CSS, JavaScript (Vanilla)
- [SheetJS (xlsx)](https://sheetjs.com/) for Excel export

## License
This project is provided as a prototype for demonstration and educational purposes. No warranty or support is provided.
