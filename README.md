# export-as

Sheet Exporter ~ Google Apps Script project that reads data from a Google Sheet and creates a new Google Sheet and an Excel (.xlsx) file in your Google Drive.  

## Setup Instructions

1. Open any Google Sheet (create a new one if needed).
2. Go to Extensions > App Script.
3. Delete any code in `Code.gs` and paste the contents of `Code.gs` from this repo.
4. Save the code. You can simply press `Ctrl+S` to save.

## How to Run

1. In the Apps Script editor, select `writeSheetFile` from the function dropdown, then click `Run` to execute it.
2. The first time you run it, Google will ask for authorization. Accept the permissions.
3. Once finished, the data will be exported to a new Google Sheet and an Excel file in your Google Drive.

## Configuration

- In `Code.gs` in Apps Scripts, locate the following line and update it:

```javascript
  // 👇 Modify this to the sheet you want to export from
  const sourceSheetName = "Sheet1"; 
```

- You can also modify the base name of the exported files:

```javascript
  // 👇 Modify this to customize the base name of the exported files
  const baseFileName = "Exported_Data"; 
```

## Output

- Data from the source sheet will be written to the new Google Sheet.
- An Excel `(.xlsx)` file will also be created in your Google Drive with the same content.
- Timestamps in filenames ensure uniqueness.

## Limitations

- The script exports the entire content of the source sheet. Partial ranges are not configurable in this version.
- It does not interact with external websites or perform web scraping.
- For extracting tables from webpages, refer to the [scraper-as](https://github.com/erujs/scraper-as) project.  
These two projects work best together: `scraper-as` can populate a sheet with web table data, which `export-as` can then export to new sheets or Excel files.
- Requires the source sheet to exist and contain data.

## Notes

- This project is an upgrade from the previous `vba-file` repo. Since VBA is approaching the end of its mainstream support and requires Excel for execution.
- This Google Apps Script version provides a more stable, modern, and free solution for exporting data from Google Sheets.
- This project is designed to be simple, lightweight, and entirely free to use with Google Sheets.
- We retain `Module.bas` in the repository for anyone who wishes to continue using the legacy VBA version.

## 🧾 License
MIT — do whatever you want with it.

✨ Happy coding!
If you find this project useful, a ⭐ on the repo is always appreciated!