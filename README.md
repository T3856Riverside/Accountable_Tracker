# Employee Occurrence Tracker

A browser-based employee occurrence tracker styled with Target-inspired red colors.

## Features

- Upload `.xlsx`, `.xls`, `.xlsm`, or `.csv` occurrence data.
- Reads the `Input Data` worksheet when available.
- Search by TM name, DC Key, or employee number.
- View employee details, occurrence counts, total hours, and occurrence history.
- Filter occurrence history by `TYPE`.
- Export the selected TM's report to PDF, including the currently filtered occurrence history.
- Open a separate Top 10 Accountables page ranked by occurrence-record count.
- Navigate between pages without re-uploading during the same browser session.
- Spreadsheet data is processed in the browser.

## Files

- `index.html`: Employee search page and report controls.
- `top10.html`: Top 10 Accountables leaderboard page.
- `styles.css`: Shared styling and responsive layout.
- `app.js`: Spreadsheet loading, searching, type filtering, and PDF report generation.
- `top10.js`: Top 10 leaderboard calculations and rendering.
- `assets/`: Target-themed mascot and reference images.

## Run locally

For best results, run the folder through a local web server or publish it to GitHub Pages. Open `index.html`, upload the spreadsheet, and use the navigation bar to visit the Top 10 page.

## Notes

- Parsed spreadsheet records are saved in `sessionStorage` so the two pages can share the uploaded data during the current browser session. Use **Remove loaded data** to clear the session data.
- PDF generation uses jsPDF and AutoTable from CDN-hosted libraries, so an internet connection is needed to load them.
- Each qualifying row in the input data is counted as one accountable record in the Top 10 ranking. If the business rule should count only specific `TYPE` values, adjust the ranking logic in `top10.js`.
