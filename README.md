# Garmin Data Processor

This project processes CSV data files exported from Garmin (or similar) devices, and compiles them into a single Excel file with a summary tab and one tab per CSV file.

## Features
- Processes all CSV files in the `data_files` folder
- Skips metadata and summary rows in each CSV
- Each CSV is added as its own tab in the output Excel file
- The summary tab (first tab) includes:
  - File name
  - Date and time (from the Date field in each CSV)
  - Number of shots (rows with numeric `#`)
  - Average Speed (from summary row in CSV)
  - Standard Deviation (from summary row in CSV)
  - Spread (from summary row in CSV)
- Summary rows are sorted by date and time in ascending order
- Uses [exceljs](https://www.npmjs.com/package/exceljs) for Excel file creation and formatting

## Prerequisites
- [Node.js](https://nodejs.org/) (v18+ recommended)
- [pnpm](https://pnpm.io/)

## Setup
1. Install dependencies:
   ```sh
   pnpm install
   ```
2. Place your CSV files in the `data_files` folder.

## Running the Processor
To process the CSV files and generate the Excel summary:

```sh
pnpm exec ts-node index.ts
```

The output file will be created at `output/summary.xlsx`.

## Project Structure
- `index.ts` - Main script for processing CSVs and generating the Excel file
- `data_files/` - Place your CSV files here
- `output/` - The generated Excel file will be saved here

## Notes
- The script expects each CSV to have a metadata line, a header line, shot data, and summary rows (see sample files in `data_files/`).
- The `.gitignore` file excludes dependencies, output, and common IDE/system files from version control.
