console.log('Garmin Data Processor initialized');

import * as fs from 'fs';
import * as path from 'path';
import { parse } from 'csv-parse/sync';
import * as XLSX from 'xlsx';

const dataDir = path.join(__dirname, 'data_files');
const outputDir = path.join(__dirname, 'output');
const outputFile = path.join(outputDir, 'processed.xlsx');

// Ensure output directory exists
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir);
}

const files = fs.readdirSync(dataDir).filter(f => f.endsWith('.csv'));
const workbook = XLSX.utils.book_new();
const summaryData: any[] = [];

files.forEach((file) => {
  const filePath = path.join(dataDir, file);
  const csvContent = fs.readFileSync(filePath, 'utf8');
  // Skip the first two lines (metadata and header), then parse
  const csvLines = csvContent.split(/\r?\n/).filter(line => line.trim() !== '');
  if (csvLines.length < 2) {
    console.warn(`File ${file} does not have enough lines to process.`);
    return;
  }
  const header = csvLines[1].replace(/^\uFEFF/, ''); // Remove BOM if present
  const cleanedCsv = [header, ...csvLines.slice(2)].join('\n');
  const records = parse(cleanedCsv, { columns: true, skip_empty_lines: true, relax_column_count: true });
  // Count only rows where the '#' column is a number (shot rows)
  const shotCount = records.filter((row: any) => !isNaN(Number(row['#']))).length;

  // Extract summary stats from the raw lines
  const findStat = (label: string) => {
    const line = csvLines.find(l => l.toUpperCase().startsWith(label));
    if (!line) return '';
    const parts = line.split(',');
    // Find the first numeric value after the label
    const stat = parts.find((v, i) => i > 0 && !isNaN(Number(v.trim())));
    return stat ? Number(stat.trim()) : '';
  };
  const avgSpeed = findStat('AVERAGE SPEED');
  const stdDev = findStat('STD DEV');
  const spread = findStat('SPREAD');

  // Add a worksheet for this CSV
  const ws = XLSX.utils.json_to_sheet(records);
  const sheetName = path.basename(file, '.csv').slice(0, 31); // Excel tab name max 31 chars
  XLSX.utils.book_append_sheet(workbook, ws, sheetName);
  // Add summary info (shot count and stats)
  summaryData.push({ File: file, Shots: shotCount, 'Average Speed': avgSpeed, 'Std Dev': stdDev, Spread: spread });
});

// Add summary tab as the first sheet
const summarySheet = XLSX.utils.json_to_sheet(summaryData);
XLSX.utils.book_append_sheet(workbook, summarySheet, 'Summary');
// Move the summary sheet to the first position
const sheetNames = workbook.SheetNames;
if (sheetNames[sheetNames.length - 1] !== 'Summary') {
  // Defensive: should always be last, but just in case
  sheetNames.splice(sheetNames.indexOf('Summary'), 1);
  sheetNames.unshift('Summary');
} else {
  sheetNames.unshift(sheetNames.pop()!);
}
workbook.SheetNames = sheetNames;

// Write the workbook
XLSX.writeFile(workbook, outputFile);
console.log('Processed all CSV files into', outputFile);
