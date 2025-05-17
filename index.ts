console.log('Garmin Data Processor initialized');

import * as fs from 'fs';
import * as path from 'path';
import { parse } from 'csv-parse/sync';
import ExcelJS from 'exceljs';

const dataDir = path.join(__dirname, 'data_files');
const outputDir = path.join(__dirname, 'output');
const outputFile = path.join(outputDir, 'summary.xlsx');

// Ensure output directory exists
if (!fs.existsSync(outputDir)) {
    fs.mkdirSync(outputDir);
}

const files = fs.readdirSync(dataDir).filter(f => f.endsWith('.csv'));

async function processCSVsToExcel() {
    const workbook = new ExcelJS.Workbook();
    // Create the summary worksheet
    const summarySheet = workbook.addWorksheet('Summary');
    // Add headers
    const summaryHeaders = ['File', 'Date', 'Shots', 'Avg Speed', 'Std Dev', 'Spread'];
    summarySheet.addRow(summaryHeaders);
    // Style the header row
    summarySheet.getRow(1).eachCell(cell => {
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
        cell.font = { bold: true, underline: true, size: 14 };
    });
    // Set column widths
    summarySheet.columns = [
        { header: 'File', width: 40 },
        { header: 'Date', width: 18 },
        { header: 'Shots', width: 8 },
        { header: 'Avg Speed', width: 12 },
        { header: 'Std Dev', width: 10 },
        { header: 'Spread', width: 8 },
    ];

    // Collect summary rows
    const summaryRows: any[] = [];

    for (const file of files) {
        const filePath = path.join(dataDir, file);
        const csvContent = fs.readFileSync(filePath, 'utf8');
        const csvLines = csvContent.split(/\r?\n/).filter(line => line.trim() !== '');
        if (csvLines.length < 2) {
            console.warn(`File ${file} does not have enough lines to process.`);
            continue;
        }
        const header = csvLines[1].replace(/^\uFEFF/, '');
        const cleanedCsv = [header, ...csvLines.slice(2)].join('\n');
        const records = parse(cleanedCsv, { columns: true, skip_empty_lines: true, relax_column_count: true });
        const shotCount = records.filter((row: any) => !isNaN(Number(row['#']))).length;
        const findStat = (label: string) => {
            const line = csvLines.find(l => l.toUpperCase().startsWith(label));
            if (!line) return '';
            const parts = line.split(',');
            const stat = parts.find((v, i) => i > 0 && !isNaN(Number(v.trim())));
            return stat ? Number(stat.trim()) : '';
        };
        const avgSpeed = findStat('AVERAGE SPEED');
        const stdDev = findStat('STD DEV');
        const spread = findStat('SPREAD');
        // Find the Date field (look for line starting with 'Date')
        const dateLine = csvLines.find(l => l.toUpperCase().startsWith('DATE'));
        let dateValue = '';
        if (dateLine) {
            // Use a regex to extract the quoted string after the first comma
            const match = dateLine.match(/Date,"([^"]+)"/i);
            dateValue = match ? match[1].trim() : '';
        }
        // Add to summaryRows array
        summaryRows.push({ file, shotCount, avgSpeed, stdDev, spread, dateValue });
        // Add worksheet for this CSV
        const ws = workbook.addWorksheet(path.basename(file, '.csv').slice(0, 31));
        // Add CSV data
        ws.addRow(header.split(','));
        records.forEach((row: any) => {
            ws.addRow(Object.values(row));
        });
        // Optionally, set column widths for data sheets
        ws.columns = header.split(',').map(h => ({ header: h, width: 15 }));
    }
    // Sort summaryRows by dateValue ascending (parse as date)
    summaryRows.sort((a, b) => {
        const da = new Date(a.dateValue).getTime();
        const db = new Date(b.dateValue).getTime();
        return da - db;
    });
    // Add sorted rows to summarySheet
    summaryRows.forEach(row => {
        summarySheet.addRow([
            row.file,
            row.dateValue,
            row.shotCount,
            row.avgSpeed,
            row.stdDev,
            row.spread
        ]);
    });
    // Move summary to first position
    workbook.worksheets.splice(0, 0, workbook.worksheets.pop()!);
    // Write the workbook
    await workbook.xlsx.writeFile(outputFile);
    console.log('Processed all CSV files into', outputFile);
}

processCSVsToExcel();
