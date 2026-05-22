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
    // Create the Summary sheet first so it stays the first tab; it is populated
    // (transposed) after all rows are collected and sorted.
    const summarySheet = workbook.addWorksheet('Summary');

    // Collect summary rows.
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
        const records = parse(cleanedCsv, { columns: true, skip_empty_lines: true, relax_column_count: true, relax_quotes: true });
        // Per-shot speeds (rows with a valid numeric shot index and parseable speed)
        const speeds: number[] = records
            .filter((row: any) => !isNaN(Number(row['#'])))
            .map((row: any) => Number(row['SPEED (FPS)']))
            .filter((n: number) => !isNaN(n));
        const shotCount = records.filter((row: any) => !isNaN(Number(row['#']))).length;
        const highest = speeds.length ? Math.max(...speeds) : '';
        const lowest = speeds.length ? Math.min(...speeds) : '';
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
            const match = dateLine.match(/Date,\s*"([^"]+)"/i);
            dateValue = match ? match[1].trim() : '';
        }
        // Add to summaryRows array
        summaryRows.push({ file, shotCount, avgSpeed, highest, lowest, stdDev, spread, dateValue });
        // Add worksheet for this CSV
        const ws = workbook.addWorksheet(path.basename(file, '.csv').slice(0, 31));
        // Add CSV data
        // Convert values to numbers where possible for the header and each row
        const headerFields = header.split(',');
        ws.addRow(headerFields);
        records.forEach((row: any) => {
            const values = Object.values(row).map((v, i) => {
                // Try to convert to number if not empty and not the first column (which may be '#')
                if (typeof v === 'string' && v.trim() !== '' && !isNaN(Number(v))) {
                    return Number(v);
                }
                return v;
            });
            ws.addRow(values);
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
    // Build the transposed Summary sheet: one column per session (numbered by
    // chronological order), one row per statistic.
    // Header row: "Row" label followed by a session number per file.
    summarySheet.addRow(['Row', ...summaryRows.map((_, i) => i + 1)]);
    // Each stat becomes a row, in the order of the target layout (Grains omitted).
    const statRows: Array<[string, (r: any) => any]> = [
        ['Shots', r => r.shotCount],
        ['V0 Avg', r => r.avgSpeed],
        ['Highest', r => r.highest],
        ['Lowest', r => r.lowest],
        ['Spread', r => r.spread],
        ['Std Dev', r => r.stdDev],
    ];
    statRows.forEach(([label, getter]) => {
        summarySheet.addRow([label, ...summaryRows.map(getter)]);
    });
    // Bold the header row and the stat-label column.
    summarySheet.getRow(1).eachCell(cell => {
        cell.font = { bold: true };
        cell.alignment = { horizontal: 'center', vertical: 'middle' };
    });
    summarySheet.getColumn(1).eachCell(cell => {
        cell.font = { bold: true };
    });
    // Column widths: wider label column, uniform session columns.
    summarySheet.getColumn(1).width = 12;
    for (let c = 2; c <= summaryRows.length + 1; c++) {
        summarySheet.getColumn(c).width = 10;
        summarySheet.getColumn(c).alignment = { horizontal: 'center' };
    }
    // Write the workbook
    await workbook.xlsx.writeFile(outputFile);
    console.log('Processed all CSV files into', path.relative(process.cwd(), outputFile));
}

processCSVsToExcel();
