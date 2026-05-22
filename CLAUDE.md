# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

- **Run the processor:** `pnpm exec ts-node index.ts` (also available as the VS Code build task "Run TypeScript with ts-node")
- **Install dependencies:** `pnpm install`
- No build, lint, or test setup exists (`package.json` `test` script is a placeholder that exits 1).

## What this does

Single-script tool that consolidates Garmin chronograph CSV exports (rifle shot velocity sessions) from `data_files/*.csv` into one formatted Excel workbook at `output/summary.xlsx`. All logic lives in `index.ts`; there is no module structure.

## Input CSV format

The processor depends on the exact shape of Garmin's export, which is **not** standard CSV:

- **Line 1:** session metadata (e.g. `"Rifle session started at 10:45"`) — skipped.
- **Line 2:** the real column header (`#,SPEED (FPS),Δ AVG (FPS),...`), may carry a leading BOM.
- **Line 3+:** shot rows. A row is counted as a shot only when its `#` column is numeric.
- **Trailing summary block:** key/value lines located by case-insensitive line-prefix match — `AVERAGE SPEED`, `STD DEV`, `SPREAD`, and `DATE`.

Quirks that have already bitten and are easy to reintroduce:
- Column keys are **case-sensitive** after parsing — the header is `SPEED (FPS)` (all caps), not `Speed (FPS)`.
- The `DATE` line is `DATE, "November 10, 2024 at 10:45 AM"` with a **space before the opening quote**, which strict CSV rejects. `parse()` uses `relax_quotes: true` to tolerate this, and the date regex allows the space: `/Date,\s*"([^"]+)"/i`.

## Processing flow (`index.ts`)

1. For each CSV: drop line 1, treat line 2 as header, re-join, and parse the shot rows with `csv-parse/sync`.
2. Compute per-file `shotCount` (numeric `#`) and `speedSum` (sum of `SPEED (FPS)` over shot rows); read `AVERAGE SPEED` / `STD DEV` / `SPREAD` / `DATE` from the summary block.
3. Write one worksheet per CSV (sheet name truncated to 31 chars — Excel's hard limit).
4. Build the `Summary` sheet, sort its rows by parsed `DATE` ascending, append `TOTAL` (shots + speed sum) and `OVERALL AVG` (totalSpeedSum / totalShots) rows, then move `Summary` to the first tab position before writing.

`exceljs` produces the workbook. `xlsx` is listed as a dependency but is currently unused.

## Conventions

- `node_modules/` is committed to this repo, so dependency installs show up as large diffs in `git status`.
- The generated `output/summary.xlsx` is intentionally tracked (not gitignored).
