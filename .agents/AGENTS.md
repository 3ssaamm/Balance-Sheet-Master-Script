# Project Rules

## Deployment
- **Never push to GitHub (`git push`)** for this project.
- To deploy code changes, always use **`clasp push`** — this syncs directly to Google Apps Script.

## Architecture & Calculations
- **Formula-First Principle**: Always favor native Google Sheets equations/formulas over static script-calculated values.
- Save Apps Script for ingestion, file parsing, structural syncing, and complex non-formula math.
- Never set aggressive `onEdit` auto-rebuild triggers that wipe or overwrite the sheet on minor edits; use manual menu runs or non-destructive updates.
