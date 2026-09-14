// =================================================================
// MASTER SCRIPT — CENTRALIZED BALANCE SHEET CONTROLLER
// =================================================================
// Change this ID each time you switch to a new monthly balance file.
// This is the ONLY line you need to update.
const TARGET_SHEET_ID = "1kCE23R_uDKgRRukvxPNcusFFeD3yxivEjbaY5kHLE4M";

// =================================================================
// MENU — Appears at top of the Google Sheets menu bar when the
//         spreadsheet is opened.
// =================================================================
function onOpen() {
    try {
        const ui = SpreadsheetApp.getUi();
        ui.createMenu('⚙️ Balance Sheet')
            .addItem('▶ Run Balance Report', 'runDailyBalance')
            .addSeparator()
            .addItem('📥 Import: NET Data  (from "Raw Data - NET" sheet)', 'importNetData')
            .addItem('📥 Import: Trips Data  (from "Raw Data - Trips" sheet)', 'importTripsData')
            .addItem('🕒 Sync Working Hours Sheet', 'generateWorkingHoursSheetFromMenu')
            .addSeparator()
            .addItem('🔧 Setup Auto-Run Trigger', 'setupAutoRun')
            .addToUi();
    } catch (e) {
        // onOpen() cannot be run from the Apps Script editor.
        // It runs automatically when the spreadsheet is opened in a browser.
        Logger.log("ℹ️ onOpen() skipped: " + e.message);
    }
}

// =================================================================
// IMPORT: NET DATA — Merges "Raw Data - NET" into "Raw Data".
//
// How it works:
//  1. Reads "Raw Data - NET" — same layout as Raw Data but daily
//     values are NET pay (Credit - Cash) per driver per day.
//  2. For each date column in the NET sheet:
//     a. If the date already exists in Raw Data → overwrites that
//        driver's NET cell value.
//     b. If the date is NEW → inserts a new column in Raw Data,
//        fills it with 0 credit for all drivers (you can add
//        credit data later), and writes the NET values.
//  3. Recalculates the "Driver NET" summary column (last column)
//     in Raw Data by summing all daily NET columns.
//  4. After import, runs the full balance report automatically.
// =================================================================
function importNetData() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

    const netSheet = ss.getSheetByName("Raw Data - NET");
    if (!netSheet) {
        Logger.log('❌ Sheet "Raw Data - NET" not found. Create it, paste your NET data, then try again.');
        return;
    }
    const rawSheet = ss.getSheetByName("Raw Data");
    if (!rawSheet) { Logger.log('❌ Sheet "Raw Data" not found.'); return; }

    const rawWeeks = parseRawDataWeeks(rawSheet);
    if (rawWeeks.length === 0) {
        Logger.log('⚠️ No valid week tables found in "Raw Data".');
        return;
    }

    const netWeeks = parseRawDataWeeks(netSheet);
    let updatedDrivers = 0;

    if (netWeeks.length > 0) {
        // Multi-week vertical tables in NET sheet
        rawWeeks.forEach((rawW, wIdx) => {
            const netW = netWeeks[wIdx];
            if (!netW) return;

            const netDriverMap = {};
            netW.drivers.forEach(d => {
                netDriverMap[d.name.toLowerCase()] = d.weeklyNet;
            });

            rawW.drivers.forEach(rawD => {
                const canon = rawD.name.toLowerCase();
                let netVal = netDriverMap[canon];
                if (netVal === undefined) {
                    // Try middle-initial-free match
                    for (const netName in netDriverMap) {
                        if (netName.replace(/\s+[a-z]\.?\s+/g, ' ') === canon.replace(/\s+[a-z]\.?\s+/g, ' ')) {
                            netVal = netDriverMap[netName];
                            break;
                        }
                    }
                }
                if (netVal !== undefined) {
                    rawSheet.getRange(rawD.rowIdx + 1, rawW.netColIdx + 1).setValue(netVal);
                    updatedDrivers++;
                }
            });
        });
    } else {
        // Fallback: single table in NET sheet
        const netData = netSheet.getDataRange().getValues();
        const netHeaders = netSheet.getRange(1, 1, 1, netSheet.getLastColumn()).getDisplayValues()[0];
        const netDriverNetColIdx = findColByLabel(netHeaders, ['driver net', 'net'], netHeaders.length - 1);

        const netDriverMap = {};
        for (let r = 2; r < netData.length; r++) {
            const name = (netData[r][0] || '').toString().trim();
            if (name && name !== 'Total') {
                netDriverMap[name.toLowerCase()] = parseNumber(netData[r][netDriverNetColIdx]) || 0;
            }
        }

        rawWeeks.forEach(rawW => {
            rawW.drivers.forEach(rawD => {
                const canon = rawD.name.toLowerCase();
                let netVal = netDriverMap[canon];
                if (netVal !== undefined) {
                    rawSheet.getRange(rawD.rowIdx + 1, rawW.netColIdx + 1).setValue(netVal);
                    updatedDrivers++;
                }
            });
        });
    }

    SpreadsheetApp.flush();
    Logger.log(`✅ NET Data import complete! Updated Driver NET for ${updatedDrivers} driver-week records.`);
    Logger.log(`   Running balance report...`);
    runDailyBalance();
}

// =================================================================
// IMPORT: TRIPS DATA — Merges "Raw Data - Trips" into "Raw Data".
// =================================================================
function importTripsData() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

    const tripsSheet = ss.getSheetByName("Raw Data - Trips");
    if (!tripsSheet) {
        Logger.log('❌ Sheet "Raw Data - Trips" not found. Create it, paste your Trips data, then try again.');
        return;
    }
    const rawSheet = ss.getSheetByName("Raw Data");
    if (!rawSheet) { Logger.log('❌ Sheet "Raw Data" not found.'); return; }

    const rawWeeks = parseRawDataWeeks(rawSheet);
    if (rawWeeks.length === 0) {
        Logger.log('⚠️ No valid week tables found in "Raw Data".');
        return;
    }

    const tripsWeeks = parseRawDataWeeks(tripsSheet);
    let updatedCount = 0;

    if (tripsWeeks.length > 0) {
        rawWeeks.forEach((rawW, wIdx) => {
            const tripsW = tripsWeeks[wIdx];
            if (!tripsW) return;

            const tripsDriverMap = {};
            tripsW.drivers.forEach(d => {
                tripsDriverMap[d.name.toLowerCase()] = d.weeklyTrips;
            });

            rawW.drivers.forEach(rawD => {
                const canon = rawD.name.toLowerCase();
                let tripVal = tripsDriverMap[canon];
                if (tripVal === undefined) {
                    for (const tripsName in tripsDriverMap) {
                        if (tripsName.replace(/\s+[a-z]\.?\s+/g, ' ') === canon.replace(/\s+[a-z]\.?\s+/g, ' ')) {
                            tripVal = tripsDriverMap[tripsName];
                            break;
                        }
                    }
                }
                if (tripVal !== undefined) {
                    rawSheet.getRange(rawD.rowIdx + 1, rawW.countColIdx + 1).setValue(tripVal);
                    updatedCount++;
                }
            });
        });
    } else {
        const tripsData = tripsSheet.getDataRange().getValues();
        const tripsHeaders = tripsSheet.getRange(1, 1, 1, tripsSheet.getLastColumn()).getDisplayValues()[0];
        const tripsCountColIdx = findColByLabel(tripsHeaders, ['count', 'trip', 'trips', 'total'], tripsHeaders.length - 2);

        const driverTripsMap = {};
        for (let r = 2; r < tripsData.length; r++) {
            const name = (tripsData[r][0] || '').toString().trim();
            if (name && name !== 'Total') {
                driverTripsMap[name.toLowerCase()] = parseNumber(tripsData[r][tripsCountColIdx]) || 0;
            }
        }

        rawWeeks.forEach(rawW => {
            rawW.drivers.forEach(rawD => {
                const canon = rawD.name.toLowerCase();
                if (driverTripsMap[canon] !== undefined) {
                    rawSheet.getRange(rawD.rowIdx + 1, rawW.countColIdx + 1).setValue(driverTripsMap[canon]);
                    updatedCount++;
                }
            });
        });
    }

    SpreadsheetApp.flush();
    Logger.log(`✅ Trips Data import complete! Updated ${updatedCount} driver-week records.`);
    Logger.log(`   Running balance report...`);
    runDailyBalance();
}

/**
 * Helper: scans a header row array for the first column whose label matches
 * one of the given keywords (case-insensitive). Returns that column's 0-based
 * index. Falls back to `defaultIdx` if nothing matches.
 */
function findColByLabel(headers, keywords, defaultIdx) {
    for (let c = headers.length - 1; c >= 0; c--) {   // scan from right (summary cols first)
        const h = (headers[c] || '').toString().toLowerCase().trim();
        for (const kw of keywords) {
            if (h.includes(kw)) return c;
        }
    }
    return defaultIdx;
}

/**
 * Scans the main Summary table headers (from left to right, within columns 0-8)
 * looking for a column matching one of the given keywords.
 * Never scans past column index 8 to avoid matching adjacent tables (stats/daily).
 */
function findMainSummaryCol(headers, keywords, defaultIdx) {
    if (!headers || !headers.length) return defaultIdx;
    const maxCol = Math.min(headers.length, 9);
    for (let c = 0; c < maxCol; c++) {
        const h = (headers[c] || '').toString().toLowerCase().trim();
        for (const kw of keywords) {
            if (h === kw || h.includes(kw)) return c;
        }
    }
    return defaultIdx;
}

/**
 * Reads a sheet containing one or more weekly tables stacked vertically.
 * Weekly tables:
 *  - Start at rows 1, 21, 41, 61, 81... (or any row where Column A is "Driver")
 *  - Header row: Column A is "Driver", followed by date columns (DD-MM-YY), then Total, Count, Driver NET
 *  - Row 2 of each table is a "Total" row (Column A = "Total")
 *  - Driver rows follow until an empty cell in Column A or the next table header
 *
 * Returns an array of weekly table objects:
 * [
 *   {
 *     weekIndex: number,
 *     headerRowIdx: number,      // 0-based
 *     dateColumns: [ { colIdx, date, headerStr } ],
 *     totalColIdx: number,
 *     countColIdx: number,
 *     netColIdx: number,
 *     drivers: [
 *       {
 *         rowIdx: number,        // 0-based row index in sheet
 *         name: string,
 *         dailyCredits: { [dateTimestamp]: number },
 *         weeklyTotal: number,
 *         weeklyTrips: number,
 *         weeklyNet: number,
 *         weeklyCash: number
 *       }
 *     ]
 *   }
 * ]
 */
function parseRawDataWeeks(sheet) {
    if (!sheet) return [];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const numCols = sheet.getLastColumn();
    const displayValues = sheet.getRange(1, 1, data.length, numCols).getDisplayValues();

    const weeks = [];
    let r = 0;

    while (r < data.length) {
        const firstCell = (displayValues[r][0] || "").toString().trim().toLowerCase();
        if (firstCell === "driver") {
            const headerRowIdx = r;
            const headerRow = displayValues[headerRowIdx];
            const numHeaderCols = headerRow.length;

            const dateColumns = [];
            for (let c = 1; c < numHeaderCols; c++) {
                const d = parseDDMMYY(headerRow[c]);
                if (d) {
                    dateColumns.push({ colIdx: c, date: d, headerStr: headerRow[c] });
                }
            }

            if (dateColumns.length > 0) {
                const lastDateCol = dateColumns[dateColumns.length - 1].colIdx;
                const totalColIdx = findColByLabel(headerRow, ["total"], lastDateCol + 1);
                const countColIdx = findColByLabel(headerRow, ["count", "trip", "trips"], lastDateCol + 2);
                const netColIdx = findColByLabel(headerRow, ["driver net", "net"], lastDateCol + 3);

                const weekDrivers = [];
                let dr = headerRowIdx + 1;

                while (dr < data.length) {
                    const cellVal = (displayValues[dr][0] || "").toString().trim();
                    const cellLower = cellVal.toLowerCase();

                    if (cellLower === "driver") {
                        break;
                    }

                    if (cellLower === "total") {
                        dr++;
                        continue;
                    }

                    if (cellVal !== "") {
                        const driverName = cellVal;
                        const weeklyTotal = parseNumber(data[dr][totalColIdx]) || 0;
                        const weeklyTrips = parseNumber(data[dr][countColIdx]) || 0;
                        const weeklyNet = parseNumber(data[dr][netColIdx]) || 0;
                        const weeklyCash = roundToTwo(weeklyTotal - weeklyNet);

                        const dailyCredits = {};
                        dateColumns.forEach(dc => {
                            const cred = parseNumber(data[dr][dc.colIdx]) || 0;
                            dailyCredits[dc.date.getTime()] = cred;
                        });

                        weekDrivers.push({
                            rowIdx: dr,
                            name: driverName,
                            dailyCredits: dailyCredits,
                            weeklyTotal: weeklyTotal,
                            weeklyTrips: weeklyTrips,
                            weeklyNet: weeklyNet,
                            weeklyCash: weeklyCash
                        });
                    }

                    dr++;
                    if (dr - headerRowIdx >= 20 && displayValues[dr] && (displayValues[dr][0] || "").toString().trim().toLowerCase() === "driver") {
                        break;
                    }
                }

                weeks.push({
                    weekIndex: weeks.length,
                    headerRowIdx: headerRowIdx,
                    dateColumns: dateColumns,
                    totalColIdx: totalColIdx,
                    countColIdx: countColIdx,
                    netColIdx: netColIdx,
                    drivers: weekDrivers
                });

                r = dr;
                continue;
            }
        }
        r++;
    }

    return weeks;
}

/**
 * Normalizes driver names to consolidate variations across weeks
 * (e.g., "Hassan L En Nejjari" vs "Hassan En Nejjari").
 */
function getCanonicalDriverName(name, canonicalList) {
    if (!name) return "";
    const trimmed = name.toString().trim();
    if (canonicalList.includes(trimmed)) return trimmed;

    function simplify(n) {
        return n.toLowerCase()
            .replace(/\s+[a-z]\.?\s+/g, " ")
            .replace(/\s+/g, " ")
            .trim();
    }

    const simplified = simplify(trimmed);
    for (const canon of canonicalList) {
        if (simplify(canon) === simplified) {
            return canon;
        }
    }

    const matched = matchOldDriverName(trimmed, canonicalList);
    if (matched && canonicalList.includes(matched)) {
        return matched;
    }

    return trimmed;
}




/**
 * Run this function ONCE every time you change the TARGET_SHEET_ID above.
 * Installs the onChange trigger so the balance report auto-runs on data change.
 */
function setupAutoRun() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

    // Remove all existing triggers to prevent duplicates
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

    // Install the auto-run trigger: re-runs balance report whenever data changes
    ScriptApp.newTrigger('runDailyBalance')
        .forSpreadsheet(ss)
        .onChange()
        .create();

    Logger.log("✅ Auto-run trigger installed for: " + ss.getName());
    Logger.log("   • onChange → balance report re-runs automatically on data change");
    Logger.log("");
    Logger.log("To run imports manually, select the function from the dropdown in the Apps Script editor:");
    Logger.log("   importNetData()   — merges 'Raw Data - NET' sheet into Raw Data");
    Logger.log("   importTripsData() — merges 'Raw Data - Trips' sheet into Raw Data");
}


// =================================================================
// WORKING HOURS SHEET GENERATION & LOAD HELPERS
// =================================================================

function generateWorkingHoursSheetFromMenu() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    generateWorkingHoursSheet(ss);
    SpreadsheetApp.flush();
    try {
        SpreadsheetApp.getUi().alert("✅ 'Working hours' sheet synced successfully!");
    } catch (e) {}
}

/**
 * Reads the "Working hours" sheet and returns:
 * - map: { driverName -> { dateTimestamp -> hours } }
 * - driverTotalHours: { driverName -> totalHours }
 * - dailyTotalHours: { dateTimestamp -> totalHours }
 */
function loadWorkingHoursMap(ss) {
    const hoursSheet = ss.getSheetByName("Working hours");
    if (!hoursSheet) return { map: {}, driverTotalHours: {}, dailyTotalHours: {} };

    const data = hoursSheet.getDataRange().getValues();
    if (data.length < 2) return { map: {}, driverTotalHours: {}, dailyTotalHours: {} };

    const headers = hoursSheet.getRange(1, 1, 1, hoursSheet.getLastColumn()).getDisplayValues()[0];
    const dateCols = {}; // colIdx -> timestamp
    for (let c = 1; c < headers.length; c++) {
        if (headers[c].toLowerCase().includes("total")) continue;
        const d = parseDDMMYY(headers[c]);
        if (d) dateCols[c] = d.getTime();
    }

    const map = {};
    const driverTotalHours = {};
    const dailyTotalHours = {};

    for (let r = 1; r < data.length; r++) {
        const driver = (data[r][0] || "").toString().trim();
        if (!driver || driver.toLowerCase().startsWith("total")) continue;

        map[driver] = {};
        let driverSum = 0;

        for (const c in dateCols) {
            const dateKey = dateCols[c];
            const val = parseNumber(data[r][c]);
            if (val !== null && !isNaN(val) && val > 0) {
                map[driver][dateKey] = val;
                driverSum += val;
                dailyTotalHours[dateKey] = (dailyTotalHours[dateKey] || 0) + val;
            }
        }
        driverTotalHours[driver] = roundToTwo(driverSum);
    }

    return { map, driverTotalHours, dailyTotalHours };
}

/**
 * Generates/syncs the "Working hours" sheet based on "Raw Data".
 * Preserves existing entered hours, defaults $0 credit days to 0 hours,
 * adds a Total Hours column (formula =SUM(...)) and a Total Hours row at bottom.
 * Applies custom data validation for 0.5-hour increments.
 */
function generateWorkingHoursSheet(ss) {
    let hoursSheet = ss.getSheetByName("Working hours");
    if (!hoursSheet) hoursSheet = ss.insertSheet("Working hours");

    const rawSheet = ss.getSheetByName("Raw Data");
    if (!rawSheet) {
        Logger.log('❌ Sheet "Raw Data" not found.');
        return;
    }

    const weeks = parseRawDataWeeks(rawSheet);
    if (weeks.length === 0) {
        Logger.log('⚠️ No valid week tables found in "Raw Data".');
        return;
    }

    // Collect all date columns across all weeks (deduplicate and sort chronologically)
    const dateMap = {}; // timestamp -> { date, headerStr }
    weeks.forEach(w => {
        w.dateColumns.forEach(dc => {
            const t = dc.date.getTime();
            if (!dateMap[t]) {
                dateMap[t] = { date: dc.date, headerStr: dc.headerStr };
            }
        });
    });

    const sortedTimestamps = Object.keys(dateMap).map(Number).sort((a, b) => a - b);
    if (sortedTimestamps.length === 0) return;

    const allDateCols = sortedTimestamps.map(t => dateMap[t]);

    // Seed canonical drivers from Settings
    const settingsSheet = ss.getSheetByName("Settings");
    const canonicalDrivers = [];
    if (settingsSheet) {
        const sData = settingsSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
            const d = (sData[i][0] || "").toString().trim();
            if (d && !canonicalDrivers.includes(d)) canonicalDrivers.push(d);
        }
    }

    // Collect all unique drivers and their credit across all weeks
    const driverList = [];
    const driverCreditMap = {}; // canonDriver -> dateTimestamp -> credit

    weeks.forEach(w => {
        w.drivers.forEach(d => {
            const canon = getCanonicalDriverName(d.name, canonicalDrivers);
            if (!canonicalDrivers.includes(canon)) canonicalDrivers.push(canon);
            if (!driverList.includes(canon)) driverList.push(canon);

            if (!driverCreditMap[canon]) driverCreditMap[canon] = {};
            for (const t in d.dailyCredits) {
                driverCreditMap[canon][t] = (driverCreditMap[canon][t] || 0) + d.dailyCredits[t];
            }
        });
    });

    // Read existing entered working hours to preserve them
    const { map: existingHoursMap } = loadWorkingHoursMap(ss);

    // Build Header
    const headerRow = ["Driver"];
    allDateCols.forEach(dc => headerRow.push(dc.headerStr));
    headerRow.push("Total Hours");

    const numDateCols = allDateCols.length;
    const lastDateColLetter = getColumnLetter(numDateCols + 1); // Col B is 2

    const matrixRows = [];

    // Build driver rows
    driverList.forEach((driver, idx) => {
        const rowIndex = idx + 2; // Row 2 is first driver
        const row = [driver];

        allDateCols.forEach(dc => {
            const dateKey = dc.date.getTime();
            const rawCredit = driverCreditMap[driver] ? (driverCreditMap[driver][dateKey] || 0) : 0;

            // Existing value takes priority
            if (existingHoursMap[driver] && existingHoursMap[driver][dateKey] !== undefined) {
                row.push(existingHoursMap[driver][dateKey]);
            } else if (rawCredit === 0) {
                // If $0 in raw data, default to 0 in working time
                row.push(0);
            } else {
                row.push(""); // blank for user to enter working hours
            }
        });

        // Formula for Total Hours row
        row.push(`=SUM(B${rowIndex}:${lastDateColLetter}${rowIndex})`);
        matrixRows.push(row);
    });

    // Summary Bottom Row (Total Hours Per Date)
    const bottomRowIndex = driverList.length + 2;
    const bottomRow = ["Total Hours"];
    for (let c = 0; c < numDateCols; c++) {
        const colLetter = getColumnLetter(c + 2);
        bottomRow.push(`=SUM(${colLetter}2:${colLetter}${bottomRowIndex - 1})`);
    }
    bottomRow.push(`=SUM(B${bottomRowIndex}:${lastDateColLetter}${bottomRowIndex})`);
    matrixRows.push(bottomRow);

    // Write to Sheet
    hoursSheet.clear();
    hoursSheet.getRange(1, 1, 1, headerRow.length).setValues([headerRow]);
    hoursSheet.getRange(2, 1, matrixRows.length, headerRow.length).setValues(matrixRows);

    // Formatting
    const numRows = matrixRows.length;
    const numCols = headerRow.length;

    // Header styling
    hoursSheet.getRange(1, 1, 1, numCols)
        .setFontWeight("bold")
        .setBackground("#3c78d8")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    // Matrix numbers format & alignment
    hoursSheet.getRange(2, 2, numRows, numCols - 1)
        .setNumberFormat("0.0")
        .setHorizontalAlignment("center");

    // Driver names column formatting
    hoursSheet.getRange(2, 1, numRows, 1).setFontWeight("bold");

    // Bottom total row styling
    hoursSheet.getRange(bottomRowIndex, 1, 1, numCols)
        .setFontWeight("bold")
        .setBackground("#d0e0e3")
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);

    // Outer and grid borders
    hoursSheet.getRange(1, 1, numRows + 1, numCols)
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    // Data Validation & Conditional Formatting Rules (cells B2 to lastDateCol, drivers only)
    if (driverList.length > 0 && numDateCols > 0) {
        const matrixRange = hoursSheet.getRange(2, 2, driverList.length, numDateCols);

        const firstCell = "B2";
        const rule = SpreadsheetApp.newDataValidation()
            .requireFormulaSatisfied(`=AND(ISNUMBER(${firstCell}), ${firstCell}>=0, MOD(${firstCell}*2, 1)=0)`)
            .setAllowInvalid(true)
            .setHelpText("Hours must be entered in 0.5 increments (e.g. 0, 0.5, 1, 1.5, 2, ...)")
            .build();

        matrixRange.setDataValidation(rule);

        // Conditional Formatting Rules: 0 = Red, < 8 = Yellow, >= 8 = Green
        const redRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberEqualTo(0)
            .setBackground("#f4cccc") // Soft pastel red
            .setFontColor("#990000") // Dark red text
            .setRanges([matrixRange])
            .build();

        const yellowRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(0.001, 7.999)
            .setBackground("#fff2cc") // Soft pastel yellow
            .setFontColor("#7f6000") // Dark yellow/brown text
            .setRanges([matrixRange])
            .build();

        const greenRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThanOrEqualTo(8)
            .setBackground("#d9ead3") // Soft pastel green
            .setFontColor("#274e13") // Dark green text
            .setRanges([matrixRange])
            .build();

        hoursSheet.setConditionalFormatRules([redRule, yellowRule, greenRule]);
    }

    hoursSheet.autoResizeColumns(1, numCols);
    for (let c = 1; c <= numCols; c++) {
        if (hoursSheet.getColumnWidth(c) < 80) hoursSheet.setColumnWidth(c, 80);
    }
}

/**
 * Helper to convert 1-based column index to letter (1->A, 2->B, etc.)
 */
function getColumnLetter(colIndex) {
    let temp = "";
    let letter = "";
    while (colIndex > 0) {
        temp = (colIndex - 1) % 26;
        letter = String.fromCharCode(65 + temp) + letter;
        colIndex = (colIndex - temp - 1) / 26;
    }
    return letter;
}

/**
 * Main entry point — runs all reporting functions in the correct order.
 */
function runDailyBalance() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

    // 0. Sync working hours sheet first so downstream reports can use working hours data
    generateWorkingHoursSheet(ss);

    // 1. Update the main Summary sheet from Raw Data.
    updateSummaryAndCharts(ss);

    // 2. Wait for all pending spreadsheet changes to apply.
    SpreadsheetApp.flush();

    // 3. Generate the weekly summary.
    generateWeeklySummary(ss);

    // 4. Generate the final Bonus report.
    generateBonusReport(ss);

    // 5. Generate Angel's specific progressive report.
    generateAngelReport(ss);
}


// =================================================================
// SUMMARY BUILD — Reads "Raw Data" tab and builds the Summary sheet
// =================================================================
function updateSummaryAndCharts(ss) {
    let summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) summarySheet = ss.insertSheet("Summary");

    const rawSheet = ss.getSheetByName("Raw Data");
    if (!rawSheet) throw new Error('No sheet named "Raw Data" was found.');

    // --- Helper to load exact daily data from auxiliary sheets ---
    function loadAuxMap(sheetName) {
        const auxSheet = ss.getSheetByName(sheetName);
        if (!auxSheet) return {};
        const weeks = parseRawDataWeeks(auxSheet);
        if (weeks.length > 0) {
            const map = {};
            weeks.forEach(w => {
                w.drivers.forEach(d => {
                    const dName = d.name;
                    if (!map[dName]) map[dName] = {};
                    for (const t in d.dailyCredits) {
                        map[dName][t] = d.dailyCredits[t];
                    }
                });
            });
            return map;
        }

        const data = auxSheet.getDataRange().getValues();
        if (data.length < 2) return {};
        const headers = auxSheet.getRange(1, 1, 1, auxSheet.getLastColumn()).getDisplayValues()[0];

        const dateCols = {};
        for (let c = 1; c < headers.length; c++) {
            const d = parseDDMMYY(headers[c]);
            if (d) dateCols[c] = d.getTime();
        }

        const map = {};
        for (let r = 2; r < data.length; r++) {
            const driver = (data[r][0] || "").toString().trim();
            if (driver && driver !== "Total") {
                map[driver] = {};
                for (const c in dateCols) {
                    const val = parseNumber(data[r][c]);
                    if (val !== null && !isNaN(val)) map[driver][dateCols[c]] = val;
                }
            }
        }
        return map;
    }

    // Load auxiliary daily maps (driverName -> { timestamp -> value })
    const tripsDataMap = loadAuxMap("Raw Data - Trips");
    const netDataMap = loadAuxMap("Raw Data - NET");
    const noShowMap = loadNoShowMap(ss, "No Show");
    const { map: workingHoursMap, driverTotalHours, dailyTotalHours } = loadWorkingHoursMap(ss);

    // --- Read the entire Raw Data sheet as weekly tables ---
    const weeks = parseRawDataWeeks(rawSheet);
    if (weeks.length === 0) {
        Logger.log("⚠️ No valid week tables found in Raw Data.");
        return;
    }

    // Seed canonical drivers from Settings
    let settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) settingsSheet = ss.insertSheet("Settings");
    const settingsData = settingsSheet.getDataRange().getValues();
    let existingFares = {};
    const canonicalDrivers = [];
    for (let i = 1; i < settingsData.length; i++) {
        const d = (settingsData[i][0] || "").toString().trim();
        if (d) {
            if (!canonicalDrivers.includes(d)) canonicalDrivers.push(d);
            existingFares[d] = settingsData[i][1];
        }
    }

    // --- Read existing Summary data BEFORE clearing it ---
    // This acts as a memory of previously-calculated trips/cash values.
    const existingSummaryMap = {}; // { driver -> { dateKey -> { trips, cash } } }
    const existingLastRow = summarySheet.getLastRow();
    if (existingLastRow >= 3) {
        const existingHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 9)).getDisplayValues()[0];
        const existingTripsIdx = findMainSummaryCol(existingHeaders, ['trips', 'trip'], 5);
        const existingCashIdx = findMainSummaryCol(existingHeaders, ['cash'], 6);
        const readCols = Math.max(existingTripsIdx, existingCashIdx) + 1;
        const existingRows = summarySheet.getRange(3, 1, existingLastRow - 2, readCols).getValues();
        let prevSummaryDate = null;
        existingRows.forEach(row => {
            const parsedD = parseDate(row[0]);
            if (parsedD) prevSummaryDate = parsedD;
            else if (row[0] instanceof Date) prevSummaryDate = row[0];
            const dateVal = prevSummaryDate;
            const driver = (row[1] || '').toString().trim();
            if (!dateVal || !driver) return;
            const trips = parseNumber(row[existingTripsIdx]);
            const cash = parseNumber(row[existingCashIdx]);
            if (trips === null && cash === null) return;
            const dk = new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate()).getTime();
            const canonDriver = getCanonicalDriverName(driver, canonicalDrivers);
            if (!existingSummaryMap[canonDriver]) existingSummaryMap[canonDriver] = {};
            existingSummaryMap[canonDriver][dk] = { trips: trips || 0, cash: cash || 0 };
            if (!existingSummaryMap[driver]) existingSummaryMap[driver] = {};
            existingSummaryMap[driver][dk] = { trips: trips || 0, cash: cash || 0 };
        });
    }

    // --- Build output rows: one row per driver per day across all weeks ---
    let outputRows = []; // [date, driver, credit, cash, trips, total, cashTotal, noShow, hours]

    weeks.forEach(w => {
        w.drivers.forEach(d => {
            const driverName = getCanonicalDriverName(d.name, canonicalDrivers);
            if (!canonicalDrivers.includes(driverName)) canonicalDrivers.push(driverName);

            const weeklyTotal = d.weeklyTotal;
            const weeklyTrips = d.weeklyTrips;
            const weeklyCash = d.weeklyCash;

            // Collect all active days for this driver in this week
            let activeDays = [];
            w.dateColumns.forEach(dc => {
                const dateKey = dc.date.getTime();
                let credit = d.dailyCredits[dateKey] || 0;

                const hasExactTrips = tripsDataMap[driverName] && tripsDataMap[driverName][dateKey] !== undefined;
                const hasExactNet = netDataMap[driverName] && netDataMap[driverName][dateKey] !== undefined;

                const exactTrips = hasExactTrips ? tripsDataMap[driverName][dateKey] : 0;
                const exactNet = hasExactNet ? netDataMap[driverName][dateKey] : 0;
                const hasNoShow = noShowMap[driverName] && noShowMap[driverName][dateKey] !== undefined;

                if (hasNoShow) {
                    credit = roundToTwo(credit - noShowMap[driverName][dateKey].cash);
                }

                // Include day if there's credit, trips, non-zero net, or No Show
                if (credit > 0 || exactTrips > 0 || exactNet !== 0 || hasNoShow) {
                    activeDays.push({ date: dc.date, dateKey: dateKey, credit: credit });
                }
            });

            if (activeDays.length === 0) return;

            const nonAuxTripDays = activeDays.filter(day =>
                !(tripsDataMap[driverName] && tripsDataMap[driverName][day.dateKey] !== undefined));
            const nonAuxCashDays = activeDays.filter(day =>
                !(netDataMap[driverName] && netDataMap[driverName][day.dateKey] !== undefined));

            const lastNonAuxTripDay = nonAuxTripDays.length > 0 ? nonAuxTripDays[nonAuxTripDays.length - 1] : null;
            const lastNonAuxCashDay = nonAuxCashDays.length > 0 ? nonAuxCashDays[nonAuxCashDays.length - 1] : null;

            let sumExactTrips = 0;
            let sumExactCash = 0;
            let missingTripDays = [];
            let missingCashDays = [];

            activeDays.forEach(day => {
                const isLastTrip = lastNonAuxTripDay && day.dateKey === lastNonAuxTripDay.dateKey;
                const isLastCash = lastNonAuxCashDay && day.dateKey === lastNonAuxCashDay.dateKey;

                // TRIPS
                if (tripsDataMap[driverName] && tripsDataMap[driverName][day.dateKey] !== undefined) {
                    sumExactTrips += tripsDataMap[driverName][day.dateKey];
                } else if (!isLastTrip && existingSummaryMap[driverName] && existingSummaryMap[driverName][day.dateKey] !== undefined) {
                    sumExactTrips += existingSummaryMap[driverName][day.dateKey].trips;
                } else {
                    missingTripDays.push(day);
                }

                // CASH
                if (netDataMap[driverName] && netDataMap[driverName][day.dateKey] !== undefined) {
                    sumExactCash += roundToTwo(day.credit - netDataMap[driverName][day.dateKey]);
                } else if (!isLastCash && existingSummaryMap[driverName] && existingSummaryMap[driverName][day.dateKey] !== undefined) {
                    sumExactCash += existingSummaryMap[driverName][day.dateKey].cash;
                } else {
                    missingCashDays.push(day);
                }
            });

            const remainingTrips = Math.max(0, weeklyTrips - sumExactTrips);
            const remainingCash = weeklyCash - sumExactCash;

            const lastMissingTripDay = missingTripDays.length > 0 ? missingTripDays[missingTripDays.length - 1] : null;
            const lastMissingCashDay = missingCashDays.length > 0 ? missingCashDays[missingCashDays.length - 1] : null;

            activeDays.forEach(day => {
                let dailyTrips = 0;
                let dailyCash = 0;

                // TRIPS
                if (tripsDataMap[driverName] && tripsDataMap[driverName][day.dateKey] !== undefined) {
                    dailyTrips = tripsDataMap[driverName][day.dateKey];
                } else if (lastMissingTripDay && day.dateKey === lastMissingTripDay.dateKey && remainingTrips > 0) {
                    dailyTrips = remainingTrips;
                } else if (existingSummaryMap[driverName] && existingSummaryMap[driverName][day.dateKey] !== undefined) {
                    dailyTrips = existingSummaryMap[driverName][day.dateKey].trips;
                } else {
                    dailyTrips = 0;
                }

                // CASH
                if (netDataMap[driverName] && netDataMap[driverName][day.dateKey] !== undefined) {
                    dailyCash = roundToTwo(day.credit - netDataMap[driverName][day.dateKey]);
                } else if (lastMissingCashDay && day.dateKey === lastMissingCashDay.dateKey && remainingCash > 0) {
                    dailyCash = remainingCash;
                } else if (existingSummaryMap[driverName] && existingSummaryMap[driverName][day.dateKey] !== undefined) {
                    dailyCash = existingSummaryMap[driverName][day.dateKey].cash;
                } else {
                    dailyCash = 0;
                }

                let dailyNoShowCount = 0;
                if (noShowMap[driverName] && noShowMap[driverName][day.dateKey]) {
                    dailyNoShowCount = noShowMap[driverName][day.dateKey].count;
                }

                const dailyHours = (workingHoursMap[driverName] && workingHoursMap[driverName][day.dateKey]) ? workingHoursMap[driverName][day.dateKey] : 0;

                outputRows.push([
                    day.date,
                    driverName,
                    day.credit,
                    dailyCash,
                    dailyTrips,
                    weeklyTotal,
                    weeklyCash,
                    dailyNoShowCount,
                    dailyHours
                ]);
            });
        });
    });

    if (outputRows.length === 0) {
        Logger.log("⚠️ No driver data found in Raw Data.");
        return;
    }

    // SORT BY DATE CHRONOLOGICALLY, then by driver name
    outputRows.sort((a, b) => {
        const dateDiff = a[0] - b[0];
        if (dateDiff !== 0) return dateDiff;
        return a[1].localeCompare(b[1]);
    });

    // --- Write Summary headers ---
    const headers = ["Date", "Driver", "Credit", "Hours", "Avg Per Hour", "Trips", "Cash", "No Show", "Balance"];
    summarySheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    // Helper to get fare as decimal
    function getFareDecimal(driver) {
        let fareVal = existingFares[driver];
        if (String(fareVal).trim() === "80%-90%") return "80%-90%";
        if (fareVal === undefined || fareVal === null || fareVal === "") return 0.9;
        if (typeof fareVal === "string" && fareVal.includes("%")) return parseFloat(fareVal) / 100;
        if (typeof fareVal === "number") return fareVal > 1 ? fareVal / 100 : fareVal;
        return 0.9;
    }

    // Pre-cache the fare decimal for each driver to speed up loop
    let driverFareCache = {};
    outputRows.forEach(row => {
        if (!driverFareCache[row[1]]) {
            driverFareCache[row[1]] = getFareDecimal(row[1]);
        }
    });

    // --- Build the rows to write (with date grouping / merging) ---
    const rowsToWrite = [];
    let previousDateString = "";
    let dayBlocks = [];
    let currentStartRow = 3;
    let currentGroupSize = 0;

    let progressiveTrackers = {}; // { driver -> { weekStart: 0, weeklyCredit: 0 } }

    outputRows.forEach((row, index) => {
        const currentDateString = row[0] ? row[0].toDateString() : "";
        let displayDate = row[0];

        if (currentDateString === previousDateString) {
            displayDate = "";
            currentGroupSize++;
        } else {
            if (currentGroupSize > 0) {
                dayBlocks.push({ row: currentStartRow, count: currentGroupSize });
            }
            previousDateString = currentDateString;
            currentStartRow = 3 + index;
            currentGroupSize = 1;
        }

        const fare = driverFareCache[row[1]];
        const dailyCredit = Number(row[2]) || 0;
        const dailyCash = Number(row[3]) || 0;
        const trips = Number(row[4]) || 0;
        const dailyNoShow = Number(row[7]) || 0;
        const dailyHours = Number(row[8]) || 0;
        const dailyAvgPerHour = dailyHours > 0 ? roundToTwo(dailyCredit / dailyHours) : 0;

        let dailyBalance = 0;
        if (fare === "80%-90%") {
            let driver = row[1];
            let d = new Date(row[0]);
            let weekStart = getMonday(d).getTime();

            if (!progressiveTrackers[driver]) progressiveTrackers[driver] = { weekStart: 0, weeklyCredit: 0 };
            if (progressiveTrackers[driver].weekStart !== weekStart) {
                progressiveTrackers[driver] = { weekStart: weekStart, weeklyCredit: 0 };
            }

            let todayGrossPayout = 0;
            if (dailyCredit > 0) {
                let currentWeekCredit = progressiveTrackers[driver].weeklyCredit;
                if (currentWeekCredit >= 1000) {
                    todayGrossPayout = dailyCredit * 0.90;
                } else if ((currentWeekCredit + dailyCredit) <= 1000) {
                    todayGrossPayout = dailyCredit * 0.80;
                } else {
                    const creditAtBaseRate = 1000 - currentWeekCredit;
                    const creditAtTopRate = dailyCredit - creditAtBaseRate;
                    todayGrossPayout = (creditAtBaseRate * 0.80) + (creditAtTopRate * 0.90);
                }
            }
            progressiveTrackers[driver].weeklyCredit += dailyCredit;
            dailyBalance = Number(roundToTwo(todayGrossPayout - dailyCash)) || 0;
        } else {
            const numericFare = Number(fare) || 0.9;
            dailyBalance = Number(roundToTwo((dailyCredit * numericFare) - dailyCash)) || 0;
        }

        // Summary rows: Date, Driver, Credit, Hours, Avg Per Hour, Trips, Cash, No Show, Balance
        rowsToWrite.push([
            displayDate,
            row[1],
            dailyCredit,
            roundToTwo(dailyHours) || 0,
            dailyAvgPerHour,
            trips,
            Number(roundToTwo(dailyCash)) || 0,
            dailyNoShow,
            dailyBalance
        ]);
    });

    if (currentGroupSize > 0) {
        dayBlocks.push({ row: currentStartRow, count: currentGroupSize });
    }

    // --- Clear old data and write new ---
    const maxRows = summarySheet.getMaxRows();
    if (maxRows > 2) {
        summarySheet.getRange(3, 1, maxRows - 2, headers.length).clearContent();
        try { summarySheet.getRange(3, 1, maxRows - 2, 1).breakApart(); } catch (e) { }
    }

    const dataRange = summarySheet.getRange(3, 1, rowsToWrite.length, headers.length);
    dataRange.setValues(rowsToWrite);
    dataRange.setBorder(false, false, false, false, false, false);
    dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    summarySheet.getRange(3, 1, rowsToWrite.length, 1).setVerticalAlignment("middle");

    // --- Apply day block borders and merge date cells ---
    dayBlocks.forEach(block => {
        summarySheet.getRange(block.row, 1, block.count, headers.length)
            .setBorder(true, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        if (block.count > 1) {
            summarySheet.getRange(block.row, 1, block.count, 1).mergeVertically();
        }
    });

    // --- Aggregate per driver across all weeks (for the stats table) ---
    let driverStats = {};
    weeks.forEach(w => {
        w.drivers.forEach(d => {
            const driverName = getCanonicalDriverName(d.name, canonicalDrivers);
            if (!driverStats[driverName]) {
                driverStats[driverName] = {
                    totalCredit: 0,
                    totalTrips: 0,
                    totalCash: 0,
                    totalNoShow: 0,
                    driverNet: 0,
                    workingDays: 0,
                    firstWorkDate: null
                };
            }

            driverStats[driverName].totalCredit += d.weeklyTotal;
            driverStats[driverName].totalTrips += d.weeklyTrips;
            driverStats[driverName].driverNet += d.weeklyNet;
            driverStats[driverName].totalCash += d.weeklyCash;

            w.dateColumns.forEach(dc => {
                const cred = d.dailyCredits[dc.date.getTime()] || 0;
                if (cred > 0) {
                    driverStats[driverName].workingDays++;
                    if (!driverStats[driverName].firstWorkDate) {
                        driverStats[driverName].firstWorkDate = dc.date;
                    }
                }
            });
        });
    });

    // Deduct No Show and calculate hourly/daily averages
    for (let driverName in driverStats) {
        const s = driverStats[driverName];
        let totalNoShowCashForDriver = 0;
        let totalNoShowCountForDriver = 0;
        if (noShowMap[driverName]) {
            Object.values(noShowMap[driverName]).forEach(val => {
                totalNoShowCashForDriver += val.cash;
                totalNoShowCountForDriver += val.count;
            });
        }
        s.totalCredit = roundToTwo(s.totalCredit - totalNoShowCashForDriver);
        s.totalCash = roundToTwo(s.totalCash - totalNoShowCashForDriver);
        s.totalNoShow = totalNoShowCountForDriver;
        s.driverNet = roundToTwo(s.driverNet);

        const driverHours = driverTotalHours[driverName] || 0;
        s.totalHours = driverHours;
        s.avgPerHour = driverHours > 0 ? roundToTwo(s.totalCredit / driverHours) : 0;
        s.avgPerDay = s.workingDays > 0 ? roundToTwo(s.totalCredit / s.workingDays) : 0;
    }

    // =================================================================
    // --- AUTO-SYNC "SETTINGS" SHEET (Driver + Fare %) ---
    // =================================================================
    // (settingsSheet and existingFares already read above for balance calc)

    // --- Aggregated stats table (column I, one gap after the 7-col main table) ---
    const aggHeader = ["Driver", "Total Credit", "Total Trips", "Total Cash", "No Show", "Total Hours", "Avg Per Hour", "Driver NET", "Avg Per Day"];
    const aggRows = [aggHeader];

    const newSettingsRows = [["Driver", "Fare %"]];

    for (let driver in driverStats) {
        const s = driverStats[driver];
        aggRows.push([driver, s.totalCredit, s.totalTrips, s.totalCash, s.totalNoShow, s.totalHours, s.avgPerHour, s.driverNet, s.avgPerDay]);

        let fare = existingFares[driver];
        if (fare === undefined || fare === null || fare === "") {
            fare = "90%";
        }
        newSettingsRows.push([driver, fare]);
    }

    // --- Keep inactive drivers in Settings ---
    for (let oldDriver in existingFares) {
        if (!driverStats[oldDriver]) {
            newSettingsRows.push([oldDriver, existingFares[oldDriver]]);
        }
    }

    settingsSheet.clear();
    settingsSheet.getRange(1, 1, newSettingsRows.length, 2).setValues(newSettingsRows);
    settingsSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#d0e0e3");
    settingsSheet.autoResizeColumns(1, 2);

    // =================================================================
    // --- Write aggregated stats to Summary (column K = one gap after 9-col main table) ---
    // =================================================================
    const startRow = 3, startCol = 11; // Column K (A-I = 9-col main table, J = 1-col gap)
    // Clear old aggregate area
    try { summarySheet.getRange(startRow, 10, 60, 12).clearContent(); } catch (e) { }

    const aggRange = summarySheet.getRange(startRow, startCol, aggRows.length, aggRows[0].length);
    aggRange.setValues(aggRows);

    // Decrease cell size / font size of aggregated table so chart can fully cover it
    aggRange.setFontSize(7);
    summarySheet.getRange(startRow, startCol, 1, aggRows[0].length).setFontWeight("bold");

    // Compact column widths for columns K to S (cols 11 to 19)
    summarySheet.setColumnWidth(startCol, 75); // Col K (Driver name)
    for (let c = startCol + 1; c < startCol + aggRows[0].length; c++) {
        summarySheet.setColumnWidth(c, 55); // Cols L to S (Stats)
    }

    // --- Daily stats for chart (aggregate all drivers per day) ---
    let dailyStats = {};
    outputRows.forEach(r => {
        const dateVal = r[0];
        if (!dateVal) return;
        const dateKey = dateVal.getTime();
        if (!dailyStats[dateKey]) {
            dailyStats[dateKey] = { date: dateVal, credit: 0, trips: 0, cash: 0, noShow: 0, hours: 0 };
        }
        dailyStats[dateKey].credit += (r[2] || 0);
        dailyStats[dateKey].trips += (r[4] || 0);
        dailyStats[dateKey].cash += (r[3] || 0);
        dailyStats[dateKey].noShow += (r[7] || 0); // r[7] is No Show
        dailyStats[dateKey].hours += (dailyTotalHours[dateKey] || 0);
    });

    const dailyRows = Object.values(dailyStats)
        .sort((a, b) => a.date - b.date)
        .map(d => [d.date, roundToTwo(d.credit), d.trips, roundToTwo(d.cash), d.noShow, roundToTwo(d.hours)]);

    const dailyHeader = ["Date", "Total Credit", "Trips", "Cash", "No Show", "Total Hours"];
    const dailyStartCol = 22; // Column V (Daily total summary table starts at Column V)

    // Clear old daily table and chart area starting from Column T (col 20)
    summarySheet.getRange(2, 20, 100, 20).clearContent();
    summarySheet.getRange(2, dailyStartCol, 1, dailyHeader.length).setValues([dailyHeader]).setFontWeight("bold");

    if (dailyRows.length > 0) {
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, dailyRows[0].length).setValues(dailyRows);
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, 1).setNumberFormat("dddd, dd");
    }

    // --- Charts ---
    const allCharts = summarySheet.getCharts();
    allCharts.forEach(c => summarySheet.removeChart(c));

    // Chart 1: Daily Total Credit Trend (line) — positioned NEXT TO the daily credit table (Col AC = 29)
    if (dailyRows.length > 0) {
        const dailyChartCol = dailyStartCol + 7; // Col 29 (AC)
        let dailyChart = summarySheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(summarySheet.getRange(3, dailyStartCol, dailyRows.length, 1))
            .addRange(summarySheet.getRange(3, dailyStartCol + 1, dailyRows.length, 1))
            .setOption("title", "Daily Total Credit Trend")
            .setOption("titleTextStyle", { bold: true, fontSize: 24 })
            .setOption("vAxis", { title: "Total Credit" })
            .setOption("hAxis", { title: "Date", format: "MM/dd" })
            .setOption("colors", ["#1f77b4"])
            .setOption("legend", { position: "none" })
            .setOption("pointSize", 7)
            .setPosition(2, dailyChartCol, 0, 0)
            .setOption("width", 850)
            .setOption("height", 350)
            .build();
        summarySheet.insertChart(dailyChart);
    }

    const lastRow = startRow + aggRows.length - 1;
    const dataStart = startRow + 1;
    // Build column letters from startCol (11 = K) for chart ranges
    const colF = getColumnLetter(startCol);         // K (Driver)
    const colG = getColumnLetter(startCol + 1);     // L (Total Credit)
    const colH = getColumnLetter(startCol + 2);     // M (Total Trips)
    const colNoShow = getColumnLetter(startCol + 4); // O (No Show)
    const colHours = getColumnLetter(startCol + 5);  // P (Total Hours)
    const colK = getColumnLetter(startCol + 8);     // S (Avg Per Day)

    // Chart 2: Total Credit by Driver (bar) — POSITIONED AT ROW 3, COL J TO FULLY COVER AGGREGATED STATS TABLE
    let chart1 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colG + dataStart + ":" + colG + lastRow))
        .setOption("title", "Total Credit by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#1f77b4"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Credit" })
        .setOption("width", 740)
        .setOption("height", 450)
        .setPosition(3, startCol, 0, 0).build();
    summarySheet.insertChart(chart1);

    // Chart 3: Total Trips by Driver (bar) — Stacked below Chart 1
    let chart2 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colH + dataStart + ":" + colH + lastRow))
        .setOption("title", "Total Trips by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#8c564b"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Trips" })
        .setOption("width", 740)
        .setOption("height", 450)
        .setPosition(26, startCol, 0, 0).build();
    summarySheet.insertChart(chart2);

    // Chart 4: No Show by Driver (bar) — Stacked below Chart 2
    let chartNoShow = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colNoShow + dataStart + ":" + colNoShow + lastRow))
        .setOption("title", "No Show Trips by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#2ca02c"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "No Show Trips" })
        .setOption("width", 740)
        .setOption("height", 450)
        .setPosition(49, startCol, 0, 0).build();
    summarySheet.insertChart(chartNoShow);

    // Chart 5: Total Working Hours by Driver (bar) — Stacked below Chart 3
    let chartHours = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colHours + dataStart + ":" + colHours + lastRow))
        .setOption("title", "Total Working Hours by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#ff7f0e"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Hours" })
        .setOption("width", 740)
        .setOption("height", 450)
        .setPosition(72, startCol, 0, 0).build();
    summarySheet.insertChart(chartHours);

    // Chart 6: Average Per Day by Driver (bar) — Stacked below Chart 4
    let chart3 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colK + dataStart + ":" + colK + lastRow))
        .setOption("title", "Average Credit Per Day")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#d62728"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Daily Credit" })
        .setOption("width", 740)
        .setOption("height", 450)
        .setPosition(95, startCol, 0, 0).build();
    summarySheet.insertChart(chart3);
}


// =================================================================
// WEEKLY SUMMARY — Generates weekly performance breakdown
// =================================================================
function generateWeeklySummary(ss) {
    const summarySheet = ss.getSheetByName("Summary");
    let weeklySheet = ss.getSheetByName("Weekly Summary");

    if (!weeklySheet) weeklySheet = ss.insertSheet("Weekly Summary");
    weeklySheet.getRange("A:Z").clearContent();
    weeklySheet.getRange("A:Z").setBackground(null);
    weeklySheet.getRange("A:Z").setBorder(false, false, false, false, false, false);
    weeklySheet.getCharts().forEach(c => weeklySheet.removeChart(c));
    weeklySheet.getRange("A1").setValue("WEEKLY SUMMARY").setFontWeight("bold").setFontSize(14);



    // --- Read Fare % from Settings sheet ---
    let driverFareMap = {}; // driver -> fare as decimal (e.g. 0.9)
    const settingsSheet = ss.getSheetByName("Settings");
    if (settingsSheet) {
        const settingsData = settingsSheet.getDataRange().getValues();
        for (let i = 1; i < settingsData.length; i++) {
            if (settingsData[i][0]) {
                let fareVal = settingsData[i][1];
                let parsedFare = 0.9;
                if (String(fareVal).trim() === "80%-90%") {
                    parsedFare = "80%-90%";
                } else if (typeof fareVal === "string" && fareVal.includes("%")) {
                    parsedFare = parseFloat(fareVal) / 100;
                } else if (typeof fareVal === "number") {
                    parsedFare = fareVal > 1 ? fareVal / 100 : fareVal;
                }
                driverFareMap[settingsData[i][0]] = parsedFare;
            }
        }
    }

    // --- Get Main Data from Summary sheet ---
    const lastRowCurrent = summarySheet.getLastRow();
    let allRows = [];
    if (lastRowCurrent >= 3) {
        const summaryHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 9)).getDisplayValues()[0];
        const curCreditCol = findMainSummaryCol(summaryHeaders, ['credit'], 2);
        const curTripsCol = findMainSummaryCol(summaryHeaders, ['trips', 'trip'], 5);
        const curCashCol = findMainSummaryCol(summaryHeaders, ['cash'], 6);
        const maxCol = Math.max(curCreditCol, curTripsCol, curCashCol) + 1;

        const rawRows = summarySheet.getRange(3, 1, lastRowCurrent - 2, maxCol).getValues();
        let lastSeenDate = null;
        allRows = rawRows.map(r => {
            if (r[0] && r[0] !== "") lastSeenDate = r[0];
            else r[0] = lastSeenDate;
            return [
                r[0],
                r[1],
                parseNumber(r[curCreditCol]) || 0,
                parseNumber(r[curTripsCol]) || 0,
                parseNumber(r[curCashCol]) || 0
            ];
        }).filter(r => r[0] && r[1]);
    }

    const currentMonthName = ss.getName().split(" ")[0];
    const now = new Date();
    let currentYear = now.getFullYear();
    if (allRows.length > 0 && allRows[0][0] instanceof Date) {
        currentYear = allRows[0][0].getFullYear();
    } else if (currentMonthName.startsWith("Jan") && now.getMonth() === 11) {
        currentYear = currentYear + 1;
    }

    // --- Load Working Hours for Current Month ---
    const { map: workingHoursMap } = loadWorkingHoursMap(ss);

    // --- Look Back Logic ---
    const firstOfMonth = new Date(Date.parse(`${currentMonthName} 1, ${currentYear}`));
    const currentMonthIndex = firstOfMonth.getMonth();
    const firstWeekMonday = getMonday(firstOfMonth);

    if (firstWeekMonday.getMonth() < currentMonthIndex || firstWeekMonday.getFullYear() < firstOfMonth.getFullYear()) {
        try {
            const prevMonthDate = new Date(firstOfMonth);
            prevMonthDate.setMonth(currentMonthIndex - 1);
            const prevMonthName = prevMonthDate.toLocaleString('en-US', { month: 'long' });
            const prevFileName = `${prevMonthName} - Drivers Daily Balance`;
            const files = DriveApp.getFilesByName(prevFileName);

            if (files.hasNext()) {
                const prevFile = files.next();
                const prevSpreadsheet = SpreadsheetApp.openById(prevFile.getId());
                const prevSummarySheet = prevSpreadsheet.getSheetByName("Summary");
                if (prevSummarySheet && prevSummarySheet.getLastRow() >= 3) {
                    const prevHeaders = prevSummarySheet.getRange(2, 1, 1, Math.min(prevSummarySheet.getLastColumn(), 9)).getDisplayValues()[0];
                    const prevCreditCol = findMainSummaryCol(prevHeaders, ['credit'], 2);
                    const prevTripsCol = findMainSummaryCol(prevHeaders, ['trips', 'trip'], 3);
                    const prevCashCol = findMainSummaryCol(prevHeaders, ['cash'], 4);
                    const prevMaxCol = Math.max(prevCreditCol, prevTripsCol, prevCashCol) + 1;
                    const prevAllData = prevSummarySheet.getRange(3, 1, prevSummarySheet.getLastRow() - 2, prevMaxCol).getValues();

                    const currentDrivers = Array.from(new Set(allRows.map(r => r[1]).filter(n => n)));
                    let prevLastSeen = null;
                    const crossoverRows = prevAllData.map(row => {
                        if (row[0] && row[0] !== "") prevLastSeen = row[0];
                        else row[0] = prevLastSeen;

                        let mappedDriver = row[1];
                        if (mappedDriver) {
                            mappedDriver = matchOldDriverName(mappedDriver, currentDrivers);
                        }
                        return [
                            row[0],
                            mappedDriver,
                            parseNumber(row[prevCreditCol]) || 0,
                            parseNumber(row[prevTripsCol]) || 0,
                            parseNumber(row[prevCashCol]) || 0
                        ];
                    }).filter(row => {
                        const d = parseDate(row[0]);
                        return d && d >= firstWeekMonday && d < firstOfMonth;
                    });
                    allRows = [...crossoverRows, ...allRows];

                    // --- Also merge Working Hours from Previous Month ---
                    const { map: prevHoursMap } = loadWorkingHoursMap(prevSpreadsheet);
                    for (const prevDriver in prevHoursMap) {
                        const mappedDriver = matchOldDriverName(prevDriver, currentDrivers);
                        if (!workingHoursMap[mappedDriver]) {
                            workingHoursMap[mappedDriver] = {};
                        }
                        for (const dateKey in prevHoursMap[prevDriver]) {
                            workingHoursMap[mappedDriver][dateKey] = prevHoursMap[prevDriver][dateKey];
                        }
                    }
                }
            }
        } catch (e) { }
    }

    // --- Process Data ---
    const weeklyData = {};
    const dateIdx = 0, driverIdx = 1, creditIdx = 2, tripsIdx = 3, cashIdx = 4;

    allRows.forEach((row) => {
        if (!Array.isArray(row) || !row[dateIdx] || !row[driverIdx]) return;
        const date = parseDate(row[dateIdx]);
        if (!date) return;

        const driver = row[driverIdx];
        const credit = Number(row[creditIdx]) || 0;
        const trips = Number(row[tripsIdx]) || 0;
        const cash = Number(row[cashIdx]) || 0;

        const weekStart = getMonday(date);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekKey = `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;

        const dateKey = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
        const dayHours = (workingHoursMap[driver] && workingHoursMap[driver][dateKey]) ? workingHoursMap[driver][dateKey] : 0;

        if (!weeklyData[weekKey]) weeklyData[weekKey] = {};
        if (!weeklyData[weekKey][driver]) {
            weeklyData[weekKey][driver] = { credit: 0, trips: 0, cash: 0, hours: 0 };
        }

        weeklyData[weekKey][driver].credit += credit;
        weeklyData[weekKey][driver].trips += trips;
        weeklyData[weekKey][driver].cash += cash;
        weeklyData[weekKey][driver].hours += dayHours;
    });

    // --- Write Output ---
    let currentRow = 3;
    let displayedWeekCounter = 0;

    const sortedWeekKeys = Object.keys(weeklyData).sort((a, b) => {
        return parseDate(a.split(" - ")[0]) - parseDate(b.split(" - ")[0]);
    });

    sortedWeekKeys.forEach((weekKey) => {
        const weekStartDate = parseDate(weekKey.split(" - ")[0]);
        const weekEndDate = parseDate(weekKey.split(" - ")[1]);

        const isInCurrentMonth = (weekStartDate.getMonth() === currentMonthIndex && weekStartDate.getFullYear() === currentYear) ||
            (weekEndDate.getMonth() === currentMonthIndex && weekEndDate.getFullYear() === currentYear);

        if (weekStartDate && isInCurrentMonth) {
            displayedWeekCounter++;
            const weekBlock = weeklyData[weekKey];
            const startRow = currentRow;

            const titleRange = weeklySheet.getRange(currentRow, 1, 1, 7).merge();
            titleRange.setValue(`WEEK ${displayedWeekCounter} (${weekKey})`)
                .setFontWeight("bold")
                .setFontSize(11)
                .setBackground("#d0e0e3");

            currentRow++;

            const headerRange = weeklySheet.getRange(currentRow, 1, 1, 7);
            headerRange.setValues([["Driver", "Total Credit", "Trips", "Total Cash", "Total Hours", "Avg Per Hour", "Balance"]])
                .setFontWeight("bold")
                .setBackground("#3c78d8")
                .setFontColor("white")
                .setHorizontalAlignment("center");

            currentRow++;

            const drivers = Object.keys(weekBlock);
            const tableData = [];

            drivers.forEach(d => {
                const info = weeklyData[weekKey][d];
                const cash = info.cash;
                const fare = driverFareMap[d] !== undefined ? driverFareMap[d] : 0.9;

                let grossPayout = 0;
                if (fare === "80%-90%") {
                    if (info.credit <= 1000) {
                        grossPayout = info.credit * 0.80;
                    } else {
                        grossPayout = (1000 * 0.80) + ((info.credit - 1000) * 0.90);
                    }
                } else {
                    const numericFare = Number(fare) || 0.9;
                    grossPayout = info.credit * numericFare;
                }

                const balance = roundToTwo(grossPayout - cash);
                const avgPerHour = info.hours > 0 ? roundToTwo(info.credit / info.hours) : 0;
                tableData.push([d, info.credit, info.trips, cash, roundToTwo(info.hours), avgPerHour, balance]);
            });

            // Write Main Table
            const dataRange = weeklySheet.getRange(currentRow, 1, tableData.length, 7);
            dataRange.setValues(tableData);
            dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            weeklySheet.getRange(currentRow, 2, tableData.length, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow, 4, tableData.length, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow, 5, tableData.length, 1).setNumberFormat("0.0");
            weeklySheet.getRange(currentRow, 6, tableData.length, 2).setNumberFormat("$#,##0.00");

            const totalCredit = tableData.reduce((a, b) => a + b[1], 0);
            const totalTrips = tableData.reduce((a, b) => a + b[2], 0);
            const totalCash = tableData.reduce((a, b) => a + b[3], 0);
            const totalHours = tableData.reduce((a, b) => a + b[4], 0);
            const totalAvgPerHour = totalHours > 0 ? roundToTwo(totalCredit / totalHours) : 0;
            const totalBalance = tableData.reduce((a, b) => a + b[6], 0);

            const totalRowRange = weeklySheet.getRange(currentRow + tableData.length, 1, 1, 7);
            totalRowRange.setValues([["TOTAL", totalCredit, totalTrips, totalCash, roundToTwo(totalHours), totalAvgPerHour, totalBalance]])
                .setFontWeight("bold")
                .setBackground("#eeeeee")
                .setBorder(true, true, true, true, true, true);

            weeklySheet.getRange(currentRow + tableData.length, 2, 1, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow + tableData.length, 4, 1, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow + tableData.length, 5, 1, 1).setNumberFormat("0.0");
            weeklySheet.getRange(currentRow + tableData.length, 6, 1, 2).setNumberFormat("$#,##0.00");

            const tableEnd = currentRow + tableData.length;

            // Chart 1: Total Credit by Driver (Col I = col 9, width 480px, height 310px)
            const chart1 = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 9, 0, 0)
                .addRange(weeklySheet.getRange(currentRow - 1, 1, tableData.length + 1, 2))
                .setNumHeaders(1)
                .setOption("title", `WEEK ${displayedWeekCounter} - Total Credit`)
                .setOption("colors", ["#1f77b4"])
                .setOption("legend", { position: "none" })
                .setOption("hAxis", { title: "Driver" })
                .setOption("vAxis", { title: "Total Credit" })
                .setOption("width", 480)
                .setOption("height", 310)
                .build();
            weeklySheet.insertChart(chart1);

            // Chart 2: Trips by Driver (Col N = col 14, width 480px, height 310px)
            const chart2 = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 14, 0, 0)
                .addRange(weeklySheet.getRange(currentRow - 1, 1, tableData.length + 1, 1)) // Driver Names
                .addRange(weeklySheet.getRange(currentRow - 1, 3, tableData.length + 1, 1)) // Trips
                .setNumHeaders(1)
                .setOption("title", `WEEK ${displayedWeekCounter} - Trips`)
                .setOption("colors", ["#8c564b"]) // Distinct brown color
                .setOption("legend", { position: "none" })
                .setOption("hAxis", { title: "Driver" })
                .setOption("vAxis", { title: "Trips" })
                .setOption("width", 480)
                .setOption("height", 310)
                .build();
            weeklySheet.insertChart(chart2);

            currentRow = Math.max(tableEnd + 2, startRow + 16);
        }
    });
}


// =================================================================
// BONUS REPORT — Identifies drivers who hit $1495 weekly threshold
// =================================================================
/**
 * Creates the "Bonus" sheet with professional styling.
 * Uses Summary sheet data (which is now built from Raw Data).
 */
function generateBonusReport(ss) {
    Logger.log("--- Starting generateBonusReport (Styled) ---");

    const summarySheet = ss.getSheetByName("Summary");
    const bonusSheetName = "Bonus";

    // 1. --- Get Main Data from Summary (WITH DATE FILL-DOWN) ---
    const lastRowCurrent = summarySheet.getLastRow();
    let allRows = [];
    if (lastRowCurrent >= 3) {
        const rawRows = summarySheet.getRange("A3:D" + lastRowCurrent).getValues();

        let lastSeenDate = null;
        allRows = rawRows.map(r => {
            if (r[0] && r[0] !== "") lastSeenDate = r[0];
            else r[0] = lastSeenDate;
            return r;
        }).filter(r => r[0] && r[1]);
    }

    // --- FETCH NON-DISPATCH DRIVERS FROM "SETTINGS" SHEET ---
    let nonDispatchDrivers = new Set();
    const settingsSheet = ss.getSheetByName("Settings");
    if (settingsSheet) {
        const settingsData = settingsSheet.getDataRange().getValues();
        for (let i = 1; i < settingsData.length; i++) {
            const dName = settingsData[i][0];
            const dSystem = (settingsData[i][1] || "").toString().toUpperCase();

            if (dSystem === "90%" || dSystem.includes("NON-DISPATCH") || dSystem === "90%-95%") {
                nonDispatchDrivers.add(dName);
            }
        }
    }

    // 2. --- Year Detection ---
    const currentMonthName = ss.getName().split(" ")[0];
    const now = new Date();
    let currentYear = now.getFullYear();
    if (allRows.length > 0 && allRows[0][0] instanceof Date) {
        currentYear = allRows[0][0].getFullYear();
    } else if (currentMonthName.startsWith("Jan") && now.getMonth() === 11) {
        currentYear = currentYear + 1;
    }

    // 3. --- Look Back Logic (WITH DATE FILL-DOWN) ---
    const firstOfMonth = new Date(Date.parse(`${currentMonthName} 1, ${currentYear}`));
    const currentMonthIndex = firstOfMonth.getMonth();
    const firstWeekMonday = getMonday(firstOfMonth);

    if (firstWeekMonday.getMonth() < currentMonthIndex || firstWeekMonday.getFullYear() < firstOfMonth.getFullYear()) {
        try {
            const prevMonthDate = new Date(firstOfMonth);
            prevMonthDate.setMonth(currentMonthIndex - 1);
            const prevMonthName = prevMonthDate.toLocaleString('en-US', { month: 'long' });
            const prevFileName = `${prevMonthName} - Drivers Daily Balance`;
            const files = DriveApp.getFilesByName(prevFileName);
            if (files.hasNext()) {
                const prevFile = files.next();
                const prevSpreadsheet = SpreadsheetApp.openById(prevFile.getId());
                const prevSummarySheet = prevSpreadsheet.getSheetByName("Summary");
                if (prevSummarySheet && prevSummarySheet.getLastRow() >= 3) {
                    const oldLastCol = prevSummarySheet.getLastColumn();
                    const readCols = oldLastCol >= 4 ? 4 : oldLastCol;
                    const prevAllData = prevSummarySheet.getRange(3, 1, prevSummarySheet.getLastRow() - 2, readCols).getValues();

                    const currentDrivers = Array.from(new Set(allRows.map(r => r[1]).filter(n => n)));
                    let prevLastSeen = null;
                    const crossoverRows = prevAllData.map(row => {
                        if (row[0] && row[0] !== "") prevLastSeen = row[0];
                        else row[0] = prevLastSeen;

                        if (row[1]) {
                            row[1] = matchOldDriverName(row[1], currentDrivers);
                        }
                        return row;
                    }).filter(row => {
                        const d = parseDate(row[0]);
                        return d && d >= firstWeekMonday && d < firstOfMonth;
                    });
                    allRows = [...crossoverRows, ...allRows];
                }
            }
        } catch (e) { Logger.log("Bonus Script Error fetching previous month: " + e.toString()); }
    }

    // 4. --- Process Data & Find Qualifiers ---
    const weeklyData = {};
    const dateIdx = 0, driverIdx = 1, creditIdx = 2;

    allRows.forEach((row) => {
        const date = parseDate(row[dateIdx]);
        if (!date) return;

        const driver = row[driverIdx];
        const credit = Number(row[creditIdx]) || 0;

        const weekStart = getMonday(date);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekKey = `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;

        if (!weeklyData[weekKey]) weeklyData[weekKey] = {};
        if (!weeklyData[weekKey][driver]) {
            weeklyData[weekKey][driver] = { credit: 0 };
        }
        weeklyData[weekKey][driver].credit += credit;
    });

    const includeNonDispatchInBonus = true;

    let allQualifiedDrivers = [];

    Object.keys(weeklyData).forEach(weekKey => {
        const weekEndDate = parseDate(weekKey.split(" - ")[1]);

        if (weekEndDate.getMonth() === currentMonthIndex && weekEndDate.getFullYear() === currentYear) {

            const weekBlock = weeklyData[weekKey];
            for (const driver in weekBlock) {

                const info = weekBlock[driver];
                const roundedCredit = Math.round(info.credit);

                if (nonDispatchDrivers.has(driver)) {
                    if (includeNonDispatchInBonus === true && roundedCredit >= 1495) {
                        allQualifiedDrivers.push([driver, roundedCredit, weekKey]);
                    }
                    continue;
                }

                if (roundedCredit >= 1495) {
                    allQualifiedDrivers.push([driver, roundedCredit, weekKey]);
                }
            }
        }
    });

    // 5. --- Write to Bonus Sheet ---
    let bonusSheet = ss.getSheetByName(bonusSheetName);
    if (!bonusSheet) {
        bonusSheet = ss.insertSheet(bonusSheetName);
    }

    bonusSheet.clear();
    bonusSheet.getRange("A:C").setBackground(null).setBorder(false, false, false, false, false, false);

    if (allQualifiedDrivers.length > 0) {
        allQualifiedDrivers.sort((a, b) => {
            const dateA = parseDate(a[2].split(" - ")[0]);
            const dateB = parseDate(b[2].split(" - ")[0]);
            if (dateA < dateB) return -1;
            if (dateA > dateB) return 1;
            return a[0].localeCompare(b[0]);
        });

        const headerRange = bonusSheet.getRange("A1:C1");
        headerRange.setValues([["Driver", "Total Credit", "Week Period"]])
            .setFontWeight("bold")
            .setFontColor("white")
            .setBackground("#3c78d8")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontSize(11);

        const dataRange = bonusSheet.getRange(2, 1, allQualifiedDrivers.length, 3);
        dataRange.setValues(allQualifiedDrivers)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

        for (let i = 0; i < allQualifiedDrivers.length; i++) {
            if (i % 2 === 1) {
                bonusSheet.getRange(2 + i, 1, 1, 3).setBackground("#f3f3f3");
            }
        }

        bonusSheet.getRange(2, 2, allQualifiedDrivers.length, 1).setNumberFormat("$#,##0");

        const totalBonusCredit = allQualifiedDrivers.reduce((sum, row) => sum + row[1], 0);
        const halfPercentBonus = totalBonusCredit * 0.005;
        const totalRow = bonusSheet.getLastRow() + 2;

        const totalsBlock = bonusSheet.getRange(totalRow, 1, 2, 2);

        bonusSheet.getRange(totalRow, 1, 2, 1).setValues([["Grand Total"], ["0.5% Bonus"]]);
        bonusSheet.getRange(totalRow, 2, 2, 1).setValues([[totalBonusCredit], [halfPercentBonus]]);

        totalsBlock.setFontWeight("bold")
            .setBackground("#cfe2f3")
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
            .setHorizontalAlignment("center");

        bonusSheet.getRange(totalRow, 2, 2, 1).setNumberFormat("$#,##0");

    } else {
        bonusSheet.getRange("A1").setValue("No Bonuses for this period.").setFontWeight("bold");
    }

    bonusSheet.autoResizeColumns(1, 3);
    bonusSheet.setColumnWidth(1, bonusSheet.getColumnWidth(1) + 20);
    bonusSheet.setColumnWidth(2, bonusSheet.getColumnWidth(2) + 20);
    bonusSheet.setColumnWidth(3, bonusSheet.getColumnWidth(3) + 20);
}


// =================================================================
// SHARED HELPERS
// =================================================================

/**
 * Tries to map an old driver name (e.g. "Ablaye Diop") to a new driver name (e.g. "Ablaye").
 * Handles first name matching while avoiding conflicting last names.
 */
function matchOldDriverName(oldName, newNames) {
    if (!oldName) return oldName;
    const oldStr = oldName.toString().trim();
    if (newNames.includes(oldStr)) return oldStr; // exact match

    const oldParts = oldStr.split(" ");
    const oldFirst = oldParts[0].toLowerCase();
    const oldLast = oldParts.length > 1 ? oldParts.slice(1).join(" ").toLowerCase() : "";

    // Find all new names with the same first name
    const matches = newNames.filter(n => {
        const nParts = n.toString().trim().split(" ");
        return nParts[0].toLowerCase() === oldFirst;
    });

    if (matches.length === 1) {
        const nParts = matches[0].toString().trim().split(" ");
        const nLast = nParts.length > 1 ? nParts.slice(1).join(" ").toLowerCase() : "";
        // If one has a last name and the other does too, they must not conflict.
        if (oldLast && nLast) {
            // E.g. "Diop" vs "Smith" -> conflict. "Diop" vs "Di" -> ok.
            if (!oldLast.startsWith(nLast) && !nLast.startsWith(oldLast)) {
                return oldStr; // conflicting last names, don't match
            }
        }
        return matches[0];
    }

    // If multiple matches or no matches, return original
    return oldStr;
}

/**
 * Loads No Show data (Date, Driver, Cash) from the No Show sheet.
 * Returns: { driverName -> { dateKey -> { cash: X, count: Y } } }
 */
function loadNoShowMap(ss, sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    const map = {};
    if (!sheet) return map;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return map;

    const headers = data[0].map(h => (h || '').toString().toLowerCase().trim());
    let dateCol = headers.indexOf('date');
    let driverCol = headers.findIndex(h => h.includes('driver') || h.includes('name'));
    let cashCol = headers.findIndex(h => h.includes('cash') || h.includes('amount'));

    if (dateCol === -1) dateCol = 0;
    if (driverCol === -1) driverCol = 1;
    if (cashCol === -1) cashCol = 2;

    for (let r = 1; r < data.length; r++) {
        const dateVal = parseDDMMYY(data[r][dateCol]);
        const driver = (data[r][driverCol] || '').toString().trim();
        const cash = parseNumber(data[r][cashCol]);

        if (dateVal && driver && cash !== null) {
            if (!map[driver]) map[driver] = {};
            const dKey = dateVal.getTime();
            if (!map[driver][dKey]) map[driver][dKey] = { cash: 0, count: 0 };
            map[driver][dKey].cash += cash; // sum cash for deduction
            map[driver][dKey].count += 1;   // count trips for display
        }
    }
    return map;
}

/**
 * Parses DD-MM-YY date format from Raw Data headers.
 * Example: "01-06-26" → June 1, 2026
 */
function parseDDMMYY(value) {
    if (!value) return null;

    // Handle native Google Sheets Date objects directly
    if (value instanceof Date) {
        return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }

    const str = String(value).trim();

    // Handle DD-MM-YY format (e.g., "01-06-26")
    const dashParts = str.split("-");
    if (dashParts.length === 3) {
        const day = parseInt(dashParts[0], 10);
        const month = parseInt(dashParts[1], 10) - 1; // zero-based
        let year = parseInt(dashParts[2], 10);
        if (year < 100) year += 2000; // 26 → 2026
        if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
        const d = new Date(year, month, day);
        return isNaN(d.getTime()) ? null : d;
    }

    // Handle DD/MM/YYYY format (e.g., "01/06/2026")
    const slashParts = str.split("/");
    if (slashParts.length === 3) {
        const day = parseInt(slashParts[0], 10);
        const month = parseInt(slashParts[1], 10) - 1;
        let year = parseInt(slashParts[2], 10);
        if (year < 100) year += 2000;
        if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
        const d = new Date(year, month, day);
        return isNaN(d.getTime()) ? null : d;
    }

    // Final fallback: try standard JS Date parser for strings like "Thursday, June 25, 2026"
    const fallback = new Date(str);
    if (!isNaN(fallback.getTime())) {
        return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
    }

    return null;
}

// =================================================================
// ANGEL'S PROGRESSIVE REPORT
// =================================================================
function generateAngelReport(ss) {
    if (!ss) ss = SpreadsheetApp.getActiveSpreadsheet();
    const summarySheet = ss.getSheetByName("Summary");
    const targetDriver = "Angel";
    const reportSheetName = "Angel Summary";

    if (!summarySheet) return;

    // 1. --- Detect Month and Year ---
    const ssName = ss.getName();
    const monthMatch = ssName.match(/^(January|February|March|April|May|June|July|August|September|October|November|December)/i);
    const currentMonthName = monthMatch ? monthMatch[1] : null;

    if (!currentMonthName) return;

    const lastRow = summarySheet.getLastRow();
    let currentYear = new Date().getFullYear();
    if (lastRow >= 3) {
        const firstDateVal = summarySheet.getRange("A3").getValue();
        if (firstDateVal instanceof Date) {
            currentYear = firstDateVal.getFullYear();
        }
    }

    const firstOfMonth = new Date(Date.parse(`${currentMonthName} 1, ${currentYear}`));
    if (isNaN(firstOfMonth.getTime())) return;
    const endDate = new Date(currentYear, firstOfMonth.getMonth() + 1, 0);
    const reportStartDate = getMonday(firstOfMonth);

    // 2. --- Initialize Data Map ---
    let angelDataMap = {};

    function mapRowToAngelData(row, colIndices) {
        const driverName = String(row[1] || "");
        if (driverName.includes(targetDriver) && row[0] instanceof Date) {
            let dateKey = row[0].toDateString();
            const cIdx = colIndices ? colIndices.credit : 2;
            const tIdx = colIndices ? colIndices.trips : 3;
            const caIdx = colIndices ? colIndices.cash : 4;
            const nsIdx = colIndices ? colIndices.noShow : 5;
            angelDataMap[dateKey] = {
                credit: Number(row[cIdx]) || 0,
                trips: Number(row[tIdx]) || 0,
                cash: Number(row[caIdx]) || 0,
                noShow: Number(row[nsIdx]) || 0
            };
        }
    }

    // 3. --- Fetch Previous Month Data (Crossover Week) ---
    if (reportStartDate < firstOfMonth) {
        try {
            const prevMonthDate = new Date(firstOfMonth);
            prevMonthDate.setMonth(firstOfMonth.getMonth() - 1);
            const prevMonthName = prevMonthDate.toLocaleString('en-US', { month: 'long' });
            const prevFileName = `${prevMonthName} - Drivers Daily Balance`;

            const files = DriveApp.getFilesByName(prevFileName);
            if (files.hasNext()) {
                const prevFile = files.next();
                const prevSpreadsheet = SpreadsheetApp.openById(prevFile.getId());
                const prevSummarySheet = prevSpreadsheet.getSheetByName("Summary");

                if (prevSummarySheet && prevSummarySheet.getLastRow() >= 3) {
                    const prevHeaders = prevSummarySheet.getRange(2, 1, 1, Math.min(prevSummarySheet.getLastColumn(), 9)).getDisplayValues()[0];
                    const prevIndices = {
                        credit: findMainSummaryCol(prevHeaders, ['credit'], 2),
                        trips: findMainSummaryCol(prevHeaders, ['trips', 'trip'], 3),
                        cash: findMainSummaryCol(prevHeaders, ['cash'], 4),
                        noShow: findMainSummaryCol(prevHeaders, ['no show', 'noshow'], 5)
                    };
                    const prevData = prevSummarySheet.getRange(3, 1, prevSummarySheet.getLastRow() - 2, 11).getValues();

                    let prevLastSeenDate = null;
                    prevData.forEach(row => {
                        if (row[0] && row[0] !== "") prevLastSeenDate = row[0];
                        else row[0] = prevLastSeenDate;

                        const d = parseDate(row[0]);
                        if (d && d >= reportStartDate) {
                            row[0] = d;
                            mapRowToAngelData(row, prevIndices);
                        }
                    });
                }
            }
        } catch (e) { }
    }

    // 4. --- Fetch Current Month Data ---
    if (lastRow >= 3) {
        const curHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 9)).getDisplayValues()[0];
        const curIndices = {
            credit: findMainSummaryCol(curHeaders, ['credit'], 2),
            trips: findMainSummaryCol(curHeaders, ['trips', 'trip'], 5),
            cash: findMainSummaryCol(curHeaders, ['cash'], 6),
            noShow: findMainSummaryCol(curHeaders, ['no show', 'noshow'], 7)
        };
        const currentData = summarySheet.getRange(3, 1, lastRow - 2, 11).getValues();

        let currentLastSeenDate = null;
        currentData.forEach(row => {
            if (row[0] && row[0] !== "") currentLastSeenDate = row[0];
            else row[0] = currentLastSeenDate;

            if (row[0]) {
                const d = parseDate(row[0]);
                if (d) {
                    row[0] = d;
                    mapRowToAngelData(row, curIndices);
                }
            }
        });
    }

    // 5. --- Build the Report Rows ---
    let reportRows = [];
    let chartDataRows = [];
    let weekCounter = 1;

    let weeklyCredit = 0, weeklyNoShow = 0, weeklyTrips = 0, weeklyCash = 0;
    let weeklyGrossBal = 0, weeklyFee = 0, weeklyNet = 0;

    let grandCredit = 0, grandNoShow = 0, grandTrips = 0, grandCash = 0;
    let grandGrossBal = 0, grandFee = 0, grandNet = 0;

    let weekTotalRows = [];

    const threshold = 1000;
    const baseRate = 0.80;
    const topRate = 0.90;
    const transferFeeRate = 0.03;

    for (let d = new Date(reportStartDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        const dateKey = d.toDateString();
        const dayData = angelDataMap[dateKey] || { credit: 0, cash: 0, noShow: 0, trips: 0 };

        const todayCredit = dayData.credit;
        const todayCash = dayData.cash;
        let todayGrossPayout = 0;
        let appliedRate = 0;

        if (todayCredit > 0) {
            if (weeklyCredit >= threshold) {
                todayGrossPayout = todayCredit * topRate;
            }
            else if ((weeklyCredit + todayCredit) <= threshold) {
                todayGrossPayout = todayCredit * baseRate;
            }
            else {
                const creditAtBaseRate = threshold - weeklyCredit;
                const creditAtTopRate = todayCredit - creditAtBaseRate;
                todayGrossPayout = (creditAtBaseRate * baseRate) + (creditAtTopRate * topRate);
            }
            appliedRate = todayGrossPayout / todayCredit;
        }

        weeklyCredit += todayCredit;
        const grossBalance = todayGrossPayout - todayCash;
        const fee = grossBalance > 0 ? (grossBalance * transferFeeRate) : 0;
        const netPayout = grossBalance - fee;

        reportRows.push([
            new Date(d),
            todayCredit,
            weeklyCredit,
            appliedRate,
            dayData.noShow,
            dayData.trips,
            todayCash,
            grossBalance,
            fee,
            netPayout
        ]);

        chartDataRows.push([new Date(d), todayCredit]);

        weeklyNoShow += dayData.noShow;
        weeklyTrips += dayData.trips;
        weeklyCash += todayCash;
        weeklyGrossBal += grossBalance;
        weeklyFee += fee;
        weeklyNet += netPayout;

        const isSunday = d.getDay() === 0;
        const isLastDayOfReport = d.toDateString() === endDate.toDateString();

        if (isSunday || isLastDayOfReport) {
            grandCredit += weeklyCredit;
            grandNoShow += weeklyNoShow;
            grandTrips += weeklyTrips;
            grandCash += weeklyCash;
            grandGrossBal += weeklyGrossBal;
            grandFee += weeklyFee;
            grandNet += weeklyNet;

            let weekLabel = `Week ${weekCounter} Total`;
            let rowColor = "#e6f4ea";
            if (weeklyCredit > threshold) rowColor = "#fff2cc";

            let weeklyEffectiveRate = weeklyCredit > 0 ? ((weeklyGrossBal + weeklyCash) / weeklyCredit) : 0;;

            reportRows.push([
                weekLabel,
                weeklyCredit,
                "",
                weeklyEffectiveRate,
                weeklyNoShow,
                weeklyTrips,
                weeklyCash,
                weeklyGrossBal,
                weeklyFee,
                weeklyNet
            ]);

            weekTotalRows.push({
                index: reportRows.length + 1,
                color: rowColor
            });

            weeklyCredit = 0; weeklyNoShow = 0; weeklyTrips = 0; weeklyCash = 0;
            weeklyGrossBal = 0; weeklyFee = 0; weeklyNet = 0;

            if (!isLastDayOfReport) {
                weekCounter++;
                reportRows.push(["", "", "", "", "", "", "", "", "", ""]);
            }
        }
    }

    reportRows.push(["", "", "", "", "", "", "", "", "", ""]);
    const grandRate = grandCredit > 0 ? ((grandGrossBal + grandCash) / grandCredit) : 0;
    reportRows.push([
        "GRAND TOTAL", grandCredit, "", grandRate, grandNoShow, grandTrips, grandCash, grandGrossBal, grandFee, grandNet
    ]);
    const grandTotalRowIndex = reportRows.length + 1;

    let reportSheet = ss.getSheetByName(reportSheetName);
    if (!reportSheet) {
        reportSheet = ss.insertSheet(reportSheetName);
    } else {
        reportSheet.clear();
        const charts = reportSheet.getCharts();
        charts.forEach(c => reportSheet.removeChart(c));
        try { reportSheet.showColumns(1, 20); } catch (e) { }
    }

    const headers = [["Date", "Daily Credit", "Weekly Cum.", "Applied Rate", "No Show", "Trips", "Cash", "Gross Bal.", "Fee (3%)", "Net Payout"]];

    reportSheet.getRange("A1:J1").setValues(headers)
        .setFontWeight("bold")
        .setFontSize(11)
        .setHorizontalAlignment("center")
        .setBackground("#f3f3f3")
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    if (reportRows.length > 0) {
        const range = reportSheet.getRange(2, 1, reportRows.length, 10);
        range.setValues(reportRows);

        reportSheet.getRange(2, 1, reportRows.length, 1).setNumberFormat("dddd, dd");
        reportSheet.getRange(2, 2, reportRows.length, 2).setNumberFormat("$#,##0.00");
        reportSheet.getRange(2, 4, reportRows.length, 1).setNumberFormat("0.0%");
        reportSheet.getRange(2, 5, reportRows.length, 1).setNumberFormat("0");
        reportSheet.getRange(2, 6, reportRows.length, 1).setNumberFormat("0");
        reportSheet.getRange(2, 7, reportRows.length, 4).setNumberFormat("$#,##0.00");

        range.setHorizontalAlignment("center");

        weekTotalRows.forEach(item => {
            const rowRange = reportSheet.getRange(item.index, 1, 1, 10);
            rowRange.setFontWeight("bold");
            rowRange.setBackground(item.color);
            rowRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
            reportSheet.getRange(item.index, 1).setHorizontalAlignment("center");
        });

        const grandRowRange = reportSheet.getRange(grandTotalRowIndex, 1, 1, 10);
        grandRowRange.setFontWeight("bold");
        grandRowRange.setFontSize(12);
        grandRowRange.setBackground("#d9d2e9");
        grandRowRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

        reportSheet.autoResizeColumns(1, 10);
        for (let i = 1; i <= 10; i++) {
            let currentWidth = reportSheet.getColumnWidth(i);
            reportSheet.setColumnWidth(i, currentWidth + 15);
        }
        if (reportSheet.getColumnWidth(1) < 120) reportSheet.setColumnWidth(1, 120);

        if (chartDataRows.length > 0) {
            const chartDataRange = reportSheet.getRange(2, 12, chartDataRows.length, 2);
            chartDataRange.setValues(chartDataRows);
            reportSheet.getRange(2, 12, chartDataRows.length, 1).setNumberFormat("dddd, dd");
            reportSheet.getRange(2, 13, chartDataRows.length, 1).setNumberFormat("$#,##0");
            reportSheet.hideColumns(12, 2);
        }

        let angelChart = reportSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(reportSheet.getRange(2, 12, chartDataRows.length, 1))
            .addRange(reportSheet.getRange(2, 13, chartDataRows.length, 1))
            .setPosition(2, 12, 0, 0)
            .setOption("title", "Angel's Daily Credit Progress")
            .setOption("titleTextStyle", { bold: true, fontSize: 18 })
            .setOption("vAxis", { title: "Daily Credit Earned", format: '$#,##0' })
            .setOption("hAxis", { title: "Date", format: 'dddd, dd' })
            .setOption("colors", ["#1f77b4"])
            .setOption("legend", { position: "none" })
            .setOption("pointSize", 5)
            .setOption("width", 800)
            .setOption("height", 400)
            .build();

        reportSheet.insertChart(angelChart);
    }
}


function parseNumber(val) {
    if (val === "" || val === null || val === undefined) return null;
    if (typeof val === "number") return val;
    if (typeof val === "string" && val.includes("%")) {
        return parseFloat(val) / 100;
    }
    const s = String(val).replace(/,/g, "").trim();
    if (s === "") return null;
    const n = Number(s);
    return isNaN(n) ? null : n;
}

function roundToTwo(num) {
    return Math.round((num + Number.EPSILON) * 100) / 100;
}

function getMonday(date) {
    const d = new Date(date);
    const day = d.getDay();
    const diff = (day === 0 ? -6 : 1 - day);
    d.setDate(d.getDate() + diff);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

function formatDate(date) {
    const mm = (date.getMonth() + 1).toString().padStart(2, "0");
    const dd = date.getDate().toString().padStart(2, "0");
    const yyyy = date.getFullYear();
    return `${mm}/${dd}/${yyyy}`;
}

function parseDate(value) {
    if (!value) return null;
    if (value instanceof Date) return value;
    if (typeof value === "string") {
        const parts = value.split("/");
        if (parts.length === 3) {
            const month = parseInt(parts[0], 10) - 1;
            const day = parseInt(parts[1], 10);
            const year = parseInt(parts[2], 10);
            const d = new Date(year, month, day);
            return isNaN(d.getTime()) ? null : d;
        }
    }
    return null;
}