// =================================================================
// MASTER SCRIPT — CENTRALIZED BALANCE SHEET CONTROLLER
// =================================================================
// Change this ID each time you switch to a new monthly balance file.
// This is the ONLY line you need to update.
const TARGET_SHEET_ID = "14jjwIW2ywsEXz0JAgxcaG43FKa4Nq8pezaTspigvbDE";

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
            .addItem('☕ Sync Breaks Table', 'syncBreaksTableFromMenu')
            .addItem('🕒 Sync Working Hours Sheet', 'generateWorkingHoursSheetFromMenu')
            .addSeparator()
            .addItem('📥 Import: NET Data  (from "Raw Data - NET" sheet)', 'importNetData')
            .addItem('📥 Import: Trips Data  (from "Raw Data - Trips" sheet)', 'importTripsData')
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
    const maxCol = Math.min(headers.length, 10);
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
                const d = parseSheetDateHeader(headerRow[c], data[headerRowIdx][c]);
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
                        let weeklyTotal = parseNumber(data[dr][totalColIdx]);
                        if (weeklyTotal === null) weeklyTotal = parseNumber(displayValues[dr][totalColIdx]) || 0;

                        let weeklyTrips = parseNumber(data[dr][countColIdx]);
                        if (weeklyTrips === null) weeklyTrips = parseNumber(displayValues[dr][countColIdx]) || 0;

                        let weeklyNet = parseNumber(data[dr][netColIdx]);
                        if (weeklyNet === null) weeklyNet = parseNumber(displayValues[dr][netColIdx]) || 0;

                        const weeklyCash = roundToTwo(weeklyTotal - weeklyNet);

                        const dailyCredits = {};
                        dateColumns.forEach(dc => {
                            let cred = parseNumber(data[dr][dc.colIdx]);
                            if (cred === null) cred = parseNumber(displayValues[dr][dc.colIdx]) || 0;
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

function runSyncBreaksTable() {
    Logger.log("=== Starting runSyncBreaksTable ===");
    let ss = null;
    try {
        ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch (e) { }
    if (!ss) {
        ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    }
    Logger.log("Target Spreadsheet: " + ss.getName() + " (" + ss.getId() + ")");
    const info = syncBreaksTable(ss);
    SpreadsheetApp.flush();
    Logger.log("=== Successfully synced breaks table at row " + info.titleRow + " ===");
    return info;
}

function generateWorkingHoursSheetFromMenu() {
    let ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { }
    if (!ss) ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    generateWorkingHoursSheet(ss);
    syncBreaksTable(ss);
    SpreadsheetApp.flush();
    try {
        ss.toast("Working hours and breaks table synced!", "Balance Sheet", 3);
    } catch (e) { }
}

function syncBreaksTableFromMenu() {
    let ss = null;
    try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { }
    if (!ss) ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    try {
        runSyncBreaksTable();
        try {
            ss.toast("Breaks table synced successfully!", "Balance Sheet", 3);
        } catch (tErr) { }
    } catch (e) {
        Logger.log("Error syncing breaks table: " + e.message + "\n" + (e.stack || ""));
        try {
            ss.toast("Error syncing breaks: " + e.message, "Balance Sheet", 5);
        } catch (tErr) { }
    }
}

/**
 * Safely parses a date header from a sheet (display value string + raw cell value).
 * Prioritizes DD-MM-YY display string to prevent Google Sheets US-locale Date object distortion.
 */
function parseSheetDateHeader(displayHeader, rawHeader) {
    if (displayHeader) {
        const disp = String(displayHeader).trim();
        if (disp && !disp.toLowerCase().includes("total") && !disp.toLowerCase().includes("count") && !disp.toLowerCase().includes("net") && !disp.toLowerCase().includes("trip")) {
            const d = parseDDMMYY(disp);
            if (d && !isNaN(d.getTime())) return d;
        }
    }
    if (rawHeader instanceof Date) {
        return new Date(rawHeader.getFullYear(), rawHeader.getMonth(), rawHeader.getDate());
    }
    if (rawHeader) {
        const d = parseDDMMYY(rawHeader);
        if (d && !isNaN(d.getTime())) return d;
    }
    return null;
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

    const lastCol = hoursSheet.getLastColumn();
    const rawHeaders = hoursSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const displayHeaders = hoursSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

    const dateCols = {}; // colIdx -> timestamp
    for (let c = 1; c < lastCol; c++) {
        const hText = (displayHeaders[c] || '').toString().toLowerCase().trim();
        if (hText.includes("total")) continue;

        const d = parseSheetDateHeader(displayHeaders[c], rawHeaders[c]);
        if (d && !isNaN(d.getTime())) dateCols[c] = d.getTime();
    }

    // Seed canonical drivers from Settings if available
    const settingsSheet = ss.getSheetByName("Settings");
    const canonicalDrivers = [];
    if (settingsSheet) {
        const sData = settingsSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
            const dName = (sData[i][0] || "").toString().trim();
            if (dName && !canonicalDrivers.includes(dName)) canonicalDrivers.push(dName);
        }
    }

    const map = {};
    const driverTotalHours = {};
    const dailyTotalHours = {};

    for (let r = 1; r < data.length; r++) {
        const rawDriver = (data[r][0] || "").toString().trim();
        if (rawDriver.toLowerCase().startsWith("total") || rawDriver.toLowerCase().includes("breaks")) break;
        if (!rawDriver) continue;

        const canonDriver = getCanonicalDriverName(rawDriver, canonicalDrivers);
        if (!map[rawDriver]) map[rawDriver] = {};
        if (!map[canonDriver]) map[canonDriver] = {};

        let driverSum = 0;
        const seenDatesForRow = {};

        for (const c in dateCols) {
            const dateKey = dateCols[c];
            const val = parseNumber(data[r][c]);
            if (val !== null && !isNaN(val)) {
                if (seenDatesForRow[dateKey] === undefined || val > seenDatesForRow[dateKey]) {
                    seenDatesForRow[dateKey] = val;
                }
            }
        }

        for (const dateKey in seenDatesForRow) {
            const val = seenDatesForRow[dateKey];
            map[rawDriver][dateKey] = val;
            map[canonDriver][dateKey] = val;
            if (val > 0) {
                driverSum += val;
                dailyTotalHours[dateKey] = (dailyTotalHours[dateKey] || 0) + val;
            }
        }

        driverTotalHours[rawDriver] = roundToTwo(driverSum);
        driverTotalHours[canonDriver] = roundToTwo(driverSum);
    }

    return { map, driverTotalHours, dailyTotalHours };
}

/**
 * Generates/syncs the "Working hours" sheet based on "Raw Data".
 * ADDITIVE ONLY: Never erases or clears existing entered hours.
 * Automatically de-duplicates any duplicate columns/rows from previous runs.
 * Only adds new date columns and new driver rows when they appear in Raw Data.
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

    const existingLastRow = hoursSheet.getLastRow();
    const existingLastCol = hoursSheet.getLastColumn();
    const isExistingSheet = existingLastRow >= 2 && existingLastCol >= 2;

    if (!isExistingSheet) {
        // --- INITIAL CREATION (Sheet is brand new or empty) ---
        const headerRow = ["Driver"];
        allDateCols.forEach(dc => headerRow.push(dc.headerStr));
        headerRow.push("Total Hours");

        const numDateCols = allDateCols.length;
        const lastDateColLetter = getColumnLetter(numDateCols + 1);

        const matrixRows = [];

        driverList.forEach((driver, idx) => {
            const rowIndex = idx + 2;
            const row = [driver];

            allDateCols.forEach(dc => {
                const dateKey = dc.date.getTime();
                const rawCredit = driverCreditMap[driver] ? (driverCreditMap[driver][dateKey] || 0) : 0;
                row.push(rawCredit === 0 ? 0 : "");
            });

            row.push(`=SUM(B${rowIndex}:${lastDateColLetter}${rowIndex})`);
            matrixRows.push(row);
        });

        const bottomRowIndex = driverList.length + 2;
        const bottomRow = ["Total Hours"];
        for (let c = 0; c < numDateCols; c++) {
            const colLetter = getColumnLetter(c + 2);
            bottomRow.push(`=SUM(${colLetter}2:${colLetter}${bottomRowIndex - 1})`);
        }
        bottomRow.push(`=SUM(B${bottomRowIndex}:${lastDateColLetter}${bottomRowIndex})`);
        matrixRows.push(bottomRow);

        hoursSheet.getRange(1, 1, 1, headerRow.length).setNumberFormat("@").setValues([headerRow]);
        hoursSheet.getRange(2, 1, matrixRows.length, headerRow.length).setValues(matrixRows);

        applyWorkingHoursFormatting(hoursSheet, driverList.length, numDateCols, bottomRowIndex);
        syncBreaksTable(ss);
        return;
    }

    // --- ADDITIVE SYNC & AUTOMATIC CLEANUP OF ANY DUPLICATES ---
    // STEP 1: Delete any trailing columns beyond the first "Total" header
    const initialRawHeaders = hoursSheet.getRange(1, 1, 1, existingLastCol).getValues()[0];
    const initialDisplayHeaders = hoursSheet.getRange(1, 1, 1, existingLastCol).getDisplayValues()[0];

    let firstTotalCol = -1;
    for (let c = 1; c < existingLastCol; c++) {
        const hText = (initialDisplayHeaders[c] || '').toString().toLowerCase().trim();
        if (hText.includes("total")) {
            firstTotalCol = c + 1; // 1-based column index
            break;
        }
    }

    if (firstTotalCol > 1 && existingLastCol > firstTotalCol) {
        hoursSheet.deleteColumns(firstTotalCol + 1, existingLastCol - firstTotalCol);
        SpreadsheetApp.flush();
    }

    // STEP 2: Deduplicate columns between Col B and Total Hours
    const midLastCol = hoursSheet.getLastColumn();
    const midLastRow = hoursSheet.getLastRow();
    const midDisplayHeaders = hoursSheet.getRange(1, 1, 1, midLastCol).getDisplayValues()[0];
    const midRawHeaders = hoursSheet.getRange(1, 1, 1, midLastCol).getValues()[0];

    let totalColIndex = midLastCol;
    for (let c = 1; c < midLastCol; c++) {
        const hText = (midDisplayHeaders[c] || '').toString().toLowerCase().trim();
        if (hText.includes("total")) {
            totalColIndex = c + 1;
            break;
        }
    }

    const seenDates = {}; // normKey -> 1-based primaryColNumber
    const colsToDelete = [];

    for (let c = 1; c < totalColIndex - 1; c++) {
        const colNum = c + 1;
        const disp = (midDisplayHeaders[c] || '').toString().trim();
        const raw = midRawHeaders[c];

        if (!disp || disp.toLowerCase().includes("total")) {
            colsToDelete.push(colNum);
            continue;
        }

        const d = parseSheetDateHeader(disp, raw);
        const normKey = d ? d.getTime() : disp.replace(/[^0-9]/g, '');

        if (seenDates[normKey]) {
            const primaryCol = seenDates[normKey];
            if (midLastRow > 1) {
                const primaryRange = hoursSheet.getRange(2, primaryCol, midLastRow - 1, 1);
                const dupRange = hoursSheet.getRange(2, colNum, midLastRow - 1, 1);
                const pVals = primaryRange.getValues();
                const dVals = dupRange.getValues();
                let needUpdate = false;
                for (let r = 0; r < pVals.length; r++) {
                    const pNum = parseNumber(pVals[r][0]);
                    const dNum = parseNumber(dVals[r][0]);
                    if ((pNum === null || pNum === 0) && dNum !== null && dNum > 0) {
                        pVals[r][0] = dNum;
                        needUpdate = true;
                    }
                }
                if (needUpdate) primaryRange.setValues(pVals);
            }
            colsToDelete.push(colNum);
        } else {
            seenDates[normKey] = colNum;
        }
    }

    if (colsToDelete.length > 0) {
        colsToDelete.sort((a, b) => b - a);
        colsToDelete.forEach(colNum => {
            hoursSheet.deleteColumn(colNum);
        });
        SpreadsheetApp.flush();
    }

    // STEP 3: Identify the bottom "Total Hours" row of the main table
    let curLastRow = hoursSheet.getLastRow();
    let curLastCol = hoursSheet.getLastColumn();
    const colAData = hoursSheet.getRange(1, 1, curLastRow, 1).getDisplayValues();
    let firstBottomTotalRow = -1;

    for (let r = 1; r < colAData.length; r++) {
        const txt = (colAData[r][0] || '').toString().toLowerCase().trim();
        if (txt.startsWith("total") && !txt.includes("break")) {
            firstBottomTotalRow = r + 1; // 1-based index
            break;
        }
    }

    // STEP 4: Now scan the cleaned sheet to find existing columns and drivers
    curLastRow = hoursSheet.getLastRow();
    curLastCol = hoursSheet.getLastColumn();

    let totalColNumber = curLastCol;
    for (let c = 1; c < curLastCol; c++) {
        const hText = (hoursSheet.getRange(1, c + 1).getDisplayValue() || '').toString().toLowerCase().trim();
        if (hText.includes("total")) {
            totalColNumber = c + 1;
            break;
        }
    }

    let bottomTotalRowNumber = (firstBottomTotalRow > -1) ? firstBottomTotalRow : curLastRow + 1;

    const cleanedHeaders = hoursSheet.getRange(1, 1, 1, totalColNumber).getDisplayValues()[0];
    const cleanedRawHeaders = hoursSheet.getRange(1, 1, 1, totalColNumber).getValues()[0];

    const existingDateCols = {}; // timestamp -> 1-based colNumber
    const existingDateDigits = {}; // digits -> 1-based colNumber
    const existingDateStrings = {}; // lowercase str -> 1-based colNumber

    for (let c = 1; c < totalColNumber - 1; c++) {
        const disp = (cleanedHeaders[c] || '').toString().trim();
        const raw = cleanedRawHeaders[c];
        const d = parseSheetDateHeader(disp, raw);
        if (d && !isNaN(d.getTime())) {
            existingDateCols[d.getTime()] = c + 1;
        }
        const digits = disp.replace(/[^0-9]/g, '');
        if (digits.length >= 6) {
            existingDateDigits[digits] = c + 1;
        }
        if (disp) {
            existingDateStrings[disp.toLowerCase()] = c + 1;
        }
    }

    const numDriversExisting = Math.max(0, bottomTotalRowNumber - 2);
    const existingDriverRows = {};
    if (numDriversExisting > 0) {
        const driverColData = hoursSheet.getRange(2, 1, numDriversExisting, 1).getDisplayValues();
        for (let r = 0; r < driverColData.length; r++) {
            const rawName = (driverColData[r][0] || '').toString().trim();
            if (!rawName) continue;
            const canon = getCanonicalDriverName(rawName, canonicalDrivers);
            existingDriverRows[canon] = r + 2;
            existingDriverRows[rawName] = r + 2;
        }
    }

    // STEP 5: Insert any truly missing date columns before "Total Hours"
    const missingDates = allDateCols.filter(dc => {
        const t = dc.date.getTime();
        const digits = (dc.headerStr || '').replace(/[^0-9]/g, '');
        const str = (dc.headerStr || '').trim().toLowerCase();
        return !existingDateCols[t] && !existingDateDigits[digits] && !existingDateStrings[str];
    });

    for (let i = 0; i < missingDates.length; i++) {
        const dc = missingDates[i];
        hoursSheet.insertColumnBefore(totalColNumber);
        hoursSheet.getRange(1, totalColNumber).setNumberFormat("@").setValue(dc.headerStr);

        const numExistingDrivers = bottomTotalRowNumber - 2;
        if (numExistingDrivers > 0) {
            const curDriverData = hoursSheet.getRange(2, 1, numExistingDrivers, 1).getDisplayValues();
            const colValues = [];
            for (let r = 0; r < numExistingDrivers; r++) {
                const driverName = (curDriverData[r][0] || '').toString().trim();
                const canon = getCanonicalDriverName(driverName, canonicalDrivers);
                const rawCredit = (driverCreditMap[canon] && driverCreditMap[canon][dc.date.getTime()]) ? driverCreditMap[canon][dc.date.getTime()] : 0;
                colValues.push([rawCredit === 0 ? 0 : ""]);
            }
            hoursSheet.getRange(2, totalColNumber, numExistingDrivers, 1).setValues(colValues);
        }

        existingDateCols[dc.date.getTime()] = totalColNumber;
        const digits = (dc.headerStr || '').replace(/[^0-9]/g, '');
        if (digits.length >= 6) existingDateDigits[digits] = totalColNumber;
        existingDateStrings[(dc.headerStr || '').trim().toLowerCase()] = totalColNumber;
        totalColNumber++;
    }

    // STEP 6: Insert any missing driver rows before bottom "Total Hours"
    const missingDrivers = driverList.filter(d => !existingDriverRows[d]);
    for (let i = 0; i < missingDrivers.length; i++) {
        const missingDriver = missingDrivers[i];
        hoursSheet.insertRowBefore(bottomTotalRowNumber);

        const rowValues = [missingDriver];
        for (let c = 2; c < totalColNumber; c++) {
            let colTimestamp = null;
            for (const t in existingDateCols) {
                if (existingDateCols[t] === c) { colTimestamp = Number(t); break; }
            }
            const rawCredit = (colTimestamp && driverCreditMap[missingDriver] && driverCreditMap[missingDriver][colTimestamp]) ? driverCreditMap[missingDriver][colTimestamp] : 0;
            rowValues.push(rawCredit === 0 ? 0 : "");
        }
        hoursSheet.getRange(bottomTotalRowNumber, 1, 1, rowValues.length).setValues([rowValues]);

        existingDriverRows[missingDriver] = bottomTotalRowNumber;
        bottomTotalRowNumber++;
    }

    // STEP 7: Refresh SUM formulas
    const numDrivers = bottomTotalRowNumber - 2;
    const numDateCols = totalColNumber - 2;
    const lastDateColLetter = getColumnLetter(totalColNumber - 1);

    if (numDrivers > 0 && numDateCols > 0) {
        hoursSheet.getRange(1, totalColNumber).setValue("Total Hours");

        const totalFormulas = [];
        for (let r = 2; r < bottomTotalRowNumber; r++) {
            totalFormulas.push([`=SUM(B${r}:${lastDateColLetter}${r})`]);
        }
        hoursSheet.getRange(2, totalColNumber, numDrivers, 1).setFormulas(totalFormulas);

        const bottomRowFormulas = ["Total Hours"];
        for (let c = 2; c < totalColNumber; c++) {
            const cLetter = getColumnLetter(c);
            bottomRowFormulas.push(`=SUM(${cLetter}2:${cLetter}${bottomTotalRowNumber - 1})`);
        }
        bottomRowFormulas.push(`=SUM(B${bottomTotalRowNumber}:${lastDateColLetter}${bottomTotalRowNumber})`);
        hoursSheet.getRange(bottomTotalRowNumber, 1, 1, bottomRowFormulas.length).setValues([bottomRowFormulas]);
    }

    applyWorkingHoursFormatting(hoursSheet, numDrivers, numDateCols, bottomTotalRowNumber);
    syncBreaksTable(ss);
}

function applyWorkingHoursFormatting(hoursSheet, numDrivers, numDateCols, bottomRowIndex) {
    const totalCols = numDateCols + 2; // Col A (Driver) + Date cols + Col Total Hours
    const totalRows = numDrivers + 2;  // Header + Driver rows + Bottom Total

    // Header styling
    hoursSheet.getRange(1, 1, 1, totalCols)
        .setFontWeight("bold")
        .setBackground("#3c78d8")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    // Ensure header row date columns are plain text
    hoursSheet.getRange(1, 1, 1, totalCols - 1).setNumberFormat("@");

    // Matrix numbers format & alignment
    if (numDrivers > 0 && totalCols > 1) {
        hoursSheet.getRange(2, 2, numDrivers + 1, totalCols - 1)
            .setNumberFormat("0.0")
            .setHorizontalAlignment("center");
        hoursSheet.getRange(2, 1, numDrivers, 1)
            .setFontWeight("bold")
            .setHorizontalAlignment("left");
    }

    // Bottom total row styling
    hoursSheet.getRange(bottomRowIndex, 1, 1, totalCols)
        .setFontWeight("bold")
        .setBackground("#d0e0e3")
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
    hoursSheet.getRange(bottomRowIndex, 1).setHorizontalAlignment("left");

    // Grid borders
    hoursSheet.getRange(1, 1, totalRows, totalCols)
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    // Data validation and conditional formatting
    if (numDrivers > 0 && numDateCols > 0) {
        const matrixRange = hoursSheet.getRange(2, 2, numDrivers, numDateCols);

        const firstCell = "B2";
        const rule = SpreadsheetApp.newDataValidation()
            .requireFormulaSatisfied(`=AND(ISNUMBER(${firstCell}), ${firstCell}>=0, MOD(${firstCell}*2, 1)=0)`)
            .setAllowInvalid(true)
            .setHelpText("Hours must be entered in 0.5 increments (e.g. 0, 0.5, 1, 1.5, 2, ...)")
            .build();

        matrixRange.setDataValidation(rule);

        const redRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberEqualTo(0)
            .setBackground("#f4cccc")
            .setFontColor("#990000")
            .setRanges([matrixRange])
            .build();

        const yellowRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberBetween(0.001, 7.999)
            .setBackground("#fff2cc")
            .setFontColor("#7f6000")
            .setRanges([matrixRange])
            .build();

        const greenRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThanOrEqualTo(8)
            .setBackground("#d9ead3")
            .setFontColor("#274e13")
            .setRanges([matrixRange])
            .build();

        hoursSheet.clearConditionalFormatRules();
        hoursSheet.setConditionalFormatRules([redRule, yellowRule, greenRule]);
    }

    hoursSheet.autoResizeColumns(1, totalCols);
    for (let c = 1; c <= totalCols; c++) {
        if (hoursSheet.getColumnWidth(c) < 80) hoursSheet.setColumnWidth(c, 80);
    }
}

// =================================================================
// BREAKS SHEET LOAD & TABLE RENDERING HELPERS
// =================================================================

/**
 * Converts a time representation (Date object, string like "10:30 AM", "17:30") to minutes from midnight.
 */
function parseTimeToMinutes(val) {
    if (!val) return null;
    if (val instanceof Date) {
        return val.getHours() * 60 + val.getMinutes() + (val.getSeconds() / 60);
    }
    const str = val.toString().trim();
    const match = str.match(/(\d+):(\d+)(?::(\d+))?\s*(AM|PM)?/i);
    if (match) {
        let h = parseInt(match[1], 10);
        const m = parseInt(match[2], 10);
        const s = match[3] ? parseInt(match[3], 10) : 0;
        const ampm = match[4] ? match[4].toUpperCase() : null;
        if (ampm === "PM" && h < 12) h += 12;
        if (ampm === "AM" && h === 12) h = 0;
        return h * 60 + m + (s / 60);
    }
    return null;
}

/**
 * Parses break duration into decimal hours (e.g. 0:30:00 -> 0.5).
 * Returns 0 if break is incomplete (empty 'To'), negative, or invalid.
 */
function parseBreakPeriodToHours(periodVal, periodDisp, fromVal, toVal) {
    if (toVal === undefined || toVal === null || String(toVal).trim() === "") {
        return 0;
    }

    const str = (periodDisp || '').toString().trim();
    if (str.startsWith("-") || str === "0:00:00" || str === "0:00" || str === "0") {
        return 0;
    }

    if (str && str.includes(":")) {
        const parts = str.split(":");
        const h = parseFloat(parts[0]) || 0;
        const m = parseFloat(parts[1]) || 0;
        const s = (parts.length > 2 ? parseFloat(parts[2]) : 0) || 0;
        const totalHours = h + (m / 60) + (s / 3600);
        if (totalHours > 0) return roundToTwo(totalHours);
    }

    if (periodVal instanceof Date) {
        const h = periodVal.getHours();
        const m = periodVal.getMinutes();
        const s = periodVal.getSeconds();
        const totalHours = h + (m / 60) + (s / 3600);
        if (totalHours > 0) return roundToTwo(totalHours);
    }

    if (typeof periodVal === "number" && periodVal > 0) {
        return roundToTwo(periodVal * 24);
    }

    if (fromVal && toVal) {
        const fromMins = parseTimeToMinutes(fromVal);
        const toMins = parseTimeToMinutes(toVal);
        if (fromMins !== null && toMins !== null && toMins > fromMins) {
            return roundToTwo((toMins - fromMins) / 60);
        }
    }

    return 0;
}

/**
 * Helper to find a sheet in a spreadsheet case-insensitively, trimming spaces.
 */
function findSheetCaseInsensitive(ss, name) {
    if (!ss || !name) return null;
    const direct = ss.getSheetByName(name);
    if (direct) return direct;
    const target = name.toLowerCase().replace(/\s+/g, ' ').trim();
    const sheets = ss.getSheets();
    for (let i = 0; i < sheets.length; i++) {
        const sName = sheets[i].getName().toLowerCase().replace(/\s+/g, ' ').trim();
        if (sName === target) return sheets[i];
    }
    return null;
}

/**
 * Reads the "breaks" sheet and returns:
 * - map: { canonDriver -> { dateTimestamp -> totalBreakHours } }
 * - driverTotals: { canonDriver -> totalBreakHours }
 * - dailyTotals: { dateTimestamp -> totalBreakHours }
 */
function loadBreaksMap(ss, canonicalDrivers) {
    if (!ss) {
        try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { }
        if (!ss) ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    }
    const breaksSheet = findSheetCaseInsensitive(ss, "breaks");
    if (!breaksSheet) {
        Logger.log("ℹ️ Breaks sheet not found, proceeding with 0 break hours.");
        return { map: {}, driverTotals: {}, dailyTotals: {} };
    }

    const lastRow = breaksSheet.getLastRow();
    const lastCol = breaksSheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) return { map: {}, driverTotals: {}, dailyTotals: {} };

    const rawHeaders = breaksSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const dispHeaders = breaksSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

    let dateColIdx = -1, driverColIdx = -1, fromColIdx = -1, toColIdx = -1, periodColIdx = -1;
    for (let c = 0; c < lastCol; c++) {
        const h = (dispHeaders[c] || rawHeaders[c] || '').toString().toLowerCase().trim();
        if (h.includes("date")) dateColIdx = c;
        else if (h.includes("driver") || h.includes("name")) driverColIdx = c;
        else if (h.includes("from") || h.includes("start")) fromColIdx = c;
        else if (h.includes("to") || h.includes("end")) toColIdx = c;
        else if (h.includes("period") || h.includes("duration") || h.includes("hour") || h.includes("break")) periodColIdx = c;
    }

    if (dateColIdx === -1) dateColIdx = 0;
    if (driverColIdx === -1) driverColIdx = 1;
    if (fromColIdx === -1) fromColIdx = 2;
    if (toColIdx === -1) toColIdx = 3;
    if (periodColIdx === -1) periodColIdx = 4;

    const rawData = breaksSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    const dispData = breaksSheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

    const map = {};
    const driverTotals = {};
    const dailyTotals = {};
    const rawToCanon = {};

    for (let r = 0; r < rawData.length; r++) {
        const rawDriver = (rawData[r][driverColIdx] || dispData[r][driverColIdx] || '').toString().trim();
        if (!rawDriver) continue;

        const canonDriver = getCanonicalDriverName(rawDriver, canonicalDrivers) || rawDriver;
        rawToCanon[rawDriver] = canonDriver;

        const rawDate = rawData[r][dateColIdx];
        const dispDate = dispData[r][dateColIdx];
        const d = (rawDate instanceof Date) ? rawDate : (parseDate(rawDate) || parseDate(dispDate));
        if (!d || isNaN(d.getTime())) continue;

        const dateKey = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();

        const periodVal = rawData[r][periodColIdx];
        const periodDisp = dispData[r][periodColIdx];
        const fromVal = rawData[r][fromColIdx] || dispData[r][fromColIdx];
        const toVal = rawData[r][toColIdx] || dispData[r][toColIdx];

        const breakHours = parseBreakPeriodToHours(periodVal, periodDisp, fromVal, toVal);
        if (breakHours <= 0) continue;

        if (!map[canonDriver]) map[canonDriver] = {};
        map[canonDriver][dateKey] = roundToTwo((map[canonDriver][dateKey] || 0) + breakHours);

        driverTotals[canonDriver] = roundToTwo((driverTotals[canonDriver] || 0) + breakHours);
        dailyTotals[dateKey] = roundToTwo((dailyTotals[dateKey] || 0) + breakHours);
    }

    // Link raw driver name aliases so lookups by raw or canonical driver return identical totals
    for (const raw in rawToCanon) {
        const canon = rawToCanon[raw];
        if (raw !== canon && map[canon]) {
            map[raw] = map[canon];
            driverTotals[raw] = driverTotals[canon];
        }
    }

    return { map, driverTotals, dailyTotals };
}

/**
 * Standalone function to sync and render the Breaks Hours table directly
 * below the main table in the "Working hours" sheet.
 * Reads existing dates and drivers directly from the sheet to guarantee exact alignment.
 */
function syncBreaksTable(ss) {
    if (!ss) {
        try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { }
        if (!ss) ss = SpreadsheetApp.openById(TARGET_SHEET_ID);
    }
    if (!ss) throw new Error("Could not access spreadsheet.");

    let hoursSheet = null;
    try {
        const activeS = ss.getActiveSheet();
        if (activeS && activeS.getName().toLowerCase().replace(/\s+/g, ' ').trim() === "working hours") {
            hoursSheet = activeS;
        }
    } catch (e) { }
    if (!hoursSheet) {
        hoursSheet = findSheetCaseInsensitive(ss, "Working hours");
    }
    if (!hoursSheet) {
        throw new Error("Sheet 'Working hours' was not found in spreadsheet '" + ss.getName() + "'. Available sheets: " + ss.getSheets().map(s => '"' + s.getName() + '"').join(", "));
    }

    const lastRow = hoursSheet.getLastRow();
    const lastCol = hoursSheet.getLastColumn();
    if (lastRow < 2 || lastCol < 2) {
        throw new Error("Working hours sheet has insufficient data (lastRow: " + lastRow + ", lastCol: " + lastCol + ").");
    }

    // 1. Find the bottom "Total Hours" row of the main table
    const colAData = hoursSheet.getRange(1, 1, lastRow, 1).getDisplayValues();
    let mainBottomRow = -1;
    for (let r = 1; r < colAData.length; r++) {
        const txt = (colAData[r][0] || '').toString().toLowerCase().trim();
        if (txt.startsWith("total")) {
            mainBottomRow = r + 1; // 1-based index
            break;
        }
    }
    if (mainBottomRow === -1) mainBottomRow = lastRow;

    // 2. Find the "Total Hours" column in Row 1
    const row1Data = hoursSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
    let totalColNumber = -1;
    for (let c = 1; c < lastCol; c++) {
        const hText = (row1Data[c] || '').toString().toLowerCase().trim();
        if (hText.includes("total")) {
            totalColNumber = c + 1;
            break;
        }
    }
    if (totalColNumber === -1) totalColNumber = lastCol;

    const numDateCols = totalColNumber - 2;
    if (numDateCols <= 0) {
        throw new Error("Could not find date columns in Row 1 of 'Working hours'. totalColNumber: " + totalColNumber + ", headers: " + row1Data.join(" | "));
    }

    // 3. Read date columns directly from Row 1
    const row1Raw = hoursSheet.getRange(1, 1, 1, totalColNumber).getValues()[0];
    const allDateCols = [];
    for (let c = 2; c < totalColNumber; c++) {
        const disp = (row1Data[c - 1] || '').toString().trim();
        const raw = row1Raw[c - 1];
        const d = parseSheetDateHeader(disp, raw) || parseDate(disp);
        allDateCols.push({
            headerStr: disp,
            dateKey: d ? new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime() : null
        });
    }

    // 4. Read drivers from main table (rows 2 to mainBottomRow - 1)
    const numDrivers = mainBottomRow - 2;
    if (numDrivers <= 0) {
        throw new Error("No driver rows found between Row 1 and Total Hours row (" + mainBottomRow + ").");
    }
    const driverColData = hoursSheet.getRange(2, 1, numDrivers, 1).getDisplayValues();
    const finalDrivers = [];
    for (let r = 0; r < driverColData.length; r++) {
        const dName = (driverColData[r][0] || '').toString().trim();
        if (dName && !dName.toLowerCase().startsWith("total")) {
            finalDrivers.push(dName);
        }
    }
    if (finalDrivers.length === 0) {
        throw new Error("Found 0 driver names in Column A between Row 2 and Row " + (mainBottomRow - 1) + ".");
    }

    // 5. Seed canonical drivers from Settings
    const settingsSheet = findSheetCaseInsensitive(ss, "Settings");
    const canonicalDrivers = [];
    if (settingsSheet) {
        const sData = settingsSheet.getDataRange().getValues();
        for (let i = 1; i < sData.length; i++) {
            const d = (sData[i][0] || "").toString().trim();
            if (d && !canonicalDrivers.includes(d)) canonicalDrivers.push(d);
        }
    }

    // 6. Load breaks map from 'breaks' sheet
    let breaksMap = {};
    try {
        const bRes = loadBreaksMap(ss, canonicalDrivers);
        breaksMap = (bRes && bRes.map) ? bRes.map : {};
    } catch (e) {
        Logger.log("Warning: loadBreaksMap encountered an error: " + e.message);
    }

    const totalCols = numDateCols + 2;
    const lastDateColLetter = getColumnLetter(numDateCols + 1);

    // Position: 2 empty rows after main table bottom row
    const titleRow = mainBottomRow + 3;
    const headerRow = titleRow + 1;
    const firstDataRow = headerRow + 1;
    const bottomRow = firstDataRow + finalDrivers.length;

    // Ensure sheet has enough rows
    const currentMaxRows = hoursSheet.getMaxRows();
    if (currentMaxRows < bottomRow + 5) {
        hoursSheet.insertRowsAfter(currentMaxRows, (bottomRow + 5) - currentMaxRows);
    }

    const curLastCol = Math.max(hoursSheet.getLastColumn(), totalCols);

    // Unmerge any existing merges in the target table area without wiping contents
    try {
        hoursSheet.getRange(titleRow, 1, (bottomRow - titleRow + 1), curLastCol).breakApart();
    } catch (e) { }

    const stepErrors = [];

    // A. Title Block (Row 15)
    try {
        hoursSheet.getRange(titleRow, 1).setValue("☕ BREAKS HOURS (Reference Only)");
        const titleRange = hoursSheet.getRange(titleRow, 1, 1, totalCols);
        titleRange
            .setFontWeight("bold")
            .setFontSize(11)
            .setBackground("#b45f06")
            .setFontColor("white")
            .setVerticalAlignment("middle");

        const frozenCols = hoursSheet.getFrozenColumns();
        if (frozenCols > 0 && frozenCols < totalCols) {
            hoursSheet.getRange(titleRow, 1).setHorizontalAlignment("left");
            try {
                hoursSheet.getRange(titleRow, frozenCols + 1, 1, totalCols - frozenCols).merge();
            } catch (mErr) { }
        } else {
            hoursSheet.getRange(titleRow, 1).setHorizontalAlignment("center");
            try {
                titleRange.merge();
            } catch (mErr) { }
        }
        SpreadsheetApp.flush();
    } catch (eA) {
        stepErrors.push("Step A (Title): " + eA.message);
    }

    // B. Header Row (Row 16)
    try {
        const headers = ["Driver"];
        allDateCols.forEach(dc => headers.push(dc.headerStr));
        headers.push("Total Breaks");

        const headerRange = hoursSheet.getRange(headerRow, 1, 1, headers.length);
        headerRange.setValues([headers]);
        headerRange.setFontWeight("bold")
            .setBackground("#e69138")
            .setFontColor("white")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle");
        SpreadsheetApp.flush();
    } catch (eB) {
        stepErrors.push("Step B (Header): " + eB.message);
    }

    // C. Driver Data Rows (Rows 17 to 26)
    try {
        const matrixRows = [];
        finalDrivers.forEach((driver, idx) => {
            const rIdx = firstDataRow + idx;
            const row = [driver];
            const canon = getCanonicalDriverName(driver, canonicalDrivers);

            allDateCols.forEach(dc => {
                let bVal = 0;
                if (dc.dateKey) {
                    if (breaksMap[canon] && breaksMap[canon][dc.dateKey] !== undefined) {
                        bVal = breaksMap[canon][dc.dateKey];
                    } else if (breaksMap[driver] && breaksMap[driver][dc.dateKey] !== undefined) {
                        bVal = breaksMap[driver][dc.dateKey];
                    }
                }
                row.push(bVal > 0 ? roundToTwo(bVal) : 0);
            });

            row.push(`=SUM(B${rIdx}:${lastDateColLetter}${rIdx})`);
            matrixRows.push(row);
        });

        const dataRange = hoursSheet.getRange(firstDataRow, 1, matrixRows.length, totalCols);
        dataRange.setValues(matrixRows);
        hoursSheet.getRange(firstDataRow, 1, matrixRows.length, 1)
            .setFontWeight("bold")
            .setHorizontalAlignment("left");
        hoursSheet.getRange(firstDataRow, 2, matrixRows.length, numDateCols + 1)
            .setNumberFormat("0.0")
            .setHorizontalAlignment("center");
        SpreadsheetApp.flush();
    } catch (eC) {
        stepErrors.push("Step C (Drivers): " + eC.message);
    }

    // D. Bottom Total Row (Row 27)
    try {
        const bottomRowFormulas = ["Total Breaks"];
        for (let c = 2; c <= numDateCols + 1; c++) {
            const cLetter = getColumnLetter(c);
            bottomRowFormulas.push(`=SUM(${cLetter}${firstDataRow}:${cLetter}${bottomRow - 1})`);
        }
        bottomRowFormulas.push(`=SUM(B${bottomRow}:${lastDateColLetter}${bottomRow})`);

        const bottomRange = hoursSheet.getRange(bottomRow, 1, 1, totalCols);
        bottomRange.setValues([bottomRowFormulas]);
        bottomRange.setFontWeight("bold")
            .setBackground("#fce5cd")
            .setHorizontalAlignment("center")
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        hoursSheet.getRange(bottomRow, 1).setHorizontalAlignment("left");
        hoursSheet.getRange(bottomRow, 2, 1, numDateCols + 1).setNumberFormat("0.0");
        SpreadsheetApp.flush();
    } catch (eD) {
        stepErrors.push("Step D (Totals): " + eD.message);
    }

    // E. Grid borders around table
    try {
        hoursSheet.getRange(headerRow, 1, finalDrivers.length + 2, totalCols)
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    } catch (eE) {
        stepErrors.push("Step E (Borders): " + eE.message);
    }

    // F. Highlight break days > 0
    try {
        const breakDataRange = hoursSheet.getRange(firstDataRow, 2, finalDrivers.length, numDateCols);
        const breakHighlightRule = SpreadsheetApp.newConditionalFormatRule()
            .whenNumberGreaterThan(0)
            .setBackground("#fff2cc")
            .setFontColor("#b45f06")
            .setRanges([breakDataRange])
            .build();

        const currentRules = hoursSheet.getConditionalFormatRules() || [];
        currentRules.push(breakHighlightRule);
        hoursSheet.setConditionalFormatRules(currentRules);
    } catch (eF) { }

    // G. Clear ONLY leftover rows strictly below bottomRow (if any)
    try {
        const maxR = hoursSheet.getMaxRows();
        if (maxR > bottomRow) {
            const leftoverRows = maxR - bottomRow;
            const leftoverRange = hoursSheet.getRange(bottomRow + 1, 1, leftoverRows, curLastCol);
            leftoverRange.breakApart();
            leftoverRange.clearContent().clearFormat();
        }
    } catch (eG) { }

    // Ensure all rows in Working hours are unhidden and visible
    try {
        hoursSheet.showRows(1, hoursSheet.getMaxRows());
    } catch (eH) { }

    // (Do not activate or change the user's active sheet tab)

    SpreadsheetApp.flush();

    const writtenVal = hoursSheet.getRange(titleRow, 1).getValue() || hoursSheet.getRange(titleRow, 1).getDisplayValue();
    const writtenColor = hoursSheet.getRange(titleRow, 1).getBackground();
    Logger.log("✅ syncBreaksTable completed at row " + titleRow + ": '" + writtenVal + "' (" + writtenColor + ")");

    return {
        sheetName: hoursSheet.getName(),
        spreadsheetName: ss.getName(),
        titleRow: titleRow,
        headerRow: headerRow,
        firstDataRow: firstDataRow,
        bottomRow: bottomRow,
        numDrivers: finalDrivers.length,
        totalCols: totalCols,
        writtenTitle: writtenVal,
        writtenBg: writtenColor,
        errors: stepErrors
    };
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
    const ss = SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.openById(TARGET_SHEET_ID);

    // 0. Sync working hours sheet first so downstream reports can use working hours data
    try {
        generateWorkingHoursSheet(ss);
    } catch (e) {
        Logger.log("Error in generateWorkingHoursSheet: " + e.toString() + "\n" + (e.stack || ""));
    }

    // 0b. Sync breaks table directly
    try {
        syncBreaksTable(ss);
    } catch (e) {
        Logger.log("Error in syncBreaksTable: " + e.toString() + "\n" + (e.stack || ""));
    }

    // 1. Update the main Summary sheet from Raw Data.
    updateSummaryAndCharts(ss);

    // 2. Wait for all pending spreadsheet changes to apply.
    SpreadsheetApp.flush();

    // 3. Generate the weekly summary.
    try {
        generateWeeklySummary(ss);
    } catch (e) {
        Logger.log("Error in generateWeeklySummary: " + e.toString() + "\n" + (e.stack || ""));
    }

    // 4. Generate the final Bonus report.
    try {
        generateBonusReport(ss);
    } catch (e) {
        Logger.log("Error in generateBonusReport: " + e.toString() + "\n" + (e.stack || ""));
    }

    // 5. Generate Angel's specific progressive report.
    try {
        generateAngelReport(ss);
    } catch (e) {
        Logger.log("Error in generateAngelReport: " + e.toString() + "\n" + (e.stack || ""));
    }
}


// =================================================================
// SUMMARY BUILD — Reads "Raw Data" tab and builds the Summary sheet
// =================================================================
function updateSummaryAndCharts(ss) {
    let summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) summarySheet = ss.insertSheet("Summary");

    const rawSheet = ss.getSheetByName("Raw Data");
    if (!rawSheet) throw new Error('No sheet named "Raw Data" was found.');

    // Seed canonical drivers from Settings first so auxiliary loaders can normalize names
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

    // --- Helper to load exact daily data from auxiliary sheets ---
    function loadAuxMap(sheetName) {
        const auxSheet = ss.getSheetByName(sheetName);
        if (!auxSheet) return {};
        const weeks = parseRawDataWeeks(auxSheet);
        if (weeks.length > 0) {
            const map = {};
            weeks.forEach(w => {
                w.drivers.forEach(d => {
                    const rawName = d.name;
                    const canonName = getCanonicalDriverName(rawName, canonicalDrivers);
                    if (!map[rawName]) map[rawName] = {};
                    if (!map[canonName]) map[canonName] = {};
                    for (const t in d.dailyCredits) {
                        map[rawName][t] = d.dailyCredits[t];
                        map[canonName][t] = d.dailyCredits[t];
                    }
                });
            });
            return map;
        }

        const data = auxSheet.getDataRange().getValues();
        if (data.length < 2) return {};
        const lastCol = auxSheet.getLastColumn();
        const rawHeaders = auxSheet.getRange(1, 1, 1, lastCol).getValues()[0];
        const displayHeaders = auxSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

        const dateCols = {};
        for (let c = 1; c < lastCol; c++) {
            const d = parseSheetDateHeader(displayHeaders[c], rawHeaders[c]);
            if (d && !isNaN(d.getTime())) dateCols[c] = d.getTime();
        }

        const map = {};
        for (let r = 1; r < data.length; r++) {
            const rawDriver = (data[r][0] || "").toString().trim();
            if (rawDriver && !rawDriver.toLowerCase().startsWith("total") && rawDriver.toLowerCase() !== "driver") {
                const canonDriver = getCanonicalDriverName(rawDriver, canonicalDrivers);
                if (!map[rawDriver]) map[rawDriver] = {};
                if (!map[canonDriver]) map[canonDriver] = {};
                for (const c in dateCols) {
                    const val = parseNumber(data[r][c]);
                    if (val !== null && !isNaN(val)) {
                        map[rawDriver][dateCols[c]] = val;
                        map[canonDriver][dateCols[c]] = val;
                    }
                }
            }
        }
        return map;
    }

    // --- Helper to load persistent memory of previously-calculated daily cash and trips ---
    function loadPersistentMemory() {
        const memory = {}; // { [canonDriver]: { [dateKey]: { cash, trips } } }

        // Seed known historical values for Sep 1, 2026 so they are never lost if Summary previously had zeros
        const sep1Key = new Date(2026, 8, 1).getTime();
        const knownSep1 = {
            "Hassan En Nejjari": { cash: 33.00, trips: 14 },
            "Angel Yoy": { cash: 48.00, trips: 17 },
            "Koba Svanadze": { cash: 30.00, trips: 11 },
            "Biaoming Feng": { cash: 25.50, trips: 9 },
            "Brian Macancela": { cash: 42.00, trips: 16 },
            "Benjamin Douglass": { cash: 12.00, trips: 4 },
            "Prince Verma": { cash: 30.00, trips: 11 },
            "Romer Elias": { cash: 9.00, trips: 5 }
        };
        for (const d in knownSep1) {
            const canon = getCanonicalDriverName(d, canonicalDrivers);
            if (!memory[canon]) memory[canon] = {};
            memory[canon][sep1Key] = { cash: knownSep1[d].cash, trips: knownSep1[d].trips };
        }

        // 1. Read from PropertiesService
        try {
            const props = PropertiesService.getScriptProperties();
            const stored = props.getProperty("DAILY_RECORDED_MEMORY");
            if (stored) {
                const parsed = JSON.parse(stored);
                for (const dr in parsed) {
                    const canon = getCanonicalDriverName(dr, canonicalDrivers);
                    if (!memory[canon]) memory[canon] = {};
                    for (const dk in parsed[dr]) {
                        if (!memory[canon][dk]) memory[canon][dk] = {};
                        if (parsed[dr][dk].cash !== undefined && parsed[dr][dk].cash > 0) {
                            memory[canon][dk].cash = parsed[dr][dk].cash;
                        }
                        if (parsed[dr][dk].trips !== undefined && parsed[dr][dk].trips > 0) {
                            memory[canon][dk].trips = parsed[dr][dk].trips;
                        }
                    }
                }
            }
        } catch (e) {
            Logger.log("⚠️ Could not load ScriptProperties: " + e.message);
        }

        // 2. Read from existing Summary sheet rows (before clearing)
        const summarySheet = ss.getSheetByName("Summary");
        if (summarySheet && summarySheet.getLastRow() >= 3) {
            const lastRow = summarySheet.getLastRow();
            const lastCol = Math.min(summarySheet.getLastColumn(), 10);
            const headers = summarySheet.getRange(2, 1, 1, lastCol).getDisplayValues()[0];
            const tripsIdx = findMainSummaryCol(headers, ['trips', 'trip'], 6);
            const cashIdx = findMainSummaryCol(headers, ['cash'], 7);
            const numCols = Math.max(tripsIdx, cashIdx) + 1;
            const rows = summarySheet.getRange(3, 1, lastRow - 2, numCols).getValues();

            let prevDate = null;
            rows.forEach(r => {
                const parsedD = parseDate(r[0]);
                if (parsedD) prevDate = parsedD;
                else if (r[0] instanceof Date) prevDate = r[0];
                const dateVal = prevDate;
                const rawDriver = (r[1] || '').toString().trim();
                if (!dateVal || !rawDriver) return;

                const dk = new Date(dateVal.getFullYear(), dateVal.getMonth(), dateVal.getDate()).getTime();
                const canon = getCanonicalDriverName(rawDriver, canonicalDrivers);
                if (!memory[canon]) memory[canon] = {};
                if (!memory[canon][dk]) memory[canon][dk] = {};

                const cVal = parseNumber(r[cashIdx]);
                const tVal = parseNumber(r[tripsIdx]);

                if (cVal !== null && cVal > 0) memory[canon][dk].cash = cVal;
                if (tVal !== null && tVal > 0) memory[canon][dk].trips = tVal;
            });
        }

        return memory;
    }

    function savePersistentMemory(memory) {
        try {
            const props = PropertiesService.getScriptProperties();
            props.setProperty("DAILY_RECORDED_MEMORY", JSON.stringify(memory));
        } catch (e) {
            Logger.log("⚠️ Could not save ScriptProperties: " + e.message);
        }
    }

    const memory = loadPersistentMemory();

    // Load auxiliary daily maps (driverName -> { timestamp -> value })
    const tripsDataMap = loadAuxMap("Raw Data - Trips");
    const noShowMap = loadNoShowMap(ss, "No Show");
    const { map: workingHoursMap, driverTotalHours, dailyTotalHours } = loadWorkingHoursMap(ss);

    // --- Read the entire Raw Data sheet as weekly tables ---
    const weeks = parseRawDataWeeks(rawSheet);
    if (weeks.length === 0) {
        Logger.log("⚠️ No valid week tables found in Raw Data.");
        return;
    }

    // --- Build output rows: one row per driver per day across all weeks ---
    let outputRows = []; // [date, driver, credit, cash, trips, total, cashTotal, noShow, hours]

    weeks.forEach(w => {
        w.drivers.forEach(d => {
            const driverName = getCanonicalDriverName(d.name, canonicalDrivers);
            if (!canonicalDrivers.includes(driverName)) canonicalDrivers.push(driverName);

            const weeklyTotal = d.weeklyTotal;
            const weeklyTrips = d.weeklyTrips;
            const weeklyCash = d.weeklyCash; // Cash is Total - Driver NET

            // Collect all active days for this driver in this week
            let activeDays = [];
            w.dateColumns.forEach(dc => {
                const dateKey = dc.date.getTime();
                let credit = d.dailyCredits[dateKey] || 0;

                const hasExactTrips = tripsDataMap[driverName] && tripsDataMap[driverName][dateKey] !== undefined;
                const exactTrips = hasExactTrips ? tripsDataMap[driverName][dateKey] : 0;
                const hasNoShow = noShowMap[driverName] && noShowMap[driverName][dateKey] !== undefined;

                if (hasNoShow) {
                    credit = roundToTwo(credit - noShowMap[driverName][dateKey].cash);
                }

                // Include day if there's credit, trips, or No Show
                if (credit > 0 || exactTrips > 0 || hasNoShow) {
                    activeDays.push({ date: dc.date, dateKey: dateKey, credit: credit });
                }
            });

            // If no active day found by credit/trips but the driver has weeklyTotal or weeklyCash, fallback to first date column
            if (activeDays.length === 0 && (weeklyTotal > 0 || weeklyCash > 0) && w.dateColumns.length > 0) {
                const dc = w.dateColumns[0];
                activeDays.push({ date: dc.date, dateKey: dc.date.getTime(), credit: d.dailyCredits[dc.date.getTime()] || weeklyTotal });
            }

            if (activeDays.length === 0) return;

            // Sort active days chronologically: Day 1, Day 2, Day 3...
            activeDays.sort((a, b) => a.date - b.date);

            const dailyCashMap = {};
            const dailyTripMap = {};

            if (activeDays.length === 1) {
                // Exactly 1 active day for this driver in this table (e.g. Day 1 of the week)
                const onlyDay = activeDays[0];
                dailyCashMap[onlyDay.dateKey] = weeklyCash;

                let trips = weeklyTrips;
                if (tripsDataMap[driverName] && tripsDataMap[driverName][onlyDay.dateKey] !== undefined) {
                    trips = tripsDataMap[driverName][onlyDay.dateKey];
                }
                dailyTripMap[onlyDay.dateKey] = trips;

                // Memorize this day
                if (!memory[driverName]) memory[driverName] = {};
                memory[driverName][onlyDay.dateKey] = { cash: weeklyCash, trips: trips };
            } else {
                // Multiple active days in this week (e.g. Day 1, Day 2, etc.)
                let sumPriorCash = 0;
                let sumPriorTrips = 0;

                // For all prior days (all days before the newest/last day): use memorized numbers
                for (let i = 0; i < activeDays.length - 1; i++) {
                    const priorDay = activeDays[i];
                    let pCash = 0;
                    let pTrips = 0;

                    if (memory[driverName] && memory[driverName][priorDay.dateKey]) {
                        pCash = memory[driverName][priorDay.dateKey].cash || 0;
                        pTrips = memory[driverName][priorDay.dateKey].trips || 0;
                    }

                    if (tripsDataMap[driverName] && tripsDataMap[driverName][priorDay.dateKey] !== undefined) {
                        pTrips = tripsDataMap[driverName][priorDay.dateKey];
                    }

                    dailyCashMap[priorDay.dateKey] = pCash;
                    dailyTripMap[priorDay.dateKey] = pTrips;

                    sumPriorCash = roundToTwo(sumPriorCash + pCash);
                    sumPriorTrips += pTrips;
                }

                // The last (newest) day gets: cumulative weekly total − sum of all known prior days
                const lastDay = activeDays[activeDays.length - 1];
                const lastDayCash = roundToTwo(weeklyCash - sumPriorCash);
                let lastDayTrips = Math.max(0, weeklyTrips - sumPriorTrips);

                if (tripsDataMap[driverName] && tripsDataMap[driverName][lastDay.dateKey] !== undefined) {
                    lastDayTrips = tripsDataMap[driverName][lastDay.dateKey];
                }

                dailyCashMap[lastDay.dateKey] = lastDayCash;
                dailyTripMap[lastDay.dateKey] = lastDayTrips;

                // Memorize the newest day
                if (!memory[driverName]) memory[driverName] = {};
                memory[driverName][lastDay.dateKey] = { cash: lastDayCash, trips: lastDayTrips };
            }

            activeDays.forEach(day => {
                const dailyCash = dailyCashMap[day.dateKey] !== undefined ? dailyCashMap[day.dateKey] : 0;
                const dailyTrips = dailyTripMap[day.dateKey] !== undefined ? dailyTripMap[day.dateKey] : 0;

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

    savePersistentMemory(memory);

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
    const headers = ["Date", "Driver", "Credit", "Hours", "Driver Avg/ Hour", "CPH", "Trips", "Cash", "No Show", "Balance"];
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
        const dailyCPH = dailyHours > 0 ? roundToTwo(dailyCredit / dailyHours) : 0;

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

        const dailyDEPH = dailyHours > 0 ? roundToTwo((dailyCash + dailyBalance) / dailyHours) : 0;

        // Summary rows: Date, Driver, Credit, Hours, Driver Avg/ Hour, CPH, Trips, Cash, No Show, Balance
        rowsToWrite.push([
            displayDate,
            row[1],
            dailyCredit,
            roundToTwo(dailyHours) || 0,
            dailyDEPH,
            dailyCPH,
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
    // --- Write aggregated stats to Summary (column L = one gap after 10-col main table) ---
    // =================================================================
    const startRow = 3, startCol = 12; // Column L (A-J = 10-col main table, K = 1-col gap)
    // Clear old aggregate area
    try { summarySheet.getRange(startRow, 11, 60, 12).clearContent(); } catch (e) { }

    const aggRange = summarySheet.getRange(startRow, startCol, aggRows.length, aggRows[0].length);
    aggRange.setValues(aggRows);

    // Decrease cell size / font size of aggregated table so chart can fully cover it
    aggRange.setFontSize(7);
    summarySheet.getRange(startRow, startCol, 1, aggRows[0].length).setFontWeight("bold");

    // Compact column widths for columns L to T (cols 12 to 20)
    summarySheet.setColumnWidth(startCol, 75); // Col L (Driver name)
    for (let c = startCol + 1; c < startCol + aggRows[0].length; c++) {
        summarySheet.setColumnWidth(c, 55); // Cols M to T (Stats)
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
    const dailyStartCol = 23; // Column W (Daily total summary table starts at Column W)

    // Clear old daily table and chart area starting from Column T (col 20)
    summarySheet.getRange(2, 20, 100, 25).clearContent();
    summarySheet.getRange(2, dailyStartCol, 1, dailyHeader.length).setValues([dailyHeader]).setFontWeight("bold");

    if (dailyRows.length > 0) {
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, dailyRows[0].length).setValues(dailyRows);
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, 1).setNumberFormat("dddd, dd");
    }

    // --- Charts ---
    const allCharts = summarySheet.getCharts();
    allCharts.forEach(c => summarySheet.removeChart(c));

    // Chart 1: Daily Total Credit Trend (line) — positioned NEXT TO the daily credit table (Col AD = 30)
    if (dailyRows.length > 0) {
        const dailyChartCol = dailyStartCol + 7; // Col 30 (AD)
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


/**
 * Detects the active month and year for reporting.
 * 1. Analyzes valid dates in `allRows` or `summarySheet` to pick the most frequent month/year.
 * 2. If no dates available, searches spreadsheet name for any English month name and 4-digit year.
 * 3. Falls back to current system date.
 */
function getSpreadsheetMonthAndYear(ss, allRows) {
    const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];

    // Priority 1: Extract from valid dates in allRows
    if (Array.isArray(allRows) && allRows.length > 0) {
        const monthCounts = {};
        const yearCounts = {};
        for (let i = 0; i < allRows.length; i++) {
            const rawDate = allRows[i][0];
            const d = parseDate(rawDate);
            if (d && !isNaN(d.getTime())) {
                const m = d.getMonth();
                const y = d.getFullYear();
                monthCounts[m] = (monthCounts[m] || 0) + 1;
                yearCounts[y] = (yearCounts[y] || 0) + 1;
            }
        }
        let bestMonth = -1, maxMonthCount = 0;
        for (const m in monthCounts) {
            if (monthCounts[m] > maxMonthCount) {
                maxMonthCount = monthCounts[m];
                bestMonth = parseInt(m, 10);
            }
        }
        let bestYear = -1, maxYearCount = 0;
        for (const y in yearCounts) {
            if (yearCounts[y] > maxYearCount) {
                maxYearCount = yearCounts[y];
                bestYear = parseInt(y, 10);
            }
        }
        if (bestMonth >= 0 && bestYear > 2000) {
            return {
                monthIndex: bestMonth,
                year: bestYear,
                monthName: monthNames[bestMonth]
            };
        }
    }

    // Priority 2: Check spreadsheet name for any month name (full or short) and 4-digit year
    const ssName = (ss && typeof ss.getName === "function") ? ss.getName() : "";
    let foundMonthIndex = -1;
    for (let i = 0; i < monthNames.length; i++) {
        const regex = new RegExp("\\b" + monthNames[i] + "\\b", "i");
        if (regex.test(ssName)) {
            foundMonthIndex = i;
            break;
        }
    }
    if (foundMonthIndex === -1) {
        const shortNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
        for (let i = 0; i < shortNames.length; i++) {
            const regex = new RegExp("\\b" + shortNames[i] + "\\b", "i");
            if (regex.test(ssName)) {
                foundMonthIndex = i;
                break;
            }
        }
    }

    let foundYear = -1;
    const yearMatch = ssName.match(/\b(20\d\d)\b/);
    if (yearMatch) {
        foundYear = parseInt(yearMatch[1], 10);
    }

    const now = new Date();
    const finalMonthIndex = foundMonthIndex >= 0 ? foundMonthIndex : now.getMonth();
    const finalYear = foundYear > 2000 ? foundYear : now.getFullYear();

    return {
        monthIndex: finalMonthIndex,
        year: finalYear,
        monthName: monthNames[finalMonthIndex]
    };
}


// =================================================================
// WEEKLY SUMMARY — Generates weekly performance breakdown
// =================================================================
function generateWeeklySummary(ss) {
    const summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) {
        Logger.log("generateWeeklySummary: Summary sheet not found.");
        return;
    }
    let weeklySheet = ss.getSheetByName("Weekly Summary");
    if (!weeklySheet) weeklySheet = ss.insertSheet("Weekly Summary");

    // Ensure weeklySheet has at least 26 columns for charts and tables
    if (weeklySheet.getMaxColumns() < 26) {
        weeklySheet.insertColumnsAfter(weeklySheet.getMaxColumns(), 26 - weeklySheet.getMaxColumns());
    }

    weeklySheet.clear();
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
        const summaryHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 10)).getDisplayValues()[0];
        const curCreditCol = findMainSummaryCol(summaryHeaders, ['credit'], 2);
        const curTripsCol = findMainSummaryCol(summaryHeaders, ['trips', 'trip'], 6);
        const curCashCol = findMainSummaryCol(summaryHeaders, ['cash'], 7);
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

    // --- Month and Year Detection (from Summary dates, spreadsheet name, or current date) ---
    const { monthIndex: currentMonthIndex, year: currentYear, monthName: currentMonthName } = getSpreadsheetMonthAndYear(ss, allRows);

    // --- Load Working Hours for Current Month ---
    const { map: workingHoursMap } = loadWorkingHoursMap(ss);

    // --- Look Back Logic ---
    const firstOfMonth = new Date(currentYear, currentMonthIndex, 1);
    const firstWeekMonday = getMonday(firstOfMonth);

    if (firstWeekMonday.getMonth() < currentMonthIndex || firstWeekMonday.getFullYear() < firstOfMonth.getFullYear()) {
        try {
            const prevMonthDate = new Date(firstOfMonth);
            prevMonthDate.setMonth(currentMonthIndex - 1);
            const prevMonthName = prevMonthDate.toLocaleString('en-US', { month: 'long' });

            let files = DriveApp.getFilesByName(`${prevMonthName} - Drivers Daily Balance`);
            if (!files.hasNext()) {
                files = DriveApp.getFilesByName(`Drivers Daily Balance - ${prevMonthName}`);
            }
            if (!files.hasNext()) {
                files = DriveApp.getFilesByName(`${prevMonthName} ${prevMonthDate.getFullYear()} - Drivers Daily Balance`);
            }

            if (files.hasNext()) {
                const prevFile = files.next();
                const prevSpreadsheet = SpreadsheetApp.openById(prevFile.getId());
                const prevSummarySheet = prevSpreadsheet.getSheetByName("Summary");
                if (prevSummarySheet && prevSummarySheet.getLastRow() >= 3) {
                    const prevHeaders = prevSummarySheet.getRange(2, 1, 1, Math.min(prevSummarySheet.getLastColumn(), 10)).getDisplayValues()[0];
                    const prevCreditCol = findMainSummaryCol(prevHeaders, ['credit'], 2);
                    const prevTripsCol = findMainSummaryCol(prevHeaders, ['trips', 'trip'], 6);
                    const prevCashCol = findMainSummaryCol(prevHeaders, ['cash'], 7);
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
                    try {
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
                    } catch (eHours) {
                        Logger.log("Error merging previous month working hours: " + eHours);
                    }
                }
            }
        } catch (e) {
            Logger.log("Error in previous month lookback: " + e);
        }
    }

    // --- Process Data ---
    const weeklyData = {};
    const dateIdx = 0, driverIdx = 1, creditIdx = 2, tripsIdx = 3, cashIdx = 4;

    allRows.forEach((row) => {
        if (!Array.isArray(row) || !row[dateIdx] || !row[driverIdx]) return;
        const date = parseDate(row[dateIdx]);
        if (!date) return;

        const driver = row[driverIdx];
        const credit = parseNumber(row[creditIdx]) || 0;
        const trips = parseNumber(row[tripsIdx]) || 0;
        const cash = parseNumber(row[cashIdx]) || 0;

        const weekStart = getMonday(date);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekKey = `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;

        const dateKey = new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
        const dayHours = (workingHoursMap[driver] && workingHoursMap[driver][dateKey]) ? workingHoursMap[driver][dateKey] : 0;

        if (!weeklyData[weekKey]) {
            weeklyData[weekKey] = {
                startDate: weekStart,
                endDate: weekEnd,
                drivers: {}
            };
        }
        if (!weeklyData[weekKey].drivers[driver]) {
            weeklyData[weekKey].drivers[driver] = { credit: 0, trips: 0, cash: 0, hours: 0 };
        }

        weeklyData[weekKey].drivers[driver].credit += credit;
        weeklyData[weekKey].drivers[driver].trips += trips;
        weeklyData[weekKey].drivers[driver].cash += cash;
        weeklyData[weekKey].drivers[driver].hours += dayHours;
    });

    // --- Write Output ---
    let currentRow = 3;
    let displayedWeekCounter = 0;

    const sortedWeekKeys = Object.keys(weeklyData).sort((a, b) => {
        return weeklyData[a].startDate.getTime() - weeklyData[b].startDate.getTime();
    });

    sortedWeekKeys.forEach((weekKey) => {
        const weekBlock = weeklyData[weekKey];
        const weekStartDate = weekBlock.startDate;
        const weekEndDate = weekBlock.endDate;

        const isInCurrentMonth = (weekStartDate && weekStartDate.getMonth() === currentMonthIndex && weekStartDate.getFullYear() === currentYear) ||
            (weekEndDate && weekEndDate.getMonth() === currentMonthIndex && weekEndDate.getFullYear() === currentYear);

        if (isInCurrentMonth) {
            displayedWeekCounter++;
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

            const drivers = Object.keys(weekBlock.drivers);
            const tableData = [];

            drivers.forEach(d => {
                const info = weekBlock.drivers[d];
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

            if (tableData.length > 0) {
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

                const totalRow = [
                    "Total",
                    roundToTwo(totalCredit),
                    totalTrips,
                    roundToTwo(totalCash),
                    roundToTwo(totalHours),
                    totalAvgPerHour,
                    roundToTwo(totalBalance)
                ];

                const totalRange = weeklySheet.getRange(currentRow + tableData.length, 1, 1, 7);
                totalRange.setValues([totalRow])
                    .setFontWeight("bold")
                    .setBackground("#eeeeee")
                    .setBorder(true, true, true, true, true, true);

                weeklySheet.getRange(currentRow + tableData.length, 2, 1, 1).setNumberFormat("$#,##0.00");
                weeklySheet.getRange(currentRow + tableData.length, 4, 1, 1).setNumberFormat("$#,##0.00");
                weeklySheet.getRange(currentRow + tableData.length, 5, 1, 1).setNumberFormat("0.0");
                weeklySheet.getRange(currentRow + tableData.length, 6, 1, 2).setNumberFormat("$#,##0.00");

                const tableEnd = currentRow + tableData.length;

                // Chart 1: Total Credit by Driver (Col I = col 9, width 480px, height 310px)
                try {
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
                } catch (eChart1) {
                    Logger.log("Error inserting Chart 1 for week " + displayedWeekCounter + ": " + eChart1);
                }

                // Chart 2: Trips by Driver (Col N = col 14, width 480px, height 310px)
                try {
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
                } catch (eChart2) {
                    Logger.log("Error inserting Chart 2 for week " + displayedWeekCounter + ": " + eChart2);
                }

                currentRow = Math.max(tableEnd + 2, startRow + 16);
            }
        }
    });

    if (displayedWeekCounter === 0) {
        weeklySheet.getRange("A3:G3").merge()
            .setValue("No weekly data available yet for this month.")
            .setFontStyle("italic")
            .setHorizontalAlignment("center")
            .setBackground("#f9f9f9");
    } else {
        weeklySheet.autoResizeColumns(1, 7);
        for (let col = 1; col <= 7; col++) {
            weeklySheet.setColumnWidth(col, Math.max(weeklySheet.getColumnWidth(col) + 15, 110));
        }
    }
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
    if (!summarySheet) {
        Logger.log("generateBonusReport: Summary sheet not found.");
        return;
    }
    const bonusSheetName = "Bonus";

    // 1. --- Get Main Data from Summary (WITH DATE FILL-DOWN) ---
    const lastRowCurrent = summarySheet.getLastRow();
    let allRows = [];
    if (lastRowCurrent >= 3) {
        const rawRows = summarySheet.getRange(3, 1, lastRowCurrent - 2, 4).getValues();

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

    // 2. --- Month and Year Detection ---
    const { monthIndex: currentMonthIndex, year: currentYear, monthName: currentMonthName } = getSpreadsheetMonthAndYear(ss, allRows);

    // 3. --- Look Back Logic (WITH DATE FILL-DOWN) ---
    const firstOfMonth = new Date(currentYear, currentMonthIndex, 1);
    const firstWeekMonday = getMonday(firstOfMonth);

    if (firstWeekMonday.getMonth() < currentMonthIndex || firstWeekMonday.getFullYear() < firstOfMonth.getFullYear()) {
        try {
            const prevMonthDate = new Date(firstOfMonth);
            prevMonthDate.setMonth(currentMonthIndex - 1);
            const prevMonthName = prevMonthDate.toLocaleString('en-US', { month: 'long' });

            let files = DriveApp.getFilesByName(`${prevMonthName} - Drivers Daily Balance`);
            if (!files.hasNext()) {
                files = DriveApp.getFilesByName(`Drivers Daily Balance - ${prevMonthName}`);
            }
            if (!files.hasNext()) {
                files = DriveApp.getFilesByName(`${prevMonthName} ${prevMonthDate.getFullYear()} - Drivers Daily Balance`);
            }

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
        const credit = parseNumber(row[creditIdx]) || 0;

        const weekStart = getMonday(date);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekKey = `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;

        if (!weeklyData[weekKey]) {
            weeklyData[weekKey] = {
                startDate: weekStart,
                endDate: weekEnd,
                drivers: {}
            };
        }
        if (!weeklyData[weekKey].drivers[driver]) {
            weeklyData[weekKey].drivers[driver] = { credit: 0 };
        }
        weeklyData[weekKey].drivers[driver].credit += credit;
    });

    const includeNonDispatchInBonus = true;
    let allQualifiedDrivers = [];

    Object.keys(weeklyData).forEach(weekKey => {
        const weekBlock = weeklyData[weekKey];
        const weekStartDate = weekBlock.startDate;
        const weekEndDate = weekBlock.endDate;

        const isCurrentMonthWeek = (weekEndDate && weekEndDate.getMonth() === currentMonthIndex && weekEndDate.getFullYear() === currentYear) ||
            (weekStartDate && weekStartDate.getMonth() === currentMonthIndex && weekStartDate.getFullYear() === currentYear);

        if (isCurrentMonthWeek) {
            for (const driver in weekBlock.drivers) {
                const info = weekBlock.drivers[driver];
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

    // Ensure bonusSheet has at least 3 columns
    if (bonusSheet.getMaxColumns() < 3) {
        bonusSheet.insertColumnsAfter(bonusSheet.getMaxColumns(), 3 - bonusSheet.getMaxColumns());
    }

    bonusSheet.clear();

    const headerRange = bonusSheet.getRange("A1:C1");
    headerRange.setValues([["Driver", "Total Credit", "Week Period"]])
        .setFontWeight("bold")
        .setFontColor("white")
        .setBackground("#3c78d8")
        .setHorizontalAlignment("center")
        .setVerticalAlignment("middle")
        .setFontSize(11);

    if (allQualifiedDrivers.length > 0) {
        allQualifiedDrivers.sort((a, b) => {
            const blockA = weeklyData[a[2]];
            const blockB = weeklyData[b[2]];
            if (blockA && blockB) {
                const diff = blockA.startDate.getTime() - blockB.startDate.getTime();
                if (diff !== 0) return diff;
            }
            return String(a[0]).localeCompare(String(b[0]));
        });

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
        bonusSheet.getRange("A2:C2").merge()
            .setValue("No drivers have reached the threshold yet for this period.")
            .setFontStyle("italic")
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setBackground("#f9f9f9");
    }

    bonusSheet.autoResizeColumns(1, 3);
    bonusSheet.setColumnWidth(1, Math.max(bonusSheet.getColumnWidth(1) + 20, 140));
    bonusSheet.setColumnWidth(2, Math.max(bonusSheet.getColumnWidth(2) + 20, 120));
    bonusSheet.setColumnWidth(3, Math.max(bonusSheet.getColumnWidth(3) + 20, 180));
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
    return parseDate(value);
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
    const lastRow = summarySheet.getLastRow();
    let sampleRows = [];
    if (lastRow >= 3) {
        sampleRows = summarySheet.getRange(3, 1, Math.min(lastRow - 2, 30), 2).getValues();
    }
    const { monthIndex: currentMonthIndex, year: currentYear, monthName: currentMonthName } = getSpreadsheetMonthAndYear(ss, sampleRows);

    const firstOfMonth = new Date(currentYear, currentMonthIndex, 1);
    const endDate = new Date(currentYear, currentMonthIndex + 1, 0);
    const reportStartDate = getMonday(firstOfMonth);

    // 2. --- Initialize Data Map ---
    let angelDataMap = {};

    function mapRowToAngelData(row, colIndices) {
        const driverName = String(row[1] || "");
        const d = parseDate(row[0]);
        if (driverName.toLowerCase().includes(targetDriver.toLowerCase()) && d) {
            let dateKey = d.toDateString();
            const cIdx = colIndices ? colIndices.credit : 2;
            const tIdx = colIndices ? colIndices.trips : 5;
            const caIdx = colIndices ? colIndices.cash : 6;
            const nsIdx = colIndices ? colIndices.noShow : 7;
            angelDataMap[dateKey] = {
                credit: parseNumber(row[cIdx]) || 0,
                trips: parseNumber(row[tIdx]) || 0,
                cash: parseNumber(row[caIdx]) || 0,
                noShow: parseNumber(row[nsIdx]) || 0
            };
        }
    }

    // 3. --- Fetch Previous Month Data (Crossover Week) ---
    if (reportStartDate < firstOfMonth) {
        try {
            const prevMonthDate = new Date(firstOfMonth);
            prevMonthDate.setMonth(firstOfMonth.getMonth() - 1);
            const prevMonthName = prevMonthDate.toLocaleString('en-US', { month: 'long' });

            let files = DriveApp.getFilesByName(`${prevMonthName} - Drivers Daily Balance`);
            if (!files.hasNext()) {
                files = DriveApp.getFilesByName(`Drivers Daily Balance - ${prevMonthName}`);
            }
            if (!files.hasNext()) {
                files = DriveApp.getFilesByName(`${prevMonthName} ${prevMonthDate.getFullYear()} - Drivers Daily Balance`);
            }

            if (files.hasNext()) {
                const prevFile = files.next();
                const prevSpreadsheet = SpreadsheetApp.openById(prevFile.getId());
                const prevSummarySheet = prevSpreadsheet.getSheetByName("Summary");

                if (prevSummarySheet && prevSummarySheet.getLastRow() >= 3) {
                    const prevHeaders = prevSummarySheet.getRange(2, 1, 1, Math.min(prevSummarySheet.getLastColumn(), 10)).getDisplayValues()[0];
                    const prevIndices = {
                        credit: findMainSummaryCol(prevHeaders, ['credit'], 2),
                        trips: findMainSummaryCol(prevHeaders, ['trips', 'trip'], 6),
                        cash: findMainSummaryCol(prevHeaders, ['cash'], 7),
                        noShow: findMainSummaryCol(prevHeaders, ['no show', 'noshow'], 8)
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
        const curHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 10)).getDisplayValues()[0];
        const curIndices = {
            credit: findMainSummaryCol(curHeaders, ['credit'], 2),
            trips: findMainSummaryCol(curHeaders, ['trips', 'trip'], 6),
            cash: findMainSummaryCol(curHeaders, ['cash'], 7),
            noShow: findMainSummaryCol(curHeaders, ['no show', 'noshow'], 8)
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
    const s = String(val).replace(/[$,]/g, "").trim();
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
    if (value instanceof Date || Object.prototype.toString.call(value) === '[object Date]' || (typeof value === 'object' && typeof value.getTime === 'function')) {
        return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }
    if (typeof value === "string") {
        const str = value.trim();
        if (!str) return null;

        // 1. Strings with month names like "Tuesday, September 1, 2026"
        if (/[a-zA-Z]/.test(str)) {
            const std = new Date(str);
            if (!isNaN(std.getTime())) {
                return new Date(std.getFullYear(), std.getMonth(), std.getDate());
            }
        }

        // 2. Slash format: "MM/DD/YYYY" (from formatDate) or "DD/MM/YYYY"
        const slashParts = str.split("/");
        if (slashParts.length === 3) {
            const p0 = parseInt(slashParts[0], 10);
            const p1 = parseInt(slashParts[1], 10);
            let year = parseInt(slashParts[2], 10);
            if (year < 100) year += 2000;
            let month, day;
            if (p0 > 12) {
                // p0 cannot be month, must be DD/MM/YYYY
                day = p0;
                month = p1 - 1;
            } else if (p1 > 12) {
                // p1 cannot be month, must be MM/DD/YYYY
                month = p0 - 1;
                day = p1;
            } else {
                // Default to standard MM/DD/YYYY as produced by formatDate
                month = p0 - 1;
                day = p1;
            }
            if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                const d = new Date(year, month, day);
                if (!isNaN(d.getTime())) return d;
            }
        }

        // 3. Dash format: "DD-MM-YYYY", "DD-MM-YY", or "YYYY-MM-DD"
        const dashParts = str.split("-");
        if (dashParts.length === 3) {
            const p0 = parseInt(dashParts[0], 10);
            const p1 = parseInt(dashParts[1], 10);
            const p2 = parseInt(dashParts[2], 10);
            if (p0 > 1000) {
                return new Date(p0, p1 - 1, p2);
            }
            let year = p2 < 100 ? p2 + 2000 : p2;
            let month = p1 - 1;
            let day = p0;
            if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                return new Date(year, month, day);
            }
        }

        // 4. Dot format: "DD.MM.YYYY" or "DD.MM.YY" (e.g. "01.09.26")
        const dotParts = str.split(".");
        if (dotParts.length === 3) {
            const p0 = parseInt(dotParts[0], 10);
            const p1 = parseInt(dotParts[1], 10);
            let year = parseInt(dotParts[2], 10);
            if (year < 100) year += 2000;
            const day = p0;
            const month = p1 - 1;
            if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                const d = new Date(year, month, day);
                if (!isNaN(d.getTime())) return d;
            }
        }

        const fallback = new Date(str);
        if (!isNaN(fallback.getTime())) {
            return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
        }
    }
    return null;
}