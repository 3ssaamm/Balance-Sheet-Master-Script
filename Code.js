// =================================================================
// MASTER SCRIPT — CENTRALIZED BALANCE SHEET CONTROLLER
// =================================================================
// Change this ID each time you switch to a new monthly balance file.
// This is the ONLY line you need to update.
const TARGET_SHEET_ID = "1LG44Ry-BjAv2HCiZtkIcdI24zw43XbEH1WPDVCwUkC8";

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

    // --- Read the NET sheet (same 3-col summary structure as Raw Data) ---
    const netData    = netSheet.getDataRange().getValues();
    const netHeaders = netSheet.getRange(1, 1, 1, netSheet.getLastColumn()).getDisplayValues()[0];
    const numNetCols = netHeaders.length;

    // The NET sheet has the same trailing 3 summary columns as Raw Data:
    //   n-3 = Total NET  |  n-2 = Count  |  n-1 = Driver NET
    // We want "Driver NET" (n-1) as the definitive monthly NET per driver.
    // First try to find the column by header label; fall back to n-1.
    const netDriverNetColIdx = findColByLabel(netHeaders, ['driver net', 'net'], numNetCols - 1);
    Logger.log(`NET sheet: using column index ${netDriverNetColIdx} ("${netHeaders[netDriverNetColIdx]}") as Driver NET source.`);

    // Date columns: everything before the first non-date summary column
    const netDateCols = [];
    for (let c = 1; c < numNetCols - 3; c++) {   // stop 3 before end (same as Raw Data)
        const dateVal = parseDDMMYY(netHeaders[c]);
        if (dateVal) netDateCols.push({ colIdx: c, date: dateVal, headerStr: netHeaders[c] });
    }
    // Also try the last 3 cols in case the sheet has fewer summary cols
    if (netDateCols.length === 0) {
        for (let c = 1; c < numNetCols - 1; c++) {
            const dateVal = parseDDMMYY(netHeaders[c]);
            if (dateVal) netDateCols.push({ colIdx: c, date: dateVal, headerStr: netHeaders[c] });
        }
    }
    if (netDateCols.length === 0) {
        Logger.log('⚠️ No valid date columns found in "Raw Data - NET". Column headers must be DD-MM-YY (e.g., 01-06-26).');
        return;
    }

    // --- Build driver → row maps ---
    const netDriverRowMap = {};
    for (let r = 2; r < netData.length; r++) {
        const name = (netData[r][0] || '').toString().trim();
        if (name && name !== 'Total') netDriverRowMap[name] = r;
    }

    const rawHeaders = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getDisplayValues()[0];
    const rawNumCols = rawHeaders.length;
    const rawTotalColIdx = rawNumCols - 3;   // 0-based
    const rawNetColIdx   = rawNumCols - 1;   // 0-based ("Driver NET" in Raw Data)

    const rawData = rawSheet.getDataRange().getValues();
    const rawDriverRowMap = {};
    for (let r = 2; r < rawData.length; r++) {
        const name = (rawData[r][0] || '').toString().trim();
        if (name && name !== 'Total') rawDriverRowMap[name] = r;
    }

    // Build date-column map for Raw Data
    const rawDateColMap = {};
    for (let c = 1; c < rawTotalColIdx; c++) {
        const d = parseDDMMYY(rawHeaders[c]);
        if (d) rawDateColMap[d.getTime()] = c;
    }

    let insertedCols = 0;

    // --- Insert any missing date columns into Raw Data ---
    netDateCols.forEach(netDc => {
        const dateKey = netDc.date.getTime();
        if (rawDateColMap[dateKey] === undefined) {
            const insertAt = rawSheet.getLastColumn() - 2; // before summary cols
            rawSheet.insertColumnBefore(insertAt);
            rawSheet.getRange(1, insertAt).setValue(netDc.headerStr);
            // Fill credit = 0 for all driver rows (credit unknown from NET alone)
            for (let r = 2; r < rawData.length; r++) {
                const name = (rawData[r][0] || '').toString().trim();
                if (name) rawSheet.getRange(r + 1, insertAt).setValue(0);
            }
            rawDateColMap[dateKey] = insertAt - 1;
            insertedCols++;
            Logger.log(`📅 Inserted missing date column: ${netDc.headerStr}`);
        }
    });

    // --- Update "Driver NET" column in Raw Data with total from NET sheet ---
    // Re-read headers after possible column insertions
    const finalRawHeaders  = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getDisplayValues()[0];
    const finalRawNetColIdx = finalRawHeaders.length - 1; // last col = Driver NET (1-indexed: +1)

    let updatedDrivers = 0;
    for (const driverName in netDriverRowMap) {
        const netRowIdx = netDriverRowMap[driverName];
        const totalNet  = parseNumber(netData[netRowIdx][netDriverNetColIdx]) || 0;
        const rawRowIdx = rawDriverRowMap[driverName];
        if (rawRowIdx === undefined) {
            Logger.log(`⚠️ Driver "${driverName}" found in NET sheet but not in Raw Data — skipped.`);
            continue;
        }
        rawSheet.getRange(rawRowIdx + 1, finalRawNetColIdx + 1).setValue(totalNet);
        updatedDrivers++;
    }

    SpreadsheetApp.flush();
    Logger.log(`✅ NET Data import complete!`);
    Logger.log(`   • ${insertedCols} new date column(s) added to Raw Data.`);
    Logger.log(`   • Driver NET updated for ${updatedDrivers} driver(s).`);
    Logger.log(`   Running balance report...`);
    runDailyBalance();
}

// =================================================================
// IMPORT: TRIPS DATA — Merges "Raw Data - Trips" into "Raw Data".
//
// How it works:
//  1. Reads "Raw Data - Trips" — same layout as Raw Data but daily
//     values are trip counts per driver per day.
//  2. For each driver, the last column in the Trips sheet is the
//     Total Trips for the period.
//  3. Updates the "Count" column (2nd-to-last) in Raw Data with
//     the total from the Trips sheet.
//  4. After import, runs the full balance report automatically.
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

    // --- Read the Trips sheet (same 3-col summary structure as Raw Data) ---
    const tripsData    = tripsSheet.getDataRange().getValues();
    const tripsHeaders = tripsSheet.getRange(1, 1, 1, tripsSheet.getLastColumn()).getDisplayValues()[0];
    const numTripsCols = tripsHeaders.length;

    // The Trips sheet has the SAME trailing 3 summary columns as Raw Data:
    //   n-3 = Total (sum trips)  |  n-2 = Count (trips)  |  n-1 = Driver NET (irrelevant)
    // We want the COUNT column (n-2). Use header scanning to confirm; fall back to n-2.
    const tripsCountColIdx = findColByLabel(tripsHeaders, ['count', 'trip', 'trips', 'total'], numTripsCols - 2);
    Logger.log(`Trips sheet: using column index ${tripsCountColIdx} ("${tripsHeaders[tripsCountColIdx]}") as trip count source.`);

    // Build map: driver name → total trips (from the Count column of the Trips sheet)
    const driverTripsMap = {};
    for (let r = 2; r < tripsData.length; r++) {
        const name = (tripsData[r][0] || '').toString().trim();
        if (name && name !== 'Total') {
            driverTripsMap[name] = parseNumber(tripsData[r][tripsCountColIdx]) || 0;
        }
    }

    if (Object.keys(driverTripsMap).length === 0) {
        Logger.log('⚠️ No driver data found in "Raw Data - Trips". Make sure driver names are in column A.');
        return;
    }

    // Log what we found so the user can sanity-check
    Logger.log(`Found trips for ${Object.keys(driverTripsMap).length} driver(s):`);
    for (const d in driverTripsMap) Logger.log(`   ${d}: ${driverTripsMap[d]} trips`);

    // --- Update the Count column in Raw Data ---
    const rawData    = rawSheet.getDataRange().getValues();
    const rawNumCols = rawData[0].length;
    const rawCountCol1Idx = rawNumCols - 2; // 0-based; Count is 2nd-to-last in Raw Data

    let updatedCount = 0;
    for (let r = 2; r < rawData.length; r++) {
        const name = (rawData[r][0] || '').toString().trim();
        if (!name || name === 'Total') continue;
        if (driverTripsMap[name] !== undefined) {
            rawSheet.getRange(r + 1, rawCountCol1Idx + 1).setValue(driverTripsMap[name]); // 1-indexed
            updatedCount++;
        }
    }

    SpreadsheetApp.flush();
    Logger.log(`✅ Trips Data import complete! Updated ${updatedCount} driver(s).`);
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


/**
 * Main entry point — runs all reporting functions in the correct order.
 */
function runDailyBalance() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

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

    // --- Read the entire Raw Data sheet ---
    const rawData = rawSheet.getDataRange().getValues();
    if (rawData.length < 3) {
        Logger.log("⚠️ Raw Data has fewer than 3 rows (need header + Total + at least 1 driver).");
        return;
    }

    // Read headers as DISPLAY STRINGS so Google Sheets doesn't auto-convert DD-MM-YY dates
    const headerStrings = rawSheet.getRange(1, 1, 1, rawSheet.getLastColumn()).getDisplayValues()[0];
    const numCols = headerStrings.length;

    // --- Identify date columns vs summary columns ---
    // Last 3 columns are: Total, Count, Driver NET
    // Everything between column 1 and (numCols - 3) are date columns
    const totalColIdx = numCols - 3;
    const countColIdx = numCols - 2;
    const netColIdx = numCols - 1;

    // Parse date headers (DD-MM-YY format, read as strings)
    const dateColumns = []; // { colIdx, date }
    for (let c = 1; c < totalColIdx; c++) {
        const dateVal = parseDDMMYY(headerStrings[c]);
        if (dateVal) {
            dateColumns.push({ colIdx: c, date: dateVal });
        }
    }

    if (dateColumns.length === 0) {
        Logger.log("⚠️ No valid date columns found in Raw Data headers.");
        return;
    }

    // --- Build output rows: one row per driver per day (only days with credit > 0) ---
    let outputRows = []; // [date, driver, credit, cash, trips, total]

    for (let r = 2; r < rawData.length; r++) { // Skip header (0) and Total row (1)
        const driverName = rawData[r][0];
        if (!driverName || driverName.toString().trim() === "" || driverName === "Total") continue;

        const rawDriverTotal = parseNumber(rawData[r][totalColIdx]) || 0;
        
        let totalNoShowCashForDriver = 0;
        if (noShowMap[driverName]) {
            Object.values(noShowMap[driverName]).forEach(val => totalNoShowCashForDriver += val.cash);
        }
        
        const driverTotal = roundToTwo(rawDriverTotal - totalNoShowCashForDriver);
        const driverTrips = parseNumber(rawData[r][countColIdx]) || 0;
        const driverNet = parseNumber(rawData[r][netColIdx]) || 0;
        let driverCash = roundToTwo(driverTotal - driverNet); // Base Cash = Total - NET

        // Collect all active days (days with credit, trips, or cash)
        let activeDays = [];
        dateColumns.forEach(dc => {
            const dateKey = dc.date.getTime();
            let credit = parseNumber(rawData[r][dc.colIdx]) || 0;
            
            const hasExactTrips = tripsDataMap[driverName] && tripsDataMap[driverName][dateKey] !== undefined;
            const hasExactNet = netDataMap[driverName] && netDataMap[driverName][dateKey] !== undefined;
            
            const exactTrips = hasExactTrips ? tripsDataMap[driverName][dateKey] : 0;
            const exactNet = hasExactNet ? netDataMap[driverName][dateKey] : 0;
            const hasNoShow = noShowMap[driverName] && noShowMap[driverName][dateKey] !== undefined;

            if (hasNoShow) {
                credit = roundToTwo(credit - noShowMap[driverName][dateKey].cash);
            }

            // Include day if there's credit, trips, non-zero net (implying cash), or a No Show adjustment
            if (credit > 0 || exactTrips > 0 || exactNet !== 0 || hasNoShow) {
                activeDays.push({ date: dc.date, dateKey: dateKey, credit: credit });
            }
        });

        // Calculate exact sums for trips and cash
        let sumExactTrips = 0;
        let sumExactCash = 0;
        let missingTripDays = [];
        let missingCashDays = [];

        activeDays.forEach(day => {
            // Trips
            if (tripsDataMap[driverName] && tripsDataMap[driverName][day.dateKey] !== undefined) {
                sumExactTrips += tripsDataMap[driverName][day.dateKey];
            } else {
                missingTripDays.push(day);
            }

            // Cash
            if (netDataMap[driverName] && netDataMap[driverName][day.dateKey] !== undefined) {
                const dailyNet = netDataMap[driverName][day.dateKey];
                let exactDailyCash = roundToTwo(day.credit - dailyNet);
                sumExactCash += exactDailyCash;
            } else {
                missingCashDays.push(day);
            }
        });

        // Remainder available for missing days
        const remainingTrips = driverTrips - sumExactTrips;
        const remainingCash = driverCash - sumExactCash;

        // Build output rows
        activeDays.forEach(day => {
            let dailyTrips = 0;
            let dailyCash = 0;

            // TRIPS LOGIC
            if (tripsDataMap[driverName] && tripsDataMap[driverName][day.dateKey] !== undefined) {
                dailyTrips = tripsDataMap[driverName][day.dateKey]; // Exact match
            } else if (missingTripDays.length === 1) {
                dailyTrips = remainingTrips; // 1 missing day -> gets the remainder exactly
            } else {
                dailyTrips = 0; // >1 missing day -> leave blank/0 until user uploads Trips data
            }

            // CASH LOGIC
            if (netDataMap[driverName] && netDataMap[driverName][day.dateKey] !== undefined) {
                const dailyNet = netDataMap[driverName][day.dateKey];
                dailyCash = roundToTwo(day.credit - dailyNet); // Exact match
            } else if (missingCashDays.length === 1) {
                dailyCash = remainingCash; // 1 missing day -> gets the remainder exactly
            } else {
                dailyCash = 0; // >1 missing day -> leave blank/0 until user uploads NET data
            }

            let dailyNoShowCount = 0;
            if (noShowMap[driverName] && noShowMap[driverName][day.dateKey]) {
                dailyNoShowCount = noShowMap[driverName][day.dateKey].count;
            }

            outputRows.push([
                day.date,
                driverName,
                day.credit,
                dailyCash,
                dailyTrips,
                driverTotal,
                driverCash,
                dailyNoShowCount
            ]);
        });
    }

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
    const headers = ["Date", "Driver", "Credit", "Trips", "Cash", "No Show", "Balance"];
    summarySheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    // --- Read fare settings for balance calculation ---
    let settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) settingsSheet = ss.insertSheet("Settings");
    const settingsData = settingsSheet.getDataRange().getValues();
    let existingFares = {}; // driver -> fare value
    for (let i = 1; i < settingsData.length; i++) {
        if (settingsData[i][0]) {
            existingFares[settingsData[i][0]] = settingsData[i][1];
        }
    }

    // Helper to get fare as decimal
    function getFareDecimal(driver) {
        let fareVal = existingFares[driver];
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

        const fare = Number(driverFareCache[row[1]]) || 0.9;
        const dailyCredit = Number(row[2]) || 0;
        const dailyCash = Number(row[3]) || 0;
        const trips = Number(row[4]) || 0;
        const dailyNoShow = Number(row[7]) || 0;
        const dailyBalance = Number(roundToTwo((dailyCredit * fare) - dailyCash)) || 0;

        // Summary rows: Date, Driver, Credit, Trips, Cash, No Show, Balance
        rowsToWrite.push([
            displayDate,
            row[1],
            dailyCredit,
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
        summarySheet.getRange(3, 1, maxRows - 2, 7).clearContent();
        try { summarySheet.getRange(3, 1, maxRows - 2, 1).breakApart(); } catch (e) { }
    }

    const dataRange = summarySheet.getRange(3, 1, rowsToWrite.length, 7);
    dataRange.setValues(rowsToWrite);
    dataRange.setBorder(false, false, false, false, false, false);
    dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    summarySheet.getRange(3, 1, rowsToWrite.length, 1).setVerticalAlignment("middle");

    // --- Apply day block borders and merge date cells ---
    dayBlocks.forEach(block => {
        summarySheet.getRange(block.row, 1, block.count, 7)
            .setBorder(true, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        if (block.count > 1) {
            summarySheet.getRange(block.row, 1, block.count, 1).mergeVertically();
        }
    });

    // --- Aggregate per driver (for the stats table) ---
    let driverStats = {};
    // We need to aggregate from raw data directly (not outputRows) to get accurate totals
    for (let r = 2; r < rawData.length; r++) {
        const driverName = rawData[r][0];
        if (!driverName || driverName.toString().trim() === "" || driverName === "Total") continue;

        const driverTotal = parseNumber(rawData[r][totalColIdx]) || 0;
        const driverTrips = parseNumber(rawData[r][countColIdx]) || 0;
        const driverNet = parseNumber(rawData[r][netColIdx]) || 0;
        let driverCash = roundToTwo(driverTotal - driverNet);

        // Deduct ALL No Show CASH for this driver from their monthly total cash
        let totalNoShowCashForDriver = 0;
        let totalNoShowCountForDriver = 0;
        if (noShowMap[driverName]) {
            Object.values(noShowMap[driverName]).forEach(val => {
                totalNoShowCashForDriver += val.cash;
                totalNoShowCountForDriver += val.count;
            });
        }
        driverCash = roundToTwo(driverCash - totalNoShowCashForDriver);

        if (driverTotal <= 0) continue;

        // Count working days for this driver
        let workingDays = 0;
        let firstWorkDate = null;
        dateColumns.forEach(dc => {
            const credit = parseNumber(rawData[r][dc.colIdx]) || 0;
            if (credit > 0) {
                workingDays++;
                if (!firstWorkDate) firstWorkDate = dc.date;
            }
        });

        driverStats[driverName] = {
            totalCredit: roundToTwo(driverTotal),
            totalTrips: driverTrips,
            totalCash: driverCash,
            totalNoShow: totalNoShowCountForDriver,
            driverNet: roundToTwo(driverNet),
            workingDays: workingDays,
            avgPerDay: workingDays > 0 ? roundToTwo(driverTotal / workingDays) : 0,
            firstWorkDate: firstWorkDate
        };
    }

    // =================================================================
    // --- AUTO-SYNC "SETTINGS" SHEET (Driver + Fare %) ---
    // =================================================================
    // (settingsSheet and existingFares already read above for balance calc)

    // --- Aggregated stats table (column I, one gap after the 7-col main table) ---
    const aggHeader = ["Driver", "Total Credit", "Total Trips", "Total Cash", "No Show", "Driver NET", "Avg Per Day"];
    const aggRows = [aggHeader];

    const newSettingsRows = [["Driver", "Fare %"]];

    for (let driver in driverStats) {
        const s = driverStats[driver];
        aggRows.push([driver, s.totalCredit, s.totalTrips, s.totalCash, s.totalNoShow, s.driverNet, s.avgPerDay]);

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
    // --- Write aggregated stats to Summary (column I = one gap after 7-col main table) ---
    // =================================================================
    const startRow = 3, startCol = 9; // Column I (A-G = main table, H = gap)
    // Clear old aggregate area
    try { summarySheet.getRange(startRow, startCol, 50, 10).clearContent(); } catch (e) { }

    summarySheet.getRange(startRow, startCol, aggRows.length, aggRows[0].length).setValues(aggRows);

    // --- Daily stats for chart (aggregate all drivers per day) ---
    let dailyStats = {};
    outputRows.forEach(r => {
        const dateVal = r[0];
        if (!dateVal) return;
        const dateKey = dateVal.getTime();
        if (!dailyStats[dateKey]) {
            dailyStats[dateKey] = { date: dateVal, credit: 0, trips: 0, cash: 0, noShow: 0 };
        }
        dailyStats[dateKey].credit += (r[2] || 0);
        dailyStats[dateKey].trips += (r[4] || 0);
        dailyStats[dateKey].cash += (r[3] || 0);
        dailyStats[dateKey].noShow += (r[7] || 0); // r[7] is No Show
    });

    const dailyRows = Object.values(dailyStats)
        .sort((a, b) => a.date - b.date)
        .map(d => [d.date, roundToTwo(d.credit), d.trips, roundToTwo(d.cash), d.noShow]);

    const dailyHeader = ["Date", "Total Credit", "Trips", "Cash", "No Show"];
    const dailyStartCol = startCol + aggRows[0].length + 1; // One gap after aggregate table

    summarySheet.getRange(2, dailyStartCol, 100, 7).clearContent();
    summarySheet.getRange(2, dailyStartCol, 1, dailyHeader.length).setValues([dailyHeader]).setFontWeight("bold");

    if (dailyRows.length > 0) {
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, dailyRows[0].length).setValues(dailyRows);
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, 1).setNumberFormat("dddd, dd");
    }

    // --- Charts ---
    const allCharts = summarySheet.getCharts();
    allCharts.forEach(c => summarySheet.removeChart(c));

    // Chart 1: Daily Total Credit Trend (line) — positioned NEXT TO the daily credit table
    if (dailyRows.length > 0) {
        const dailyChartCol = dailyStartCol + 6; // Shifted right to accommodate the 5-column daily table
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
            .setOption("width", 900)
            .setOption("height", 350)
            .build();
        summarySheet.insertChart(dailyChart);
    }

    const lastRow = startRow + aggRows.length - 1;
    const dataStart = startRow + 1;
    // Build column letters from startCol for chart ranges
    const colF = String.fromCharCode(64 + startCol);         // I (Driver)
    const colG = String.fromCharCode(64 + startCol + 1);     // J (Total Credit)
    const colH = String.fromCharCode(64 + startCol + 2);     // K (Total Trips)
    const colNoShow = String.fromCharCode(64 + startCol + 4); // M (No Show)
    const colK = String.fromCharCode(64 + startCol + 6);     // O (Avg Per Day) - 6 columns offset

    // Chart 2: Total Credit by Driver (bar)
    let chart1 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colG + dataStart + ":" + colG + lastRow))
        .setOption("title", "Total Credit by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#1f77b4"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Credit" })
        .setOption("width", 720)
        .setOption("height", 450)
        .setPosition(3, startCol, 0, 0).build();
    summarySheet.insertChart(chart1);

    // Chart 3: Total Trips by Driver (bar)
    let chart2 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colH + dataStart + ":" + colH + lastRow))
        .setOption("title", "Total Trips by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#8c564b"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Trips" })
        .setOption("width", 720)
        .setOption("height", 450)
        .setPosition(27, startCol, 0, 0).build();
    summarySheet.insertChart(chart2);

    // Chart 4: No Show by Driver (bar)
    let chartNoShow = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colNoShow + dataStart + ":" + colNoShow + lastRow))
        .setOption("title", "No Show Trips by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#2ca02c"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "No Show Trips" })
        .setOption("width", 720)
        .setOption("height", 450)
        .setPosition(51, startCol, 0, 0).build();
    summarySheet.insertChart(chartNoShow);

    // Chart 5: Average Per Day by Driver (bar)
    let chart3 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange(colF + dataStart + ":" + colF + lastRow))
        .addRange(summarySheet.getRange(colK + dataStart + ":" + colK + lastRow))
        .setOption("title", "Average Credit Per Day")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#d62728"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Daily Credit" })
        .setOption("width", 720)
        .setOption("height", 450)
        .setPosition(75, startCol, 0, 0).build();
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
                if (typeof fareVal === "string" && fareVal.includes("%")) {
                    fareVal = parseFloat(fareVal) / 100;
                } else if (typeof fareVal === "number") {
                    fareVal = fareVal > 1 ? fareVal / 100 : fareVal;
                } else {
                    fareVal = 0.9; // default
                }
                driverFareMap[settingsData[i][0]] = fareVal;
            }
        }
    }

    // --- Get Main Data from Summary sheet ---
    const lastRowCurrent = summarySheet.getLastRow();
    let allRows = [];
    if (lastRowCurrent >= 3) {
        const rawRows = summarySheet.getRange("A3:E" + lastRowCurrent).getValues();
        let lastSeenDate = null;
        allRows = rawRows.map(r => {
            if (r[0] && r[0] !== "") lastSeenDate = r[0];
            else r[0] = lastSeenDate;
            return r;
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
                    const oldLastCol = prevSummarySheet.getLastColumn();
                    const readCols = oldLastCol >= 5 ? 5 : oldLastCol;
                    const prevAllData = prevSummarySheet.getRange(3, 1, prevSummarySheet.getLastRow() - 2, readCols).getValues();

                    let prevLastSeen = null;
                    const crossoverRows = prevAllData.map(row => {
                        if (row[0] && row[0] !== "") prevLastSeen = row[0];
                        else row[0] = prevLastSeen;
                        return row;
                    }).filter(row => {
                        const d = parseDate(row[0]);
                        return d && d >= firstWeekMonday && d < firstOfMonth;
                    });
                    allRows = [...crossoverRows, ...allRows];
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

        if (!weeklyData[weekKey]) weeklyData[weekKey] = {};
        if (!weeklyData[weekKey][driver]) {
            weeklyData[weekKey][driver] = { credit: 0, trips: 0, cash: 0 };
        }

        weeklyData[weekKey][driver].credit += credit;
        weeklyData[weekKey][driver].trips += trips;
        weeklyData[weekKey][driver].cash += cash;
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

            const titleRange = weeklySheet.getRange(currentRow, 1, 1, 5).merge();
            titleRange.setValue(`WEEK ${displayedWeekCounter} (${weekKey})`)
                .setFontWeight("bold")
                .setFontSize(11)
                .setBackground("#d0e0e3");

            currentRow++;

            const headerRange = weeklySheet.getRange(currentRow, 1, 1, 5);
            headerRange.setValues([["Driver", "Total Credit", "Trips", "Total Cash", "Balance"]])
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
                const fare = driverFareMap[d] || 0.9; // default 90%
                const balance = roundToTwo((info.credit * fare) - cash);
                tableData.push([d, info.credit, info.trips, cash, balance]);
            });

            // Write Main Table
            const dataRange = weeklySheet.getRange(currentRow, 1, tableData.length, 5);
            dataRange.setValues(tableData);
            dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            weeklySheet.getRange(currentRow, 2, tableData.length, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow, 4, tableData.length, 2).setNumberFormat("$#,##0.00");

            const totalCredit = tableData.reduce((a, b) => a + b[1], 0);
            const totalTrips = tableData.reduce((a, b) => a + b[2], 0);
            const totalCash = tableData.reduce((a, b) => a + b[3], 0);
            const totalBalance = tableData.reduce((a, b) => a + b[4], 0);

            const totalRowRange = weeklySheet.getRange(currentRow + tableData.length, 1, 1, 5);
            totalRowRange.setValues([["TOTAL", totalCredit, totalTrips, totalCash, totalBalance]])
                .setFontWeight("bold")
                .setBackground("#eeeeee")
                .setBorder(true, true, true, true, true, true);

            weeklySheet.getRange(currentRow + tableData.length, 2, 1, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow + tableData.length, 4, 1, 2).setNumberFormat("$#,##0.00");

            const tableEnd = currentRow + tableData.length;

            // Chart 1: Total Credit by Driver
            const chart1 = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 7, 0, 0)
                .addRange(weeklySheet.getRange(currentRow, 1, tableData.length, 2))
                .setOption("title", `WEEK ${displayedWeekCounter} - Total Credit`)
                .setOption("colors", ["#1f77b4"])
                .setOption("legend", { position: "none" })
                .setOption("hAxis", { title: "Driver" })
                .setOption("vAxis", { title: "Total Credit" })
                .build();
            weeklySheet.insertChart(chart1);

            // Chart 2: Trips by Driver
            const chart2 = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 13, 0, 0) // Place at column M
                .addRange(weeklySheet.getRange(currentRow, 1, tableData.length, 1)) // Driver Names
                .addRange(weeklySheet.getRange(currentRow, 3, tableData.length, 1)) // Trips
                .setOption("title", `WEEK ${displayedWeekCounter} - Trips`)
                .setOption("colors", ["#8c564b"]) // Distinct brown color
                .setOption("legend", { position: "none" })
                .setOption("hAxis", { title: "Driver" })
                .setOption("vAxis", { title: "Trips" })
                .build();
            weeklySheet.insertChart(chart2);

            currentRow = Math.max(tableEnd + 2, startRow + 22);
        }
    });
}


// =================================================================
// BONUS REPORT — Identifies drivers who hit $1500 weekly threshold
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

                    let prevLastSeen = null;
                    const crossoverRows = prevAllData.map(row => {
                        if (row[0] && row[0] !== "") prevLastSeen = row[0];
                        else row[0] = prevLastSeen;
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

                if (nonDispatchDrivers.has(driver)) {
                    if (includeNonDispatchInBonus === true && info.credit >= 1500) {
                        allQualifiedDrivers.push([driver, info.credit, weekKey]);
                    }
                    continue;
                }

                if (info.credit >= 1500) {
                    allQualifiedDrivers.push([driver, info.credit, weekKey]);
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

        bonusSheet.getRange(2, 2, allQualifiedDrivers.length, 1).setNumberFormat("$#,##0.00");

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

        bonusSheet.getRange(totalRow, 2, 2, 1).setNumberFormat("$#,##0.00");

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

  function mapRowToAngelData(row) {
    if (row[1] === targetDriver && row[0] instanceof Date) {
      let dateKey = row[0].toDateString(); 
      angelDataMap[dateKey] = {
        credit: Number(row[2]) || 0,
        trips: Number(row[3]) || 0,
        cash: Number(row[4]) || 0,
        noShow: Number(row[5]) || 0
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
          const prevData = prevSummarySheet.getRange(3, 1, prevSummarySheet.getLastRow() - 2, 11).getValues();
          
          let prevLastSeenDate = null;
          prevData.forEach(row => {
            if (row[0] && row[0] !== "") prevLastSeenDate = row[0];
            else row[0] = prevLastSeenDate;

            const d = parseDate(row[0]);
            if (d && d >= reportStartDate) {
               row[0] = d; 
               mapRowToAngelData(row);
            }
          });
        }
      }
    } catch (e) { }
  }

  // 4. --- Fetch Current Month Data ---
  if (lastRow >= 3) {
    const currentData = summarySheet.getRange(3, 1, lastRow - 2, 11).getValues();
    
    let currentLastSeenDate = null;
    currentData.forEach(row => {
        if (row[0] && row[0] !== "") currentLastSeenDate = row[0];
        else row[0] = currentLastSeenDate;

        if (row[0]) { 
            const d = parseDate(row[0]);
            if(d) {
                row[0] = d;
                mapRowToAngelData(row);
            }
        }
    });
  }

  // 5. --- Build the Report Rows ---
  let reportRows = [];
  let chartDataRows =[]; 
  let weekCounter = 1;
  
  let weeklyCredit = 0, weeklyNoShow = 0, weeklyTrips = 0, weeklyCash = 0;
  let weeklyGrossBal = 0, weeklyFee = 0, weeklyNet = 0;
  
  let grandCredit = 0, grandNoShow = 0, grandTrips = 0, grandCash = 0;
  let grandGrossBal = 0, grandFee = 0, grandNet = 0;

  let weekTotalRows =[];

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
    try { reportSheet.showColumns(1, 20); } catch(e) {}
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
      .setOption("colors",["#1f77b4"])
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