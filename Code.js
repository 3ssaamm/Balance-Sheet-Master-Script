// =================================================================
// MASTER SCRIPT — CENTRALIZED BALANCE SHEET CONTROLLER
// =================================================================
// Change this ID each time you switch to a new monthly balance file.
// This is the ONLY line you need to update.
const TARGET_SHEET_ID = "1XmzqDfalrqADK-O7g1SsTOersaRBNzACa-KNcEJVrP9v7w9BQ_hkC5Uu";
const DAILY_EXTRACTS_FOLDER_ID = "1wTRKJMQ07vGT4ZzzoXWFz92hLoWRkVvw";

var _ACTIVE_MONTH_NUM = 10; // Default: October (10)

/**
 * Safely resolves the active spreadsheet without throwing if TARGET_SHEET_ID is invalid.
 */
function getSafeSpreadsheet(ss) {
    if (ss) {
        if (typeof ss.getParent === "function" && typeof ss.getLastRow === "function") {
            try { return ss.getParent(); } catch (e) { }
        }
        if (typeof ss.getSheetByName === "function") {
            return ss;
        }
    }
    try {
        const active = SpreadsheetApp.getActiveSpreadsheet();
        if (active) return active;
    } catch (e) { }
    try {
        if (TARGET_SHEET_ID && TARGET_SHEET_ID.length === 44) {
            return SpreadsheetApp.openById(TARGET_SHEET_ID);
        }
    } catch (e) { }
    return null;
}

/**
 * Detects the active month and year context from the spreadsheet name, sheet names, or system date.
 */
function getActiveMonthContext(ss) {
    let spreadsheet = getSafeSpreadsheet(ss);
    let sheetNameHint = "";

    if (ss && typeof ss.getParent === "function" && typeof ss.getLastRow === "function") {
        try { sheetNameHint = ss.getName(); } catch (e) { }
    }

    const monthNames = [
        "January", "February", "March", "April", "May", "June",
        "July", "August", "September", "October", "November", "December"
    ];
    const shortNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];

    const ssName = (spreadsheet && typeof spreadsheet.getName === "function") ? spreadsheet.getName() : "";
    let foundMonthIndex = -1;

    // 1. Check spreadsheet name (e.g. "October - Drivers Daily Balance")
    for (let i = 0; i < monthNames.length; i++) {
        if (new RegExp("\\b" + monthNames[i] + "\\b", "i").test(ssName)) {
            foundMonthIndex = i;
            break;
        }
    }
    if (foundMonthIndex === -1) {
        for (let i = 0; i < shortNames.length; i++) {
            if (new RegExp("\\b" + shortNames[i] + "\\b", "i").test(ssName)) {
                foundMonthIndex = i;
                break;
            }
        }
    }
    if (foundMonthIndex === -1) {
        const ymMatch = ssName.match(/\b20\d\d[-_](0?[1-9]|1[0-2])\b/);
        if (ymMatch) foundMonthIndex = parseInt(ymMatch[1], 10) - 1;
    }

    // 2. Check sheetNameHint if available
    if (foundMonthIndex === -1 && sheetNameHint) {
        for (let i = 0; i < monthNames.length; i++) {
            if (new RegExp("\\b" + monthNames[i] + "\\b", "i").test(sheetNameHint)) {
                foundMonthIndex = i;
                break;
            }
        }
    }

    // 3. Check names of sheets within the spreadsheet
    if (foundMonthIndex === -1 && spreadsheet && typeof spreadsheet.getSheets === "function") {
        try {
            const allSheets = spreadsheet.getSheets();
            for (let s = 0; s < allSheets.length; s++) {
                const sName = allSheets[s].getName();
                for (let i = 0; i < monthNames.length; i++) {
                    if (new RegExp("\\b" + monthNames[i] + "\\b", "i").test(sName)) {
                        foundMonthIndex = i;
                        break;
                    }
                }
                if (foundMonthIndex !== -1) break;
            }
        } catch (eSheets) { }
    }

    let foundYear = -1;
    const yearMatch = (ssName + " " + sheetNameHint).match(/\b(20\d\d)\b/);
    if (yearMatch) foundYear = parseInt(yearMatch[1], 10);

    // 4. If year not found in title, inspect the Summary sheet date column
    if (foundYear === -1 && spreadsheet && typeof spreadsheet.getSheetByName === "function") {
        try {
            const sumSheet = spreadsheet.getSheetByName("Summary");
            if (sumSheet && sumSheet.getLastRow() >= 3) {
                const sampleVals = sumSheet.getRange(3, 1, Math.min(15, sumSheet.getLastRow() - 2), 1).getValues();
                for (let r = 0; r < sampleVals.length; r++) {
                    const cellVal = sampleVals[r][0];
                    if (cellVal instanceof Date && !isNaN(cellVal.getTime())) {
                        const y = cellVal.getFullYear();
                        if (y >= 2020 && y <= 2035) {
                            foundYear = y;
                            break;
                        }
                    } else if (cellVal) {
                        const m = cellVal.toString().match(/\b(20\d\d)\b/);
                        if (m) {
                            foundYear = parseInt(m[1], 10);
                            break;
                        }
                    }
                }
            }
        } catch (eYear) { }
    }

    const now = new Date();
    const finalMonthIndex = foundMonthIndex >= 0 ? foundMonthIndex : now.getMonth();
    const finalYear = foundYear > 2000 ? foundYear : (now.getFullYear() < 2026 ? 2026 : now.getFullYear());

    _ACTIVE_MONTH_NUM = finalMonthIndex + 1;

    return {
        monthIndex: finalMonthIndex,
        monthNum: finalMonthIndex + 1,
        year: finalYear,
        monthName: monthNames[finalMonthIndex]
    };
}

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
            .addItem('🔧 Setup Auto-Run Trigger', 'setupAutoRun')
            .addToUi();
    } catch (e) {
        // onOpen() cannot be run from the Apps Script editor.
        // It runs automatically when the spreadsheet is opened in a browser.
        Logger.log("ℹ️ onOpen() skipped: " + e.message);
    }
}

// =================================================================
// DAILY EXTRACTS LOADER — Reads daily .xlsx files directly from Drive
// =================================================================

/**
 * Converts column letter(s) like "A", "B", "AA" to 0-based column index.
 */
function colLetterToIdx(letters) {
    if (!letters) return 0;
    let col = 0;
    const upper = letters.toUpperCase();
    for (let i = 0; i < upper.length; i++) {
        col = col * 26 + (upper.charCodeAt(i) - 64);
    }
    return col - 1;
}

/**
 * Decodes standard XML escape entities.
 */
function unescapeXml(str) {
    if (!str) return "";
    return str
        .replace(/&amp;/g, '&')
        .replace(/&lt;/g, '<')
        .replace(/&gt;/g, '>')
        .replace(/&quot;/g, '"')
        .replace(/&apos;/g, "'")
        .replace(/&#(\d+);/g, (m, dec) => String.fromCharCode(dec))
        .replace(/&#x([0-9a-fA-F]+);/g, (m, hex) => String.fromCharCode(parseInt(hex, 16)));
}

/**
 * Unzips an Excel .xlsx blob and parses the first worksheet into a 2D array.
 * Works natively in Google Apps Script without converting to Google Sheets or calling external APIs.
 */
function parseXlsxBlob(blob) {
    const unzipped = Utilities.unzip(blob);
    let sharedStringsXml = null;
    let sheetXml = null;

    for (let i = 0; i < unzipped.length; i++) {
        const entryName = unzipped[i].getName();
        if (entryName === "xl/sharedStrings.xml") {
            sharedStringsXml = unzipped[i].getDataAsString();
        } else if (entryName === "xl/worksheets/sheet1.xml" || (!sheetXml && entryName.startsWith("xl/worksheets/sheet") && entryName.endsWith(".xml"))) {
            sheetXml = unzipped[i].getDataAsString();
        }
    }

    if (!sheetXml) {
        throw new Error("Could not find worksheet XML inside Excel file.");
    }

    // 1. Parse shared strings table
    const sharedStrings = [];
    if (sharedStringsXml) {
        const siMatches = sharedStringsXml.match(/<si[\s\S]*?<\/si>/g) || [];
        for (let j = 0; j < siMatches.length; j++) {
            const tMatches = siMatches[j].match(/<t[^>]*>([\s\S]*?)<\/t>/g) || [];
            let str = "";
            for (let k = 0; k < tMatches.length; k++) {
                const textContent = tMatches[k].replace(/^<t[^>]*>/, '').replace(/<\/t>$/, '');
                str += unescapeXml(textContent);
            }
            sharedStrings.push(str);
        }
    }

    // 2. Parse sheet cells into grid
    const grid = [];
    const rowMatches = sheetXml.match(/<row[\s\S]*?<\/row>/g) || [];

    for (let r = 0; r < rowMatches.length; r++) {
        const rowStr = rowMatches[r];
        const rNumMatch = rowStr.match(/r="(\d+)"/);
        const rowIdx = rNumMatch ? parseInt(rNumMatch[1], 10) - 1 : r;

        while (grid.length <= rowIdx) grid.push([]);

        const cMatches = rowStr.match(/<c[\s\S]*?(?:<\/c>|\/>)/g) || [];
        for (let c = 0; c < cMatches.length; c++) {
            const cellStr = cMatches[c];
            const refMatch = cellStr.match(/r="([A-Z]+)(\d+)"/);
            if (!refMatch) continue;
            const colLetters = refMatch[1];
            const colIdx = colLetterToIdx(colLetters);

            const typeMatch = cellStr.match(/t="([^"]+)"/);
            const cellType = typeMatch ? typeMatch[1] : "";

            let val = "";
            if (cellType === "s") {
                const vMatch = cellStr.match(/<v>(\d+)<\/v>/);
                if (vMatch) {
                    const sIdx = parseInt(vMatch[1], 10);
                    val = sharedStrings[sIdx] !== undefined ? sharedStrings[sIdx] : "";
                }
            } else if (cellType === "inlineStr") {
                const tMatch = cellStr.match(/<t[^>]*>([\s\S]*?)<\/t>/);
                if (tMatch) val = unescapeXml(tMatch[1]);
            } else {
                const vMatch = cellStr.match(/<v>([\s\S]*?)<\/v>/);
                if (vMatch) val = unescapeXml(vMatch[1]);
            }

            while (grid[rowIdx].length <= colIdx) grid[rowIdx].push("");
            grid[rowIdx][colIdx] = val;
        }
    }

    return grid;
}

/**
 * Parses the "Period: MM/DD/YYYY – MM/DD/YYYY" line from a daily extract.
 */
function parsePeriodLine(text, activeMonthNum) {
    if (!text) return null;
    const clean = String(text).replace(/[\u2010-\u2015\u2212]/g, '-').trim();
    const match = clean.match(/period:\s*([0-9\/\.\-]+)\s*[\-–—to\s]+\s*([0-9\/\.\-]+)/i);
    if (match) {
        const d1 = parseDate(match[1], activeMonthNum);
        const d2 = parseDate(match[2], activeMonthNum);
        if (d1 && d2) {
            return { startDate: d1, endDate: d2, startStr: match[1], endStr: match[2] };
        }
    }
    return null;
}

/**
 * Attempts to parse a date from the file name as a fallback.
 */
function parseDateFromFileName(fileName, activeMonthNum) {
    if (!fileName) return null;
    const clean = fileName.replace(/[\u2010-\u2015\u2212]/g, '-');
    const isoMatch = clean.match(/\b(20\d\d)[-_](0?[1-9]|1[0-2])[-_](0?[1-9]|[12]\d|3[01])\b/);
    if (isoMatch) {
        return new Date(parseInt(isoMatch[1], 10), parseInt(isoMatch[2], 10) - 1, parseInt(isoMatch[3], 10));
    }
    const usMatch = clean.match(/\b(0?[1-9]|1[0-2])[-_](0?[1-9]|[12]\d|3[01])[-_](20\d\d|\d\d)\b/);
    if (usMatch) {
        let yr = parseInt(usMatch[3], 10);
        if (yr < 100) yr += 2000;
        return new Date(yr, parseInt(usMatch[1], 10) - 1, parseInt(usMatch[2], 10));
    }
    return parseDate(fileName, activeMonthNum);
}

/**
 * Extracts driver records from an Excel grid.
 * Total credit = cash + final pay + amount of fees (positive value)
 * Columns captured: Name, Trips, Fees, Tollways, Final Pay, Cash
 */
function extractDailyRecordsFromGrid(grid, fileName, activeMonthNum) {
    if (!grid || grid.length === 0) {
        return { error: "Empty worksheet" };
    }

    // 1. Search top 15 rows for Period line
    let fileDate = null;
    let periodStr = "";
    let isMultiDay = false;

    for (let r = 0; r < Math.min(grid.length, 15); r++) {
        const rowText = grid[r].join(" ");
        if (/period:/i.test(rowText)) {
            const pMatch = rowText.match(/period:[^\r\n]+/i);
            periodStr = pMatch ? pMatch[0] : rowText;
            const p = parsePeriodLine(periodStr, activeMonthNum);
            if (p) {
                if (p.startDate.getTime() !== p.endDate.getTime()) {
                    isMultiDay = true;
                    fileDate = p.startDate;
                } else {
                    fileDate = p.startDate;
                }
                break;
            }
        }
    }

    if (!fileDate) {
        fileDate = parseDateFromFileName(fileName, activeMonthNum);
    }

    if (!fileDate) {
        return { error: "Could not determine date from Period line or file name (" + fileName + ")" };
    }

    if (isMultiDay) {
        return {
            date: fileDate,
            isMultiDay: true,
            periodStr: periodStr,
            error: "Multi-day extract skipped (" + periodStr + ")"
        };
    }

    // 2. Search for header row
    let headerRowIdx = -1;
    const colMap = {};

    for (let r = 0; r < Math.min(grid.length, 15); r++) {
        const row = grid[r];
        const rowLower = row.map(cell => (cell || "").toString().toLowerCase().trim());
        const hasName = rowLower.some(c => c === "name" || c === "driver name" || (c === "driver" && !c.includes("id")));
        const hasPayOrTrips = rowLower.some(c => c.includes("trip") || c.includes("pay") || c.includes("fee") || c.includes("driver id"));

        if (hasName && hasPayOrTrips) {
            headerRowIdx = r;
            rowLower.forEach((h, c) => {
                if (h === "name" || h.includes("driver name") || (h === "driver" && !h.includes("id"))) colMap.name = c;
                else if (h.includes("driver id") || h === "id") colMap.driverId = c;
                else if (h === "trips" || h === "trip" || h.includes("trips")) colMap.trips = c;
                else if (h.includes("driver net") || h === "net") colMap.driverNet = c;
                else if (h.includes("fee")) colMap.fees = c;
                else if (h.includes("toll")) colMap.tolls = c;
                else if (h.includes("final pay") || h === "pay") colMap.finalPay = c;
                else if (h.includes("cash")) colMap.cash = c;
            });
            break;
        }
    }

    if (headerRowIdx === -1 || colMap.name === undefined) {
        return { error: "Could not find valid table headers in " + fileName };
    }

    // 3. Read driver rows
    const records = [];
    let emptyCount = 0;

    for (let r = headerRowIdx + 1; r < grid.length; r++) {
        const row = grid[r];
        const nameVal = (row[colMap.name] || "").toString().trim();
        const idVal = colMap.driverId !== undefined ? (row[colMap.driverId] || "").toString().trim() : "";

        if (nameVal.toLowerCase() === "total" || nameVal.toLowerCase() === "totals" || idVal.toLowerCase() === "total") {
            break;
        }

        if (!nameVal) {
            emptyCount++;
            if (emptyCount > 5) break;
            continue;
        }
        emptyCount = 0;

        const trips = colMap.trips !== undefined ? (parseNumber(row[colMap.trips]) || 0) : 0;
        const fees = colMap.fees !== undefined ? (parseNumber(row[colMap.fees]) || 0) : 0;
        const tolls = colMap.tolls !== undefined ? (parseNumber(row[colMap.tolls]) || 0) : 0;
        const finalPay = colMap.finalPay !== undefined ? (parseNumber(row[colMap.finalPay]) || 0) : 0;
        const cash = colMap.cash !== undefined ? (parseNumber(row[colMap.cash]) || 0) : 0;

        // Total credit = cash + final pay + amount of fees (positive value)
        const credit = roundToTwo(cash + finalPay + Math.abs(fees));

        records.push({
            name: nameVal,
            driverId: idVal,
            trips: trips,
            fees: fees,
            tolls: tolls,
            finalPay: finalPay,
            cash: cash,
            credit: credit
        });
    }

    return {
        date: fileDate,
        dateKey: new Date(fileDate.getFullYear(), fileDate.getMonth(), fileDate.getDate()).getTime(),
        isMultiDay: false,
        periodStr: periodStr,
        records: records
    };
}

/**
 * Reads all daily .xlsx files from the Drive folder.
 * Uses ScriptProperties cache so unchanged files are not re-unzipped on each run.
 * Deduplicates multiple files for the same date (picks latest modified).
 * Filters strictly to the active month (October).
 */
function loadDailyExtracts(ss) {
    ss = getSafeSpreadsheet(ss);
    const monthCtx = getActiveMonthContext(ss);
    const targetMonthNum = monthCtx ? monthCtx.monthNum : _ACTIVE_MONTH_NUM;
    const targetYear = (monthCtx && monthCtx.year > 2000) ? monthCtx.year : 2026;

    let folder;
    try {
        folder = DriveApp.getFolderById(DAILY_EXTRACTS_FOLDER_ID);
    } catch (e) {
        throw new Error("Could not access Drive folder (" + DAILY_EXTRACTS_FOLDER_ID + "): " + e.message);
    }

    const files = folder.getFiles();
    const candidateFiles = [];

    while (files.hasNext()) {
        const file = files.next();
        const fName = file.getName();
        const lower = fName.toLowerCase();
        if (lower.endsWith(".xlsx") || lower.endsWith(".xls") || lower.includes("payables") || lower.includes("driver")) {
            candidateFiles.push({
                file: file,
                id: file.getId(),
                name: fName,
                lastUpdated: file.getLastUpdated().getTime(),
                lastUpdatedDate: file.getLastUpdated()
            });
        }
    }

    if (candidateFiles.length === 0) {
        Logger.log("⚠️ No extract files found in Drive folder: " + DAILY_EXTRACTS_FOLDER_ID);
        return { records: [], dates: [], allDateCols: [], warnings: ["No files found in Drive folder."], logEntries: [] };
    }

    const props = PropertiesService.getScriptProperties();
    const parsedFiles = [];
    const logEntries = [];
    const warnings = [];

    for (let i = 0; i < candidateFiles.length; i++) {
        const cf = candidateFiles[i];
        let parseResult = null;

        const cacheKey = "EXTRACT_CACHE_" + cf.id;
        try {
            const cachedStr = props.getProperty(cacheKey);
            if (cachedStr) {
                const cachedObj = JSON.parse(cachedStr);
                if (cachedObj && cachedObj.lastUpdated === cf.lastUpdated) {
                    cachedObj.date = new Date(cachedObj.dateKey);
                    parseResult = cachedObj;
                }
            }
        } catch (cErr) {}

        if (!parseResult) {
            try {
                const blob = cf.file.getBlob();
                const grid = parseXlsxBlob(blob);
                parseResult = extractDailyRecordsFromGrid(grid, cf.name, targetMonthNum);

                if (!parseResult.error && !parseResult.isMultiDay && parseResult.records) {
                    try {
                        const toCache = {
                            lastUpdated: cf.lastUpdated,
                            dateKey: parseResult.dateKey,
                            isMultiDay: parseResult.isMultiDay,
                            records: parseResult.records
                        };
                        props.setProperty(cacheKey, JSON.stringify(toCache));
                    } catch (storeErr) {}
                }
            } catch (pErr) {
                parseResult = { error: pErr.message };
            }
        }

        if (parseResult.error) {
            logEntries.push({
                name: cf.name,
                date: parseResult.date || null,
                driversCount: 0,
                totalCredit: 0,
                totalTolls: 0,
                status: "❌ " + parseResult.error,
                lastUpdated: cf.lastUpdatedDate
            });
            warnings.push(cf.name + ": " + parseResult.error);
            continue;
        }

        if (parseResult.isMultiDay) {
            logEntries.push({
                name: cf.name,
                date: parseResult.date,
                driversCount: 0,
                totalCredit: 0,
                totalTolls: 0,
                status: "⚠️ Skipped: Multi-day period (" + parseResult.periodStr + ")",
                lastUpdated: cf.lastUpdatedDate
            });
            warnings.push("Skipped multi-day file: " + cf.name);
            continue;
        }

        const fileDate = parseResult.date;
        if (fileDate.getMonth() + 1 !== targetMonthNum || fileDate.getFullYear() !== targetYear) {
            logEntries.push({
                name: cf.name,
                date: fileDate,
                driversCount: (parseResult.records || []).length,
                totalCredit: 0,
                totalTolls: 0,
                status: "⚠️ Skipped: Date " + formatDate(fileDate) + " is not in active month (" + targetMonthNum + "/" + targetYear + ")",
                lastUpdated: cf.lastUpdatedDate
            });
            continue;
        }

        const totCred = (parseResult.records || []).reduce((sum, r) => sum + (r.credit || 0), 0);
        const totTolls = (parseResult.records || []).reduce((sum, r) => sum + (r.tolls || 0), 0);

        parsedFiles.push({
            fileInfo: cf,
            date: fileDate,
            dateKey: parseResult.dateKey,
            records: parseResult.records,
            totalCredit: roundToTwo(totCred),
            totalTolls: roundToTwo(totTolls)
        });
    }

    // Handle duplicates for same date: keep latest file
    const byDate = {};
    parsedFiles.forEach(pf => {
        if (!byDate[pf.dateKey]) {
            byDate[pf.dateKey] = [];
        }
        byDate[pf.dateKey].push(pf);
    });

    const activeFiles = [];
    for (const dk in byDate) {
        const list = byDate[dk];
        list.sort((a, b) => b.fileInfo.lastUpdated - a.fileInfo.lastUpdated);
        const winner = list[0];
        activeFiles.push(winner);

        logEntries.push({
            name: winner.fileInfo.name,
            date: winner.date,
            driversCount: winner.records.length,
            totalCredit: winner.totalCredit,
            totalTolls: winner.totalTolls,
            status: "✅ Active",
            lastUpdated: winner.fileInfo.lastUpdatedDate
        });

        for (let j = 1; j < list.length; j++) {
            const dup = list[j];
            logEntries.push({
                name: dup.fileInfo.name,
                date: dup.date,
                driversCount: dup.records.length,
                totalCredit: dup.totalCredit,
                totalTolls: dup.totalTolls,
                status: "⚠️ Skipped: Duplicate date (newer file " + winner.fileInfo.name + " used)",
                lastUpdated: dup.fileInfo.lastUpdatedDate
            });
            warnings.push("Duplicate date for " + dup.fileInfo.name + "; used newer " + winner.fileInfo.name);
        }
    }

    activeFiles.sort((a, b) => a.dateKey - b.dateKey);

    const allRecords = [];
    const dateMap = {};
    activeFiles.forEach(af => {
        dateMap[af.dateKey] = {
            date: af.date,
            headerStr: formatDate(af.date)
        };
        af.records.forEach(r => {
            allRecords.push({
                date: af.date,
                dateKey: af.dateKey,
                name: r.name,
                driverId: r.driverId,
                trips: r.trips,
                fees: r.fees,
                tolls: r.tolls,
                finalPay: r.finalPay,
                cash: r.cash,
                credit: r.credit
            });
        });
    });

    const sortedDates = Object.keys(dateMap).map(Number).sort((a, b) => a - b).map(t => dateMap[t]);

    try {
        updateDailyFilesLogSheet(ss, logEntries);
    } catch (lErr) {
        Logger.log("Error updating Daily Files Log sheet: " + lErr.message);
    }

    return {
        records: allRecords,
        dates: sortedDates.map(d => d.date),
        allDateCols: sortedDates,
        warnings: warnings,
        logEntries: logEntries
    };
}

/**
 * Creates/refreshes the "Daily Files Log" sheet showing all files parsed from Drive.
 */
function updateDailyFilesLogSheet(ss, logEntries) {
    if (!ss) return;
    let logSheet = ss.getSheetByName("Daily Files Log");
    if (!logSheet) {
        logSheet = ss.insertSheet("Daily Files Log");
    }

    const headers = ["File Name", "File Date", "Drivers", "Total Credit", "Total Tolls", "Status", "Last Modified in Drive", "Processed At"];
    logSheet.clear();
    logSheet.getRange(1, 1, 1, headers.length)
        .setValues([headers])
        .setFontWeight("bold")
        .setBackground("#3c78d8")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    if (logEntries && logEntries.length > 0) {
        logEntries.sort((a, b) => {
            if (a.date && b.date) return b.date.getTime() - a.date.getTime();
            return a.name.localeCompare(b.name);
        });

        const rows = logEntries.map(e => [
            e.name,
            e.date ? formatDate(e.date) : "N/A",
            e.driversCount,
            e.totalCredit || 0,
            e.totalTolls || 0,
            e.status,
            e.lastUpdated ? e.lastUpdated.toLocaleString() : "N/A",
            new Date().toLocaleString()
        ]);

        const range = logSheet.getRange(2, 1, rows.length, headers.length);
        range.setValues(rows);
        range.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
        logSheet.getRange(2, 4, rows.length, 2).setNumberFormat("$#,##0.00");
        logSheet.getRange(2, 1, rows.length, 1).setFontWeight("bold");

        for (let i = 0; i < rows.length; i++) {
            const st = rows[i][5];
            const cell = logSheet.getRange(2 + i, 6);
            if (st.startsWith("✅")) cell.setBackground("#d9ead3").setFontColor("#274e13");
            else if (st.startsWith("⚠️")) cell.setBackground("#fff2cc").setFontColor("#7f6000");
            else if (st.startsWith("❌")) cell.setBackground("#f4cccc").setFontColor("#990000");
        }
    }

    logSheet.autoResizeColumns(1, headers.length);
}

/**
 * Helper: scans a header row array for the first column whose label matches
 * one of the given keywords (case-insensitive). Returns that column's 0-based
 * index. Falls back to `defaultIdx` if nothing matches.
 */
function findColByLabel(headers, keywords, defaultIdx) {
    for (let c = headers.length - 1; c >= 0; c--) {
        const h = (headers[c] || '').toString().toLowerCase().trim();
        for (const kw of keywords) {
            if (h.includes(kw)) return c;
        }
    }
    return defaultIdx;
}

/**
 * Scans the main Summary table headers (from left to right, within columns 0-11)
 * looking for a column matching one of the given keywords.
 * Never scans past column index 11 to avoid matching adjacent tables (stats/daily).
 */
function findMainSummaryCol(headers, keywords, defaultIdx) {
    if (!headers || !headers.length) return defaultIdx;
    const maxCol = Math.min(headers.length, 12);
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
 *  - Header row: Column A is "Driver", followed by date columns (YYYY-MM-DD or DD-MM-YY), then Total, Count, Driver NET
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
    const ss = (sheet && typeof sheet.getParent === "function") ? sheet.getParent() : null;
    const monthCtx = getActiveMonthContext(ss || sheet);
    const targetMonthNum = monthCtx ? monthCtx.monthNum : _ACTIVE_MONTH_NUM;

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
                const d = parseSheetDateHeader(headerRow[c], data[headerRowIdx][c], targetMonthNum);
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
 * Ensures clean trigger state.
 * Under Formula-First architecture, native equations update dynamically on edits,
 * so intrusive onEdit/onChange script triggers that wipe the sheet on slight edits are removed.
 */
function setupAutoRun() {
    let ss = getSafeSpreadsheet();
    if (!ss) {
        Logger.log("❌ Target spreadsheet not found.");
        return;
    }

    // Remove existing triggers for runDailyBalance to prevent full-sheet rebuilds on minor edits
    let removedCount = 0;
    ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getHandlerFunction() === 'runDailyBalance') {
            ScriptApp.deleteTrigger(t);
            removedCount++;
        }
    });

    Logger.log("✅ Removed " + removedCount + " auto-run triggers.");
    Logger.log("   • The sheet now uses live formulas that update automatically on edits without wiping.");
    Logger.log("   • To refresh daily extracts from Google Drive, click '⚙️ Balance Sheet -> ▶ Run Balance Report'.");
    try {
        ss.toast("Formula-First active: live formulas update on edits. Use menu to pull new Drive files.", "Balance Sheet", 6);
    } catch (e) { }
}


// =================================================================
// WORKING HOURS SHEET GENERATION & LOAD HELPERS
// =================================================================

function runSyncBreaksTable() {
    Logger.log("=== Starting runSyncBreaksTable ===");
    let ss = getSafeSpreadsheet();
    if (!ss) throw new Error("Could not access spreadsheet.");
    Logger.log("Target Spreadsheet: " + ss.getName() + " (" + ss.getId() + ")");
    const info = syncBreaksTable(ss);
    SpreadsheetApp.flush();
    Logger.log("=== Successfully synced breaks table at row " + info.titleRow + " ===");
    return info;
}

function generateWorkingHoursSheetFromMenu() {
    let ss = getSafeSpreadsheet();
    if (!ss) return;
    generateWorkingHoursSheet(ss);
    syncBreaksTable(ss);
    SpreadsheetApp.flush();
    try {
        ss.toast("Working hours and breaks table synced!", "Balance Sheet", 3);
    } catch (e) { }
}

function syncBreaksTableFromMenu() {
    let ss = getSafeSpreadsheet();
    if (!ss) return;
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
 * Supports YYYY-MM-DD (e.g. 2026-10-05) and DD-MM-YY display strings to prevent Google Sheets US-locale Date object distortion.
 */
function parseSheetDateHeader(displayHeader, rawHeader, targetMonthNum) {
    const activeM = (typeof targetMonthNum === 'number' && targetMonthNum >= 1 && targetMonthNum <= 12)
        ? targetMonthNum
        : _ACTIVE_MONTH_NUM;

    const disp = displayHeader ? String(displayHeader).trim() : "";

    // Ignore known non-date headers
    if (disp) {
        const lower = disp.toLowerCase();
        if (lower === "driver" || lower === "total" || lower === "count" || lower === "net" ||
            lower === "trip" || lower === "trips" || lower === "hours" || lower === "cash" ||
            lower.includes("driver net") || lower.includes("total hours")) {
            return null;
        }
    }

    // 1. Try parsing display value first with active month context
    if (disp) {
        const dDisp = parseDate(disp, activeM);
        if (dDisp && !isNaN(dDisp.getTime())) {
            // If the parsed date matches activeM, it's definitely the intended date!
            if (activeM && (dDisp.getMonth() + 1 === activeM)) {
                return dDisp;
            }
            // Even if not strictly activeM (e.g. cross-month week), if rawHeader is not a Date or is invalid, use dDisp
            if (!(rawHeader instanceof Date)) {
                return dDisp;
            }
        }
    }

    // 2. Check rawHeader if it's already a native Date object
    if (rawHeader instanceof Date && !isNaN(rawHeader.getTime())) {
        const rawDate = new Date(rawHeader.getFullYear(), rawHeader.getMonth(), rawHeader.getDate());
        // If rawHeader's month matches activeM, trust it
        if (!activeM || (rawDate.getMonth() + 1 === activeM)) {
            return rawDate;
        }
        // If rawHeader's month does NOT match activeM (e.g. Google Sheets locale parsed 10-05 as May 10 instead of Oct 5):
        // If we had a parsed date from display header, prefer display header!
        if (disp) {
            const dDisp = parseDate(disp, activeM);
            if (dDisp && !isNaN(dDisp.getTime())) {
                return dDisp;
            }
        }
        return rawDate;
    }

    // 3. If rawHeader is a string/number
    if (rawHeader) {
        const dRaw = parseDate(rawHeader, activeM);
        if (dRaw && !isNaN(dRaw.getTime())) return dRaw;
    }

    // 4. Fallback to display header parsed date if any
    if (disp) {
        const dDisp = parseDate(disp, activeM);
        if (dDisp && !isNaN(dDisp.getTime())) return dDisp;
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
    const monthCtx = getActiveMonthContext(ss);
    const targetMonthNum = monthCtx ? monthCtx.monthNum : _ACTIVE_MONTH_NUM;

    const data = hoursSheet.getDataRange().getValues();
    if (data.length < 2) return { map: {}, driverTotalHours: {}, dailyTotalHours: {} };

    const lastCol = hoursSheet.getLastColumn();
    const rawHeaders = hoursSheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const displayHeaders = hoursSheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];

    const dateCols = {}; // colIdx -> timestamp
    for (let c = 1; c < lastCol; c++) {
        const hText = (displayHeaders[c] || '').toString().toLowerCase().trim();
        if (hText.includes("total")) continue;

        const d = parseSheetDateHeader(displayHeaders[c], rawHeaders[c], targetMonthNum);
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
function generateWorkingHoursSheet(ss, dailyData) {
    let hoursSheet = ss.getSheetByName("Working hours");
    if (!hoursSheet) hoursSheet = ss.insertSheet("Working hours");

    if (!dailyData) {
        dailyData = loadDailyExtracts(ss);
    }
    if (!dailyData || !dailyData.records || dailyData.records.length === 0) {
        Logger.log('⚠️ No daily extract records found for Working Hours sheet.');
        return;
    }

    const allDateCols = dailyData.allDateCols || [];
    if (allDateCols.length === 0) return;

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

    // Collect all unique drivers and their credit across daily records
    const driverList = [];
    const driverCreditMap = {}; // canonDriver -> dateTimestamp -> credit

    dailyData.records.forEach(r => {
        const canon = getCanonicalDriverName(r.name, canonicalDrivers);
        if (!canonicalDrivers.includes(canon)) canonicalDrivers.push(canon);
        if (!driverList.includes(canon)) driverList.push(canon);

        if (!driverCreditMap[canon]) driverCreditMap[canon] = {};
        driverCreditMap[canon][r.dateKey] = (driverCreditMap[canon][r.dateKey] || 0) + (r.credit || 0);
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
    ss = getSafeSpreadsheet(ss);
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

        const monthCtx = getActiveMonthContext(ss);
        const targetMonthNum = monthCtx ? monthCtx.monthNum : _ACTIVE_MONTH_NUM;
        const rawDate = rawData[r][dateColIdx];
        const dispDate = dispData[r][dateColIdx];
        const d = parseSheetDateHeader(dispDate, rawDate, targetMonthNum) || parseDate(dispDate, targetMonthNum) || parseDate(rawDate, targetMonthNum);
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
    ss = getSafeSpreadsheet(ss);
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

    const monthCtx = getActiveMonthContext(ss);
    const targetMonthNum = monthCtx ? monthCtx.monthNum : _ACTIVE_MONTH_NUM;

    // 3. Read date columns directly from Row 1
    const row1Raw = hoursSheet.getRange(1, 1, 1, totalColNumber).getValues()[0];
    const allDateCols = [];
    for (let c = 2; c < totalColNumber; c++) {
        const disp = (row1Data[c - 1] || '').toString().trim();
        const raw = row1Raw[c - 1];
        const d = parseSheetDateHeader(disp, raw, targetMonthNum) || parseDate(disp, targetMonthNum);
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
    let ss = getSafeSpreadsheet();
    if (!ss) return;

    try {
        ss.toast("🔄 Reading daily extracts from Google Drive...", "Balance Sheet", 5);
    } catch (e) { }

    // Use LockService to prevent overlapping runs when multiple edits occur
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(15000)) {
        try {
            ss.toast("⚠️ Another report run is in progress. Please wait a moment.", "Balance Sheet", 5);
        } catch (e) { }
        Logger.log("⚠️ Another report run is already in progress. Skipping.");
        return;
    }

    try {
        // Load all daily extracts from Google Drive folder
        let dailyData = null;
        try {
            dailyData = loadDailyExtracts(ss);
        } catch (e) {
            Logger.log("Error loading daily extracts: " + e.toString() + "\n" + (e.stack || ""));
        }

        // 0. Sync working hours sheet first so downstream reports can use working hours data
        try {
            generateWorkingHoursSheet(ss, dailyData);
        } catch (e) {
            Logger.log("Error in generateWorkingHoursSheet: " + e.toString() + "\n" + (e.stack || ""));
        }

        // 0b. Sync breaks table directly
        try {
            syncBreaksTable(ss);
        } catch (e) {
            Logger.log("Error in syncBreaksTable: " + e.toString() + "\n" + (e.stack || ""));
        }

        // 1. Update the main Summary sheet from daily extracts.
        try {
            updateSummaryAndCharts(ss, dailyData);
        } catch (e) {
            Logger.log("Error in updateSummaryAndCharts: " + e.toString() + "\n" + (e.stack || ""));
        }

        // 2. Wait for all pending spreadsheet changes to apply.
        SpreadsheetApp.flush();

        // 3. Generate the weekly summary.
        try {
            generateWeeklySummary(ss, dailyData);
        } catch (e) {
            Logger.log("Error in generateWeeklySummary: " + e.toString() + "\n" + (e.stack || ""));
        }

        // 4. Generate the final Bonus report.
        try {
            generateBonusReport(ss, dailyData);
        } catch (e) {
            Logger.log("Error in generateBonusReport: " + e.toString() + "\n" + (e.stack || ""));
        }

        // 5. Generate Angel's specific progressive report.
        try {
            generateAngelReport(ss, dailyData);
        } catch (e) {
            Logger.log("Error in generateAngelReport: " + e.toString() + "\n" + (e.stack || ""));
        }

        try {
            ss.toast("✅ Balance report complete! Tolls and formulas updated.", "Balance Sheet", 10);
        } catch (e) { }
    } finally {
        lock.releaseLock();
    }
}


// =================================================================
// SUMMARY BUILD — Builds Summary sheet from Daily Extracts (Drive)
// =================================================================
function updateSummaryAndCharts(ss, dailyData) {
    let summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) summarySheet = ss.insertSheet("Summary");

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

    if (!dailyData) {
        dailyData = loadDailyExtracts(ss);
    }
    if (!dailyData || !dailyData.records || dailyData.records.length === 0) {
        Logger.log("⚠️ No daily extract records found to build Summary.");
        try {
            ss.toast("⚠️ No daily extract files found for this month in Google Drive folder.", "Balance Sheet", 10);
        } catch (e) { }
        return;
    }

    // --- Helper to load exact daily data from auxiliary sheets ---
    function loadAuxMap(sheetName) {
        const auxSheet = ss.getSheetByName(sheetName);
        if (!auxSheet) return {};
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

    // Load auxiliary daily maps
    const tripsDataMap = loadAuxMap("Raw Data - Trips");
    const noShowMap = loadNoShowMap(ss, "No Show");
    const { map: workingHoursMap } = loadWorkingHoursMap(ss);

    // Group and consolidate records from dailyData
    const consolidatedMap = {};
    dailyData.records.forEach(r => {
        const canonDriver = getCanonicalDriverName(r.name, canonicalDrivers);
        if (!canonicalDrivers.includes(canonDriver)) canonicalDrivers.push(canonDriver);

        const key = `${r.dateKey}_${canonDriver}`;
        if (!consolidatedMap[key]) {
            consolidatedMap[key] = {
                date: r.date,
                dateKey: r.dateKey,
                driver: canonDriver,
                credit: 0,
                cash: 0,
                tolls: 0,
                trips: 0
            };
        }
        consolidatedMap[key].credit += (r.credit || 0);
        consolidatedMap[key].cash += (r.cash || 0);
        consolidatedMap[key].tolls += (r.tolls || 0);
        consolidatedMap[key].trips += (r.trips || 0);
    });

    const outputRows = [];
    for (const key in consolidatedMap) {
        const item = consolidatedMap[key];
        const driverName = item.driver;
        const dateKey = item.dateKey;

        let credit = roundToTwo(item.credit);
        const hasNoShow = noShowMap[driverName] && noShowMap[driverName][dateKey] !== undefined;
        let noShowCount = 0;
        if (hasNoShow) {
            credit = roundToTwo(credit - noShowMap[driverName][dateKey].cash);
            noShowCount = noShowMap[driverName][dateKey].count;
        }

        const exactTrips = (tripsDataMap[driverName] && tripsDataMap[driverName][dateKey] !== undefined)
            ? tripsDataMap[driverName][dateKey]
            : item.trips;

        const dailyHours = (workingHoursMap[driverName] && workingHoursMap[driverName][dateKey])
            ? workingHoursMap[driverName][dateKey]
            : 0;

        outputRows.push({
            date: item.date,
            dateKey: item.dateKey,
            driver: driverName,
            credit: credit,
            cash: roundToTwo(item.cash),
            tolls: roundToTwo(item.tolls),
            trips: exactTrips,
            hours: dailyHours,
            noShowCount: noShowCount
        });
    }

    if (outputRows.length === 0) {
        Logger.log("⚠️ No driver records found to display in Summary.");
        return;
    }

    // SORT BY DATE CHRONOLOGICALLY, then by driver name
    outputRows.sort((a, b) => {
        const dateDiff = a.date.getTime() - b.date.getTime();
        if (dateDiff !== 0) return dateDiff;
        return a.driver.localeCompare(b.driver);
    });

    // --- Headers: 11 Columns (with Tolls in Col I = col 9) ---
    const headers = ["Date", "Driver", "Credit", "Hours", "Driver Avg/ Hour", "CPH", "Trips", "Cash", "Tolls", "No Show", "Balance"];
    summarySheet.getRange(2, 1, 1, headers.length).setValues([headers])
        .setFontWeight("bold")
        .setBackground("#3c78d8")
        .setFontColor("white")
        .setHorizontalAlignment("center");

    // Track progressive tiering for drivers on "80%-90%" (e.g. Angel)
    const progressiveTrackers = {}; // driver -> { weekStart, weeklyFeeBase }

    const rowsToWrite = [];
    let previousDateString = "";
    let dayBlocks = [];
    let currentStartRow = 3;
    let currentGroupSize = 0;

    outputRows.forEach((row, index) => {
        const sheetRow = 3 + index;
        const currentDateString = row.date ? row.date.toDateString() : "";
        let displayDate = row.date;

        if (currentDateString === previousDateString) {
            displayDate = "";
            currentGroupSize++;
        } else {
            if (currentGroupSize > 0) {
                dayBlocks.push({ row: currentStartRow, count: currentGroupSize, date: outputRows[index - 1].date });
            }
            previousDateString = currentDateString;
            currentStartRow = sheetRow;
            currentGroupSize = 1;
        }

        const fare = existingFares[row.driver];
        const isProgressive = String(fare).trim() === "80%-90%";

        let balanceValOrFormula = "";
        if (isProgressive) {
            let weekStart = getMonday(row.date).getTime();
            if (!progressiveTrackers[row.driver]) {
                progressiveTrackers[row.driver] = { weekStart: 0, weeklyFeeBase: 0 };
            }
            if (progressiveTrackers[row.driver].weekStart !== weekStart) {
                progressiveTrackers[row.driver] = { weekStart: weekStart, weeklyFeeBase: 0 };
            }

            const feeBase = Math.max(0, row.credit - row.tolls);
            let todayGrossPayout = 0;
            if (feeBase > 0) {
                const currentWeekFeeBase = progressiveTrackers[row.driver].weeklyFeeBase;
                if (currentWeekFeeBase >= 1000) {
                    todayGrossPayout = feeBase * 0.90;
                } else if ((currentWeekFeeBase + feeBase) <= 1000) {
                    todayGrossPayout = feeBase * 0.80;
                } else {
                    const basePortion = 1000 - currentWeekFeeBase;
                    const topPortion = feeBase - basePortion;
                    todayGrossPayout = (basePortion * 0.80) + (topPortion * 0.90);
                }
            }
            progressiveTrackers[row.driver].weeklyFeeBase += feeBase;
            balanceValOrFormula = roundToTwo(todayGrossPayout + row.tolls - row.cash);
        } else {
            // Formula-first: Standard fare balance formula referencing Settings sheet and Tolls
            // Balance = (Credit - Tolls) * Fare% + Tolls - Cash
            balanceValOrFormula = `=ROUND((C${sheetRow}-I${sheetRow})*IFERROR(VLOOKUP(B${sheetRow}, Settings!$A:$B, 2, FALSE), 0.9) + I${sheetRow} - H${sheetRow}, 2)`;
        }

        const formulaDEPH = `=IF(D${sheetRow}>0, ROUND((H${sheetRow}+K${sheetRow})/D${sheetRow}, 2), 0)`;
        const formulaCPH = `=IF(D${sheetRow}>0, ROUND(C${sheetRow}/D${sheetRow}, 2), 0)`;

        rowsToWrite.push([
            displayDate,
            row.driver,
            row.credit,
            roundToTwo(row.hours) || 0,
            formulaDEPH,
            formulaCPH,
            row.trips,
            row.cash,
            row.tolls,
            row.noShowCount,
            balanceValOrFormula
        ]);
    });

    if (currentGroupSize > 0) {
        dayBlocks.push({ row: currentStartRow, count: currentGroupSize, date: outputRows[outputRows.length - 1].date });
    }

    // --- Clear old data and write new 11-column table ---
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

    // Apply number formats
    summarySheet.getRange(3, 1, rowsToWrite.length, 1).setNumberFormat("dddd, dd");
    summarySheet.getRange(3, 3, rowsToWrite.length, 1).setNumberFormat("$#,##0.00"); // Credit
    summarySheet.getRange(3, 4, rowsToWrite.length, 1).setNumberFormat("0.0");       // Hours
    summarySheet.getRange(3, 5, rowsToWrite.length, 2).setNumberFormat("$#,##0.00"); // DEPH, CPH
    summarySheet.getRange(3, 7, rowsToWrite.length, 1).setNumberFormat("0");         // Trips
    summarySheet.getRange(3, 8, rowsToWrite.length, 2).setNumberFormat("$#,##0.00"); // Cash, Tolls
    summarySheet.getRange(3, 10, rowsToWrite.length, 1).setNumberFormat("0");        // No Show
    summarySheet.getRange(3, 11, rowsToWrite.length, 1).setNumberFormat("$#,##0.00"); // Balance

    // Apply day block borders and merge date cells
    dayBlocks.forEach(block => {
        summarySheet.getRange(block.row, 1, block.count, headers.length)
            .setBorder(true, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        if (block.count > 1) {
            summarySheet.getRange(block.row, 1, block.count, 1).mergeVertically();
        }
    });

    // Auto-sync Settings sheet (keep existing fare %, default 90% for new drivers)
    const uniqueDrivers = Array.from(new Set(rowsToWrite.map(r => r[1]).filter(Boolean))).sort();
    const newSettingsRows = [["Driver", "Fare %"]];
    uniqueDrivers.forEach(d => {
        let fare = existingFares[d];
        if (fare === undefined || fare === null || fare === "") fare = "90%";
        newSettingsRows.push([d, fare]);
    });
    for (let oldDriver in existingFares) {
        if (!uniqueDrivers.includes(oldDriver)) {
            newSettingsRows.push([oldDriver, existingFares[oldDriver]]);
        }
    }
    settingsSheet.clear();
    settingsSheet.getRange(1, 1, newSettingsRows.length, 2).setValues(newSettingsRows);
    settingsSheet.getRange("A1:B1").setFontWeight("bold").setBackground("#d0e0e3");
    settingsSheet.autoResizeColumns(1, 2);

    // =================================================================
    // DRIVER STATS TABLE (Formula-First, Col M = 13 to Col V = 22)
    // =================================================================
    const statsStartCol = 13; // Column M (Cols 1-11 = A-K, Col 12 = L gap)
    const statsHeaders = ["Driver", "Total Credit", "Total Tolls", "Total Trips", "Total Cash", "No Show", "Total Hours", "Avg Per Hour", "Driver NET", "Avg Per Day"];
    const statsRows = [statsHeaders];

    uniqueDrivers.forEach((driver, idx) => {
        const r = 4 + idx; // Data row starts at 4 (header at 3)
        statsRows.push([
            driver,
            `=SUMIF($B$3:$B, M${r}, $C$3:$C)`,
            `=SUMIF($B$3:$B, M${r}, $I$3:$I)`,
            `=SUMIF($B$3:$B, M${r}, $G$3:$G)`,
            `=SUMIF($B$3:$B, M${r}, $H$3:$H)`,
            `=SUMIF($B$3:$B, M${r}, $J$3:$J)`,
            `=SUMIF($B$3:$B, M${r}, $D$3:$D)`,
            `=IF(S${r}>0, ROUND(N${r}/S${r}, 2), 0)`,
            `=ROUND(N${r}-Q${r}, 2)`,
            `=IF(COUNTIFS($B$3:$B, M${r}, $C$3:$C, ">0")>0, ROUND(N${r}/COUNTIFS($B$3:$B, M${r}, $C$3:$C, ">0"), 2), 0)`
        ]);
    });

    try { summarySheet.getRange(3, statsStartCol, 100, statsHeaders.length).clearContent(); } catch (e) { }
    const statsRange = summarySheet.getRange(3, statsStartCol, statsRows.length, statsHeaders.length);
    statsRange.setValues(statsRows);
    statsRange.setFontSize(7);
    summarySheet.getRange(3, statsStartCol, 1, statsHeaders.length).setFontWeight("bold");

    // Format stats table columns
    summarySheet.getRange(4, statsStartCol + 1, uniqueDrivers.length, 2).setNumberFormat("$#,##0.00"); // Credit, Tolls
    summarySheet.getRange(4, statsStartCol + 3, uniqueDrivers.length, 1).setNumberFormat("0");         // Trips
    summarySheet.getRange(4, statsStartCol + 4, uniqueDrivers.length, 1).setNumberFormat("$#,##0.00"); // Cash
    summarySheet.getRange(4, statsStartCol + 5, uniqueDrivers.length, 1).setNumberFormat("0");         // No Show
    summarySheet.getRange(4, statsStartCol + 6, uniqueDrivers.length, 1).setNumberFormat("0.0");       // Total Hours
    summarySheet.getRange(4, statsStartCol + 7, uniqueDrivers.length, 3).setNumberFormat("$#,##0.00"); // Avg Per Hour, NET, Avg Per Day

    summarySheet.setColumnWidth(statsStartCol, 75); // Col M (Driver)
    for (let c = statsStartCol + 1; c < statsStartCol + statsHeaders.length; c++) {
        summarySheet.setColumnWidth(c, 55);
    }

    // =================================================================
    // DAILY SUMMARY TABLE (Formula-First, Col X = 24 to Col AD = 30)
    // =================================================================
    const dailyStartCol = 24; // Column X (Col 23 = W gap)
    const dailyHeaders = ["Date", "Total Credit", "Tolls", "Trips", "Cash", "No Show", "Total Hours"];
    const dailyTableRows = [];

    dayBlocks.forEach(block => {
        const startR = block.row;
        const endR = block.row + block.count - 1;
        dailyTableRows.push([
            block.date,
            `=SUM(C${startR}:C${endR})`,
            `=SUM(I${startR}:I${endR})`,
            `=SUM(G${startR}:G${endR})`,
            `=SUM(H${startR}:H${endR})`,
            `=SUM(J${startR}:J${endR})`,
            `=SUM(D${startR}:D${endR})`
        ]);
    });

    try { summarySheet.getRange(2, dailyStartCol, 100, dailyHeaders.length).clearContent(); } catch (e) { }
    summarySheet.getRange(2, dailyStartCol, 1, dailyHeaders.length).setValues([dailyHeaders]).setFontWeight("bold");

    if (dailyTableRows.length > 0) {
        summarySheet.getRange(3, dailyStartCol, dailyTableRows.length, dailyHeaders.length).setValues(dailyTableRows);
        summarySheet.getRange(3, dailyStartCol, dailyTableRows.length, 1).setNumberFormat("dddd, dd");
        summarySheet.getRange(3, dailyStartCol + 1, dailyTableRows.length, 2).setNumberFormat("$#,##0.00"); // Credit, Tolls
        summarySheet.getRange(3, dailyStartCol + 3, dailyTableRows.length, 1).setNumberFormat("0");         // Trips
        summarySheet.getRange(3, dailyStartCol + 4, dailyTableRows.length, 1).setNumberFormat("$#,##0.00"); // Cash
        summarySheet.getRange(3, dailyStartCol + 5, dailyTableRows.length, 1).setNumberFormat("0");         // No Show
        summarySheet.getRange(3, dailyStartCol + 6, dailyTableRows.length, 1).setNumberFormat("0.0");       // Hours
    }

    // =================================================================
    // CHARTS
    // =================================================================
    const allCharts = summarySheet.getCharts();
    allCharts.forEach(c => summarySheet.removeChart(c));

    // Chart 1: Daily Total Credit Trend (line) — Col 32 (AF, gap at Col 31 = AE)
    if (dailyTableRows.length > 0) {
        const dailyChartCol = dailyStartCol + 8; // Col 32 (AF)
        const dailyChart = summarySheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(summarySheet.getRange(3, dailyStartCol, dailyTableRows.length, 1))
            .addRange(summarySheet.getRange(3, dailyStartCol + 1, dailyTableRows.length, 1))
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

    // Driver stats charts (stacked over Driver Stats table at startCol = 13 = Col M)
    const statsDataStart = 4;
    const statsLastRow = 3 + uniqueDrivers.length;
    const colDriver = "M";
    const colCredit = "N";
    const colTrips = "P";
    const colNoShow = "R";
    const colHours = "S";
    const colAvgPerDay = "V";

    if (uniqueDrivers.length > 0) {
        // Chart 2: Total Credit by Driver (Column)
        const chartCredit = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
            .addRange(summarySheet.getRange(`${colDriver}${statsDataStart}:${colDriver}${statsLastRow}`))
            .addRange(summarySheet.getRange(`${colCredit}${statsDataStart}:${colCredit}${statsLastRow}`))
            .setOption("title", "Total Credit by Driver")
            .setOption("titleTextStyle", { bold: true, fontSize: 24 })
            .setOption("colors", ["#1f77b4"]).setOption("legend", { position: "none" })
            .setOption("vAxis", { title: "Credit" })
            .setOption("width", 740)
            .setOption("height", 450)
            .setPosition(3, statsStartCol, 0, 0).build();
        summarySheet.insertChart(chartCredit);

        // Chart 3: Total Trips by Driver
        const chartTrips = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
            .addRange(summarySheet.getRange(`${colDriver}${statsDataStart}:${colDriver}${statsLastRow}`))
            .addRange(summarySheet.getRange(`${colTrips}${statsDataStart}:${colTrips}${statsLastRow}`))
            .setOption("title", "Total Trips by Driver")
            .setOption("titleTextStyle", { bold: true, fontSize: 24 })
            .setOption("colors", ["#8c564b"]).setOption("legend", { position: "none" })
            .setOption("vAxis", { title: "Trips" })
            .setOption("width", 740)
            .setOption("height", 450)
            .setPosition(26, statsStartCol, 0, 0).build();
        summarySheet.insertChart(chartTrips);

        // Chart 4: No Show by Driver
        const chartNoShow = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
            .addRange(summarySheet.getRange(`${colDriver}${statsDataStart}:${colDriver}${statsLastRow}`))
            .addRange(summarySheet.getRange(`${colNoShow}${statsDataStart}:${colNoShow}${statsLastRow}`))
            .setOption("title", "No Show Trips by Driver")
            .setOption("titleTextStyle", { bold: true, fontSize: 24 })
            .setOption("colors", ["#2ca02c"]).setOption("legend", { position: "none" })
            .setOption("vAxis", { title: "No Show Trips" })
            .setOption("width", 740)
            .setOption("height", 450)
            .setPosition(49, statsStartCol, 0, 0).build();
        summarySheet.insertChart(chartNoShow);

        // Chart 5: Total Working Hours by Driver
        const chartHours = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
            .addRange(summarySheet.getRange(`${colDriver}${statsDataStart}:${colDriver}${statsLastRow}`))
            .addRange(summarySheet.getRange(`${colHours}${statsDataStart}:${colHours}${statsLastRow}`))
            .setOption("title", "Total Working Hours by Driver")
            .setOption("titleTextStyle", { bold: true, fontSize: 24 })
            .setOption("colors", ["#ff7f0e"]).setOption("legend", { position: "none" })
            .setOption("vAxis", { title: "Hours" })
            .setOption("width", 740)
            .setOption("height", 450)
            .setPosition(72, statsStartCol, 0, 0).build();
        summarySheet.insertChart(chartHours);

        // Chart 6: Average Per Day by Driver
        const chartAvgDay = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
            .addRange(summarySheet.getRange(`${colDriver}${statsDataStart}:${colDriver}${statsLastRow}`))
            .addRange(summarySheet.getRange(`${colAvgPerDay}${statsDataStart}:${colAvgPerDay}${statsLastRow}`))
            .setOption("title", "Average Credit Per Day")
            .setOption("titleTextStyle", { bold: true, fontSize: 24 })
            .setOption("colors", ["#d62728"]).setOption("legend", { position: "none" })
            .setOption("vAxis", { title: "Daily Credit" })
            .setOption("width", 740)
            .setOption("height", 450)
            .setPosition(95, statsStartCol, 0, 0).build();
        summarySheet.insertChart(chartAvgDay);
    }
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

    // Priority 1: Check spreadsheet name for explicit month name (e.g. "October - Drivers Daily Balance")
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
    if (foundMonthIndex === -1) {
        const ymMatch = ssName.match(/\b20\d\d[-_](0?[1-9]|1[0-2])\b/);
        if (ymMatch) foundMonthIndex = parseInt(ymMatch[1], 10) - 1;
    }

    let foundYear = -1;
    const yearMatch = ssName.match(/\b(20\d\d)\b/);
    if (yearMatch) {
        foundYear = parseInt(yearMatch[1], 10);
    }

    // If year not in title, scan allRows or Summary sheet
    if (foundYear === -1 && Array.isArray(allRows) && allRows.length > 0) {
        for (let i = 0; i < allRows.length; i++) {
            const rawDate = allRows[i][0];
            const d = parseDate(rawDate);
            if (d && !isNaN(d.getTime()) && d.getFullYear() > 2000) {
                foundYear = d.getFullYear();
                break;
            }
        }
    }

    if (foundMonthIndex >= 0) {
        const now = new Date();
        const finalYear = foundYear > 2000 ? foundYear : (now.getFullYear() < 2026 ? 2026 : now.getFullYear());
        _ACTIVE_MONTH_NUM = foundMonthIndex + 1;
        return {
            monthIndex: foundMonthIndex,
            year: finalYear,
            monthName: monthNames[foundMonthIndex]
        };
    }

    // Priority 2: Extract from valid dates in allRows
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
            _ACTIVE_MONTH_NUM = bestMonth + 1;
            return {
                monthIndex: bestMonth,
                year: bestYear,
                monthName: monthNames[bestMonth]
            };
        }
    }

    // Priority 2b: Extract from Raw Data sheet date headers
    try {
        if (ss && typeof ss.getSheetByName === "function") {
            const rawSheet = ss.getSheetByName("Raw Data");
            if (rawSheet) {
                const rawWeeks = parseRawDataWeeks(rawSheet);
                if (rawWeeks.length > 0) {
                    const monthCounts = {};
                    const yearCounts = {};
                    rawWeeks.forEach(w => {
                        w.dateColumns.forEach(dc => {
                            if (dc.date && !isNaN(dc.date.getTime())) {
                                const m = dc.date.getMonth();
                                const y = dc.date.getFullYear();
                                monthCounts[m] = (monthCounts[m] || 0) + 1;
                                yearCounts[y] = (yearCounts[y] || 0) + 1;
                            }
                        });
                    });
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
                        _ACTIVE_MONTH_NUM = bestMonth + 1;
                        return {
                            monthIndex: bestMonth,
                            year: bestYear,
                            monthName: monthNames[bestMonth]
                        };
                    }
                }
            }
        }
    } catch (eRawMonth) { }

    const now = new Date();
    const finalMonthIndex = now.getMonth();
    const finalYear = now.getFullYear();

    _ACTIVE_MONTH_NUM = finalMonthIndex + 1;

    return {
        monthIndex: finalMonthIndex,
        year: finalYear,
        monthName: monthNames[finalMonthIndex]
    };
}


// =================================================================
// WEEKLY SUMMARY — Generates weekly performance breakdown
// =================================================================
function generateWeeklySummary(ss, dailyData) {
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
        const summaryHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 12)).getDisplayValues()[0];
        const curCreditCol = findMainSummaryCol(summaryHeaders, ['credit'], 2);
        const curTripsCol = findMainSummaryCol(summaryHeaders, ['trips', 'trip'], 6);
        const curCashCol = findMainSummaryCol(summaryHeaders, ['cash'], 7);
        const curTollsCol = findMainSummaryCol(summaryHeaders, ['tolls', 'toll'], 8);
        const maxCol = Math.max(curCreditCol, curTripsCol, curCashCol, curTollsCol) + 1;

        const rawRows = summarySheet.getRange(3, 1, lastRowCurrent - 2, maxCol).getValues();
        let lastSeenDate = null;
        allRows = rawRows.map(r => {
            if (r[0] && r[0] !== "") lastSeenDate = r[0];
            else r[0] = lastSeenDate;
            return [
                r[0],
                r[1],
                parseNumber(r[curCreditCol]) || 0,
                parseNumber(r[curTollsCol]) || 0,
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
                    const prevHeaders = prevSummarySheet.getRange(2, 1, 1, Math.min(prevSummarySheet.getLastColumn(), 12)).getDisplayValues()[0];
                    const prevCreditCol = findMainSummaryCol(prevHeaders, ['credit'], 2);
                    const prevTripsCol = findMainSummaryCol(prevHeaders, ['trips', 'trip'], 6);
                    const prevCashCol = findMainSummaryCol(prevHeaders, ['cash'], 7);
                    const prevTollsCol = findMainSummaryCol(prevHeaders, ['tolls', 'toll'], -1);
                    const prevMaxCol = Math.max(prevCreditCol, prevTripsCol, prevCashCol, prevTollsCol >= 0 ? prevTollsCol : 0) + 1;
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
                            prevTollsCol >= 0 ? (parseNumber(row[prevTollsCol]) || 0) : 0,
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
    const dateIdx = 0, driverIdx = 1, creditIdx = 2, tollsIdx = 3, tripsIdx = 4, cashIdx = 5;

    allRows.forEach((row) => {
        if (!Array.isArray(row) || !row[dateIdx] || !row[driverIdx]) return;
        const date = parseDate(row[dateIdx]);
        if (!date) return;

        const driver = row[driverIdx];
        const credit = parseNumber(row[creditIdx]) || 0;
        const tolls = parseNumber(row[tollsIdx]) || 0;
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
            weeklyData[weekKey].drivers[driver] = { credit: 0, tolls: 0, trips: 0, cash: 0, hours: 0 };
        }

        weeklyData[weekKey].drivers[driver].credit += credit;
        weeklyData[weekKey].drivers[driver].tolls += tolls;
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

            const titleRange = weeklySheet.getRange(currentRow, 1, 1, 8).merge();
            titleRange.setValue(`WEEK ${displayedWeekCounter} (${weekKey})`)
                .setFontWeight("bold")
                .setFontSize(11)
                .setBackground("#d0e0e3");

            currentRow++;

            const headerRange = weeklySheet.getRange(currentRow, 1, 1, 8);
            headerRange.setValues([["Driver", "Total Credit", "Tolls", "Trips", "Total Cash", "Total Hours", "Avg Per Hour", "Balance"]])
                .setFontWeight("bold")
                .setBackground("#3c78d8")
                .setFontColor("white")
                .setHorizontalAlignment("center");

            currentRow++;

            const drivers = Object.keys(weekBlock.drivers).sort();
            const tableData = [];

            drivers.forEach((d, dIdx) => {
                const info = weekBlock.drivers[d];
                const r = currentRow + dIdx;
                const fare = driverFareMap[d] !== undefined ? driverFareMap[d] : 0.9;

                const formulaAvgPerHour = `=IF(F${r}>0, ROUND(B${r}/F${r}, 2), 0)`;
                let formulaBalance = "";
                if (fare === "80%-90%") {
                    formulaBalance = `=ROUND(IF((B${r}-C${r})<=1000, (B${r}-C${r})*0.8, 800 + (B${r}-C${r}-1000)*0.9) + C${r} - E${r}, 2)`;
                } else {
                    formulaBalance = `=ROUND((B${r}-C${r})*IFERROR(VLOOKUP(A${r}, Settings!$A:$B, 2, FALSE), 0.9) + C${r} - E${r}, 2)`;
                }

                tableData.push([
                    d,
                    roundToTwo(info.credit),
                    roundToTwo(info.tolls),
                    info.trips,
                    roundToTwo(info.cash),
                    roundToTwo(info.hours),
                    formulaAvgPerHour,
                    formulaBalance
                ]);
            });

            if (tableData.length > 0) {
                // Write Main Table
                const dataRange = weeklySheet.getRange(currentRow, 1, tableData.length, 8);
                dataRange.setValues(tableData);
                dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

                weeklySheet.getRange(currentRow, 2, tableData.length, 2).setNumberFormat("$#,##0.00"); // Credit, Tolls
                weeklySheet.getRange(currentRow, 4, tableData.length, 1).setNumberFormat("0");         // Trips
                weeklySheet.getRange(currentRow, 5, tableData.length, 1).setNumberFormat("$#,##0.00"); // Cash
                weeklySheet.getRange(currentRow, 6, tableData.length, 1).setNumberFormat("0.0");       // Hours
                weeklySheet.getRange(currentRow, 7, tableData.length, 2).setNumberFormat("$#,##0.00"); // Avg Per Hour, Balance

                const startR = currentRow;
                const endR = currentRow + tableData.length - 1;
                const totR = endR + 1;

                const totalRow = [
                    "Total",
                    `=SUM(B${startR}:B${endR})`,
                    `=SUM(C${startR}:C${endR})`,
                    `=SUM(D${startR}:D${endR})`,
                    `=SUM(E${startR}:E${endR})`,
                    `=SUM(F${startR}:F${endR})`,
                    `=IF(F${totR}>0, ROUND(B${totR}/F${totR}, 2), 0)`,
                    `=SUM(H${startR}:H${endR})`
                ];

                const totalRange = weeklySheet.getRange(totR, 1, 1, 8);
                totalRange.setValues([totalRow])
                    .setFontWeight("bold")
                    .setBackground("#eeeeee")
                    .setBorder(true, true, true, true, true, true);

                weeklySheet.getRange(totR, 2, 1, 2).setNumberFormat("$#,##0.00");
                weeklySheet.getRange(totR, 4, 1, 1).setNumberFormat("0");
                weeklySheet.getRange(totR, 5, 1, 1).setNumberFormat("$#,##0.00");
                weeklySheet.getRange(totR, 6, 1, 1).setNumberFormat("0.0");
                weeklySheet.getRange(totR, 7, 1, 2).setNumberFormat("$#,##0.00");

                const tableEnd = totR;

                // Chart 1: Total Credit by Driver (Col J = col 10, width 480px, height 310px)
                try {
                    const chart1 = weeklySheet.newChart().asColumnChart()
                        .setPosition(startRow - 1, 10, 0, 0)
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

                // Chart 2: Trips by Driver (Col O = col 15, width 480px, height 310px)
                try {
                    const chart2 = weeklySheet.newChart().asColumnChart()
                        .setPosition(startRow - 1, 15, 0, 0)
                        .addRange(weeklySheet.getRange(currentRow - 1, 1, tableData.length + 1, 1)) // Driver Names
                        .addRange(weeklySheet.getRange(currentRow - 1, 4, tableData.length + 1, 1)) // Trips
                        .setNumHeaders(1)
                        .setOption("title", `WEEK ${displayedWeekCounter} - Trips`)
                        .setOption("colors", ["#8c564b"])
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
        weeklySheet.getRange("A3:H3").merge()
            .setValue("No weekly data available yet for this month.")
            .setFontStyle("italic")
            .setHorizontalAlignment("center")
            .setBackground("#f9f9f9");
    } else {
        weeklySheet.autoResizeColumns(1, 8);
        for (let col = 1; col <= 8; col++) {
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
function generateBonusReport(ss, dailyData) {
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

        const totalRow = bonusSheet.getLastRow() + 2;
        const lastDataRow = totalRow - 2;
        const totalsBlock = bonusSheet.getRange(totalRow, 1, 2, 2);

        bonusSheet.getRange(totalRow, 1, 2, 1).setValues([["Grand Total"], ["0.5% Bonus"]]);
        bonusSheet.getRange(totalRow, 2, 2, 1).setValues([
            [`=SUM(B2:B${lastDataRow})`],
            [`=ROUND(B${totalRow}*0.005, 2)`]
        ]);

        totalsBlock.setFontWeight("bold")
            .setBackground("#cfe2f3")
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
            .setHorizontalAlignment("center");

        bonusSheet.getRange(totalRow, 2, 2, 1).setNumberFormat("$#,##0.00");

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
 * Parses date format from Raw Data headers.
 * Supports ISO "YYYY-MM-DD" (e.g. "2026-10-05" -> October 5, 2026), "DD-MM-YY", "DD-MM-YYYY", etc.
 */
function parseDDMMYY(value, activeMonthNum) {
    if (!value) return null;
    return parseDate(value, activeMonthNum);
}

// =================================================================
// ANGEL'S PROGRESSIVE REPORT
// =================================================================
function generateAngelReport(ss, dailyData) {
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
            const tollsIdx = colIndices ? colIndices.tolls : 8;
            const tIdx = colIndices ? colIndices.trips : 6;
            const caIdx = colIndices ? colIndices.cash : 7;
            const nsIdx = colIndices ? colIndices.noShow : 9;
            angelDataMap[dateKey] = {
                credit: parseNumber(row[cIdx]) || 0,
                tolls: (tollsIdx >= 0) ? (parseNumber(row[tollsIdx]) || 0) : 0,
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
                    const prevHeaders = prevSummarySheet.getRange(2, 1, 1, Math.min(prevSummarySheet.getLastColumn(), 12)).getDisplayValues()[0];
                    const prevIndices = {
                        credit: findMainSummaryCol(prevHeaders, ['credit'], 2),
                        tolls: findMainSummaryCol(prevHeaders, ['tolls', 'toll'], -1),
                        trips: findMainSummaryCol(prevHeaders, ['trips', 'trip'], 6),
                        cash: findMainSummaryCol(prevHeaders, ['cash'], 7),
                        noShow: findMainSummaryCol(prevHeaders, ['no show', 'noshow'], 8)
                    };
                    const prevData = prevSummarySheet.getRange(3, 1, prevSummarySheet.getLastRow() - 2, 12).getValues();

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
        const curHeaders = summarySheet.getRange(2, 1, 1, Math.min(summarySheet.getLastColumn(), 12)).getDisplayValues()[0];
        const curIndices = {
            credit: findMainSummaryCol(curHeaders, ['credit'], 2),
            tolls: findMainSummaryCol(curHeaders, ['tolls', 'toll'], 8),
            trips: findMainSummaryCol(curHeaders, ['trips', 'trip'], 6),
            cash: findMainSummaryCol(curHeaders, ['cash'], 7),
            noShow: findMainSummaryCol(curHeaders, ['no show', 'noshow'], 9)
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

    let weeklyCredit = 0, weeklyTolls = 0, weeklyFeeBase = 0, weeklyNoShow = 0, weeklyTrips = 0, weeklyCash = 0;
    let weeklyGrossBal = 0, weeklyFee = 0, weeklyNet = 0;

    let grandCredit = 0, grandTolls = 0, grandFeeBase = 0, grandNoShow = 0, grandTrips = 0, grandCash = 0;
    let grandGrossBal = 0, grandFee = 0, grandNet = 0;

    let weekTotalRows = [];

    const threshold = 1000;
    const baseRate = 0.80;
    const topRate = 0.90;
    const transferFeeRate = 0.03;

    for (let d = new Date(reportStartDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        const dateKey = d.toDateString();
        const dayData = angelDataMap[dateKey] || { credit: 0, tolls: 0, cash: 0, noShow: 0, trips: 0 };

        const todayCredit = dayData.credit;
        const todayTolls = dayData.tolls;
        const todayCash = dayData.cash;
        const feeBase = Math.max(0, todayCredit - todayTolls);

        let todayFeeBasePayout = 0;
        let appliedRate = 0;

        if (feeBase > 0) {
            if (weeklyFeeBase >= threshold) {
                todayFeeBasePayout = feeBase * topRate;
            } else if ((weeklyFeeBase + feeBase) <= threshold) {
                todayFeeBasePayout = feeBase * baseRate;
            } else {
                const creditAtBaseRate = threshold - weeklyFeeBase;
                const creditAtTopRate = feeBase - creditAtBaseRate;
                todayFeeBasePayout = (creditAtBaseRate * baseRate) + (creditAtTopRate * topRate);
            }
            appliedRate = todayFeeBasePayout / feeBase;
        }

        weeklyFeeBase += feeBase;
        weeklyCredit += todayCredit;
        weeklyTolls += todayTolls;

        const todayGrossPayout = todayFeeBasePayout + todayTolls;
        const grossBalance = todayGrossPayout - todayCash;
        const fee = grossBalance > 0 ? (grossBalance * transferFeeRate) : 0;
        const netPayout = grossBalance - fee;

        reportRows.push([
            new Date(d),
            todayCredit,
            todayTolls,
            weeklyFeeBase,
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
            grandTolls += weeklyTolls;
            grandFeeBase += weeklyFeeBase;
            grandNoShow += weeklyNoShow;
            grandTrips += weeklyTrips;
            grandCash += weeklyCash;
            grandGrossBal += weeklyGrossBal;
            grandFee += weeklyFee;
            grandNet += weeklyNet;

            let weekLabel = `Week ${weekCounter} Total`;
            let rowColor = "#e6f4ea";
            if (weeklyFeeBase > threshold) rowColor = "#fff2cc";

            let weeklyEffectiveRate = weeklyCredit > 0 ? ((weeklyGrossBal + weeklyCash) / weeklyCredit) : 0;

            reportRows.push([
                weekLabel,
                weeklyCredit,
                weeklyTolls,
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

            weeklyCredit = 0; weeklyTolls = 0; weeklyFeeBase = 0; weeklyNoShow = 0; weeklyTrips = 0; weeklyCash = 0;
            weeklyGrossBal = 0; weeklyFee = 0; weeklyNet = 0;

            if (!isLastDayOfReport) {
                weekCounter++;
                reportRows.push(["", "", "", "", "", "", "", "", "", "", ""]);
            }
        }
    }

    reportRows.push(["", "", "", "", "", "", "", "", "", "", ""]);
    const grandRate = grandCredit > 0 ? ((grandGrossBal + grandCash) / grandCredit) : 0;
    reportRows.push([
        "GRAND TOTAL", grandCredit, grandTolls, "", grandRate, grandNoShow, grandTrips, grandCash, grandGrossBal, grandFee, grandNet
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

    const headers = [["Date", "Daily Credit", "Tolls", "Weekly Cum.", "Applied Rate", "No Show", "Trips", "Cash", "Gross Bal.", "Fee (3%)", "Net Payout"]];

    reportSheet.getRange("A1:K1").setValues(headers)
        .setFontWeight("bold")
        .setFontSize(11)
        .setHorizontalAlignment("center")
        .setBackground("#f3f3f3")
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    if (reportRows.length > 0) {
        const range = reportSheet.getRange(2, 1, reportRows.length, 11);
        range.setValues(reportRows);

        reportSheet.getRange(2, 1, reportRows.length, 1).setNumberFormat("dddd, dd");
        reportSheet.getRange(2, 2, reportRows.length, 3).setNumberFormat("$#,##0.00"); // Credit, Tolls, Weekly Cum.
        reportSheet.getRange(2, 5, reportRows.length, 1).setNumberFormat("0.0%");       // Applied Rate
        reportSheet.getRange(2, 6, reportRows.length, 2).setNumberFormat("0");           // No Show, Trips
        reportSheet.getRange(2, 8, reportRows.length, 4).setNumberFormat("$#,##0.00"); // Cash, Gross Bal, Fee, Net Payout

        range.setHorizontalAlignment("center");

        weekTotalRows.forEach(item => {
            const rowRange = reportSheet.getRange(item.index, 1, 1, 11);
            rowRange.setFontWeight("bold");
            rowRange.setBackground(item.color);
            rowRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
            reportSheet.getRange(item.index, 1).setHorizontalAlignment("center");
        });

        const grandRowRange = reportSheet.getRange(grandTotalRowIndex, 1, 1, 11);
        grandRowRange.setFontWeight("bold");
        grandRowRange.setFontSize(12);
        grandRowRange.setBackground("#d9d2e9");
        grandRowRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

        reportSheet.autoResizeColumns(1, 11);
        for (let i = 1; i <= 11; i++) {
            let currentWidth = reportSheet.getColumnWidth(i);
            reportSheet.setColumnWidth(i, currentWidth + 15);
        }
        if (reportSheet.getColumnWidth(1) < 120) reportSheet.setColumnWidth(1, 120);

        if (chartDataRows.length > 0) {
            const chartDataRange = reportSheet.getRange(2, 13, chartDataRows.length, 2);
            chartDataRange.setValues(chartDataRows);
            reportSheet.getRange(2, 13, chartDataRows.length, 1).setNumberFormat("dddd, dd");
            reportSheet.getRange(2, 14, chartDataRows.length, 1).setNumberFormat("$#,##0");
            reportSheet.hideColumns(13, 2);
        }

        let angelChart = reportSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(reportSheet.getRange(2, 13, chartDataRows.length, 1))
            .addRange(reportSheet.getRange(2, 14, chartDataRows.length, 1))
            .setPosition(2, 13, 0, 0)
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

function parseDate(value, activeMonthNum) {
    if (!value && value !== 0) return null;
    if (value instanceof Date || Object.prototype.toString.call(value) === '[object Date]' || (typeof value === 'object' && value !== null && typeof value.getTime === 'function')) {
        return new Date(value.getFullYear(), value.getMonth(), value.getDate());
    }
    // Excel/Sheets serial day number (e.g. ~46000 for 2026 dates)
    if (typeof value === 'number' && value > 30000 && value < 60000) {
        const d = new Date(Math.round((value - 25569) * 86400 * 1000));
        return new Date(d.getFullYear(), d.getMonth(), d.getDate());
    }
    if (typeof value === "string") {
        let str = value.trim();
        if (!str) return null;

        // Normalize unicode hyphens/dashes to standard hyphen
        str = str.replace(/[\u2010-\u2015\u2212]/g, '-');

        // Resolve activeMonthNum: parameter, or global _ACTIVE_MONTH_NUM
        const targetMonth = (typeof activeMonthNum === 'number' && activeMonthNum >= 1 && activeMonthNum <= 12)
            ? activeMonthNum
            : _ACTIVE_MONTH_NUM;

        // 1. Strings with month names like "Tuesday, September 1, 2026" or "October 5, 2026"
        if (/[a-zA-Z]/.test(str)) {
            const std = new Date(str);
            if (!isNaN(std.getTime())) {
                return new Date(std.getFullYear(), std.getMonth(), std.getDate());
            }
        }

        // Split by delimiter: -, /, or .
        let delimiter = null;
        if (str.includes('-')) delimiter = '-';
        else if (str.includes('/')) delimiter = '/';
        else if (str.includes('.')) delimiter = '.';

        if (delimiter) {
            const parts = str.split(delimiter);
            if (parts.length === 3) {
                const p0 = parseInt(parts[0], 10);
                const p1 = parseInt(parts[1], 10);
                const p2 = parseInt(parts[2], 10);

                if (!isNaN(p0) && !isNaN(p1) && !isNaN(p2)) {
                    // Case A: 4-digit year at start (YYYY-MM-DD or YYYY/MM/DD or YYYY.MM.DD)
                    // e.g. "2026-10-05" -> Year: 2026, Month: 10 (October), Day: 5
                    if (p0 >= 1000) {
                        let year = p0;
                        let month, day;
                        if (targetMonth) {
                            if (p1 === targetMonth && p2 !== targetMonth) {
                                month = p1 - 1;
                                day = p2;
                            } else if (p2 === targetMonth && p1 !== targetMonth) {
                                day = p1;
                                month = p2 - 1;
                            } else if (p1 > 12 && p2 <= 12) {
                                day = p1;
                                month = p2 - 1;
                            } else {
                                month = p1 - 1;
                                day = p2;
                            }
                        } else {
                            if (p1 > 12 && p2 <= 12) {
                                day = p1;
                                month = p2 - 1;
                            } else {
                                month = p1 - 1;
                                day = p2;
                            }
                        }
                        if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                            const d = new Date(year, month, day);
                            if (!isNaN(d.getTime())) return d;
                        }
                    }

                    // Case B: 4-digit year at end (e.g. 10-05-2026 or 10/05/2026)
                    if (p2 >= 1000) {
                        let year = p2;
                        let month, day;
                        if (p0 > 12 && p1 <= 12) {
                            // DD-MM-YYYY
                            day = p0;
                            month = p1 - 1;
                        } else if (p1 > 12 && p0 <= 12) {
                            // MM-DD-YYYY
                            month = p0 - 1;
                            day = p1;
                        } else if (targetMonth) {
                            if (p0 === targetMonth && p1 !== targetMonth) {
                                // MM-DD-YYYY matches active month
                                month = p0 - 1;
                                day = p1;
                            } else if (p1 === targetMonth && p0 !== targetMonth) {
                                // DD-MM-YYYY matches active month
                                day = p0;
                                month = p1 - 1;
                            } else {
                                // Default US convention MM-DD-YYYY or MM/DD/YYYY
                                month = p0 - 1;
                                day = p1;
                            }
                        } else {
                            // Default US convention MM-DD-YYYY or MM/DD/YYYY
                            month = p0 - 1;
                            day = p1;
                        }
                        if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                            const d = new Date(year, month, day);
                            if (!isNaN(d.getTime())) return d;
                        }
                    }

                    // Case C: 2-digit years
                    // C1: YY-MM-DD (e.g. "26-10-05" -> year 2026, month 10, day 5)
                    if (p0 >= 20 && p0 <= 50 && p2 < 20) {
                        const year = 2000 + p0;
                        let month, day;
                        if (targetMonth) {
                            if (p1 === targetMonth && p2 !== targetMonth) {
                                month = p1 - 1;
                                day = p2;
                            } else if (p2 === targetMonth && p1 !== targetMonth) {
                                day = p1;
                                month = p2 - 1;
                            } else {
                                month = p1 - 1;
                                day = p2;
                            }
                        } else {
                            month = p1 - 1;
                            day = p2;
                        }
                        if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                            const d = new Date(year, month, day);
                            if (!isNaN(d.getTime())) return d;
                        }
                    }

                    // C2: 2-digit year at end (e.g. "10-05-26" -> month 10, day 5, year 2026, or "05-10-26")
                    if (p2 >= 20 && p2 <= 50) {
                        const year = 2000 + p2;
                        let month, day;
                        if (p0 > 12 && p1 <= 12) {
                            // p0 can only be day (e.g. 25-10-26)
                            day = p0;
                            month = p1 - 1;
                        } else if (p1 > 12 && p0 <= 12) {
                            // p1 can only be day (e.g. 10-25-26)
                            month = p0 - 1;
                            day = p1;
                        } else if (targetMonth) {
                            if (p0 === targetMonth && p1 !== targetMonth) {
                                // p0 is the active month! (e.g. 10-05-26 in October -> Oct 5)
                                month = p0 - 1;
                                day = p1;
                            } else if (p1 === targetMonth && p0 !== targetMonth) {
                                // p1 is the active month! (e.g. 05-10-26 in October -> Oct 5)
                                day = p0;
                                month = p1 - 1;
                            } else {
                                // Default US format MM-DD-YY
                                month = p0 - 1;
                                day = p1;
                            }
                        } else {
                            // Default US format MM-DD-YY
                            month = p0 - 1;
                            day = p1;
                        }
                        if (month >= 0 && month < 12 && day >= 1 && day <= 31) {
                            const d = new Date(year, month, day);
                            if (!isNaN(d.getTime())) return d;
                        }
                    }
                }
            }
        }

        const fallback = new Date(str);
        if (!isNaN(fallback.getTime())) {
            return new Date(fallback.getFullYear(), fallback.getMonth(), fallback.getDate());
        }
    }
    return null;
}