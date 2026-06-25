// =================================================================
// MASTER SCRIPT — CENTRALIZED BALANCE SHEET CONTROLLER
// =================================================================
// Change this ID each time you switch to a new monthly balance file.
// This is the ONLY line you need to update.
const TARGET_SHEET_ID = "1LG44Ry-BjAv2HCiZtkIcdI24zw43XbEH1WPDVCwUkC8";

/**
 * Run this function ONCE every time you change the TARGET_SHEET_ID above.
 * It sets up the script to run automatically whenever the target sheet gets updated.
 */
function setupAutoRun() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

    // 1. Delete any old triggers so we don't get duplicate runs
    const triggers = ScriptApp.getProjectTriggers();
    triggers.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    // 2. Create a new trigger that watches the target spreadsheet for changes
    ScriptApp.newTrigger('runDailyBalance')
        .forSpreadsheet(ss)
        .onChange()
        .create();

    Logger.log("✅ Auto-run successfully set up for: " + ss.getName());
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
}


// =================================================================
// SUMMARY BUILD — Reads "Raw Data" tab and builds the Summary sheet
// =================================================================
function updateSummaryAndCharts(ss) {
    const summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) throw new Error('No sheet named "Summary" was found.');

    const rawSheet = ss.getSheetByName("Raw Data");
    if (!rawSheet) throw new Error('No sheet named "Raw Data" was found.');

    // --- Read the entire Raw Data sheet ---
    const rawData = rawSheet.getDataRange().getValues();
    if (rawData.length < 3) {
        Logger.log("⚠️ Raw Data has fewer than 3 rows (need header + Total + at least 1 driver).");
        return;
    }

    const headerRow = rawData[0]; // Driver | 01-06-26 | 02-06-26 | ... | Total | Count | Driver NET
    const numCols = headerRow.length;

    // --- Identify date columns vs summary columns ---
    // Last 3 columns are: Total, Count, Driver NET
    // Everything between column 1 and (numCols - 3) are date columns
    const totalColIdx = numCols - 3;
    const countColIdx = numCols - 2;
    const netColIdx = numCols - 1;

    // Parse date headers (DD-MM-YY format)
    const dateColumns = []; // { colIdx, date }
    for (let c = 1; c < totalColIdx; c++) {
        const dateVal = parseDDMMYY(headerRow[c]);
        if (dateVal) {
            dateColumns.push({ colIdx: c, date: dateVal });
        }
    }

    if (dateColumns.length === 0) {
        Logger.log("⚠️ No valid date columns found in Raw Data headers.");
        return;
    }

    // --- Build output rows: one row per driver per day (only days with credit > 0) ---
    let outputRows = []; // [date, driver, credit]

    for (let r = 2; r < rawData.length; r++) { // Skip header (0) and Total row (1)
        const driverName = rawData[r][0];
        if (!driverName || driverName.toString().trim() === "" || driverName === "Total") continue;

        const driverTotal = parseNumber(rawData[r][totalColIdx]) || 0;
        const driverTrips = parseNumber(rawData[r][countColIdx]) || 0;
        const driverNet = parseNumber(rawData[r][netColIdx]) || 0;
        const driverCash = roundToTwo(driverTotal - driverNet); // Cash = Total - NET

        dateColumns.forEach(dc => {
            const credit = parseNumber(rawData[r][dc.colIdx]) || 0;
            if (credit > 0) {
                outputRows.push([
                    dc.date,
                    driverName,
                    credit,
                    driverCash,    // Monthly cash total (stored for aggregation)
                    driverTrips,   // Monthly trip count (stored for aggregation)
                    driverTotal    // Monthly total (stored for aggregation)
                ]);
            }
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
    const headers = ["Date", "Driver", "Credit", "Trips"];
    summarySheet.getRange(2, 1, 1, headers.length).setValues([headers]);

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

        // Summary rows: Date, Driver, Credit, Trips (trips = 0 per row since we only have monthly total)
        rowsToWrite.push([displayDate, row[1], row[2], ""]);
    });

    if (currentGroupSize > 0) {
        dayBlocks.push({ row: currentStartRow, count: currentGroupSize });
    }

    // --- Clear old data and write new ---
    const maxRows = summarySheet.getMaxRows();
    if (maxRows > 2) {
        summarySheet.getRange(3, 1, maxRows - 2, 4).clearContent();
        try { summarySheet.getRange(3, 1, maxRows - 2, 1).breakApart(); } catch (e) { }
    }

    const dataRange = summarySheet.getRange(3, 1, rowsToWrite.length, 4);
    dataRange.setValues(rowsToWrite);
    dataRange.setBorder(false, false, false, false, false, false);
    dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    summarySheet.getRange(3, 1, rowsToWrite.length, 1).setVerticalAlignment("middle");

    // --- Apply day block borders and merge date cells ---
    dayBlocks.forEach(block => {
        summarySheet.getRange(block.row, 1, block.count, 4)
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
        const driverCash = roundToTwo(driverTotal - driverNet);

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
            driverNet: roundToTwo(driverNet),
            workingDays: workingDays,
            avgPerDay: workingDays > 0 ? roundToTwo(driverTotal / workingDays) : 0,
            firstWorkDate: firstWorkDate
        };
    }

    // =================================================================
    // --- AUTO-SYNC "SETTINGS" SHEET ---
    // =================================================================
    let settingsSheet = ss.getSheetByName("Settings");
    if (!settingsSheet) settingsSheet = ss.insertSheet("Settings");

    const settingsData = settingsSheet.getDataRange().getValues();
    let existingSettings = {};
    for (let i = 1; i < settingsData.length; i++) {
        if (settingsData[i][0]) {
            existingSettings[settingsData[i][0]] = {
                system: settingsData[i][1],
                startDate: settingsData[i][2] || "",
                promoTaken: settingsData[i][3] || 0,
                promoLedger: settingsData[i][4] || ""
            };
        }
    }

    // --- Aggregated stats table (columns M+) ---
    const aggHeader = ["Driver", "Total Credit", "Total Trips", "Total Cash", "Driver NET", "Avg Per Day"];
    const aggRows = [aggHeader];

    const newSettingsRows = [["Driver", "System", "Start Date", "Promo Weeks Taken", "Promo Ledger"]];

    for (let driver in driverStats) {
        const s = driverStats[driver];
        aggRows.push([driver, s.totalCredit, s.totalTrips, s.totalCash, s.driverNet, s.avgPerDay]);

        let dInfo = existingSettings[driver];
        let systemStatus = dInfo ? dInfo.system : "";
        let startDate = dInfo ? dInfo.startDate : "";
        let promoTaken = dInfo ? dInfo.promoTaken : 0;
        let promoLedger = dInfo ? dInfo.promoLedger : "";

        // Auto-detect New Driver
        if (!systemStatus || systemStatus.toString().trim() === "") {
            systemStatus = "New Driver";
        }

        // Auto-detect Start Date if blank
        if (!startDate || startDate.toString().trim() === "") {
            if (s.firstWorkDate) {
                const m = (s.firstWorkDate.getMonth() + 1).toString().padStart(2, "0");
                const d = s.firstWorkDate.getDate().toString().padStart(2, "0");
                const y = s.firstWorkDate.getFullYear();
                startDate = `${m}/${d}/${y}`;
            }
        } else if (startDate instanceof Date) {
            const m = (startDate.getMonth() + 1).toString().padStart(2, "0");
            const d = startDate.getDate().toString().padStart(2, "0");
            const y = startDate.getFullYear();
            startDate = `${m}/${d}/${y}`;
        }

        newSettingsRows.push([driver, systemStatus, startDate, promoTaken, promoLedger]);
    }

    // --- MEMORY SAVER: Keep inactive drivers in Settings ---
    for (let oldDriver in existingSettings) {
        if (!driverStats[oldDriver]) {
            let oldInfo = existingSettings[oldDriver];
            newSettingsRows.push([
                oldDriver,
                oldInfo.system,
                oldInfo.startDate,
                oldInfo.promoTaken,
                oldInfo.promoLedger
            ]);
        }
    }

    settingsSheet.clear();
    settingsSheet.getRange(1, 1, newSettingsRows.length, 5).setValues(newSettingsRows);
    settingsSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#d0e0e3");
    settingsSheet.autoResizeColumns(1, 5);

    // =================================================================
    // --- Write aggregated stats to Summary (starting at column M) ---
    // =================================================================
    const startRow = 3, startCol = 13;
    // Clear old aggregate area
    try { summarySheet.getRange(startRow, startCol, 50, 8).clearContent(); } catch (e) { }
    summarySheet.getRange(startRow, startCol, aggRows.length, aggRows[0].length).setValues(aggRows);

    // --- Daily stats for chart (aggregate all drivers per day) ---
    let dailyStats = {};
    outputRows.forEach(r => {
        const dateVal = r[0];
        if (!dateVal) return;
        const dateKey = dateVal.getTime();
        if (!dailyStats[dateKey]) {
            dailyStats[dateKey] = { date: dateVal, credit: 0 };
        }
        dailyStats[dateKey].credit += (r[2] || 0);
    });

    const dailyRows = Object.values(dailyStats)
        .sort((a, b) => a.date - b.date)
        .map(d => [d.date, roundToTwo(d.credit)]);

    const dailyHeader = ["Date", "Total Credit"];
    const dailyStartCol = 22;

    summarySheet.getRange(2, dailyStartCol, 100, 5).clearContent();
    summarySheet.getRange(2, dailyStartCol, 1, dailyHeader.length).setValues([dailyHeader]).setFontWeight("bold");

    if (dailyRows.length > 0) {
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, dailyRows[0].length).setValues(dailyRows);
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, 1).setNumberFormat("dddd, dd");
    }

    // --- Charts ---
    const allCharts = summarySheet.getCharts();
    allCharts.forEach(c => summarySheet.removeChart(c));

    // Chart 1: Daily Total Credit Trend (line)
    if (dailyRows.length > 0) {
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
            .setPosition(2, 27, 0, 0)
            .setOption("width", 1000)
            .setOption("height", 400)
            .build();
        summarySheet.insertChart(dailyChart);
    }

    const lastRow = startRow + aggRows.length - 1;
    const dataStart = startRow + 1;

    // Chart 2: Total Credit by Driver (bar)
    let chart1 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("N" + dataStart + ":N" + lastRow))
        .setOption("title", "Total Credit by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#1f77b4"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Credit" }).setPosition(2, 13, 0, 0).build();
    summarySheet.insertChart(chart1);

    // Chart 3: Total Trips by Driver (bar)
    let chart2 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("O" + dataStart + ":O" + lastRow))
        .setOption("title", "Total Trips by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#8c564b"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Trips" }).setPosition(21, 13, 0, 0).build();
    summarySheet.insertChart(chart2);

    // Chart 4: Average Per Day by Driver (bar)
    let chart3 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("R" + dataStart + ":R" + lastRow))
        .setOption("title", "Average Credit Per Day")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#d62728"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Daily Credit" }).setPosition(39, 13, 0, 0).build();
    summarySheet.insertChart(chart3);

    // Chart 5: Driver NET by Driver (bar)
    let chart4 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("Q" + dataStart + ":Q" + lastRow))
        .setOption("title", "Driver NET Payout")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#2ca02c"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "NET ($)" }).setPosition(57, 13, 0, 0).build();
    summarySheet.insertChart(chart4);
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

    // --- Read Raw Data to get per-driver monthly cash ---
    const rawSheet = ss.getSheetByName("Raw Data");
    let driverCashMap = {}; // driver -> total monthly cash
    if (rawSheet) {
        const rawData = rawSheet.getDataRange().getValues();
        const numCols = rawData[0].length;
        const totalColIdx = numCols - 3;
        const netColIdx = numCols - 1;
        for (let r = 2; r < rawData.length; r++) {
            const name = rawData[r][0];
            if (!name || name === "Total") continue;
            const total = parseNumber(rawData[r][totalColIdx]) || 0;
            const net = parseNumber(rawData[r][netColIdx]) || 0;
            driverCashMap[name] = roundToTwo(total - net);
        }
    }

    // --- Get Main Data from Summary sheet ---
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
        } catch (e) { }
    }

    // --- Process Data ---
    const weeklyData = {};
    const dateIdx = 0, driverIdx = 1, creditIdx = 2;

    allRows.forEach((row) => {
        if (!Array.isArray(row) || !row[dateIdx] || !row[driverIdx]) return;
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

            const titleRange = weeklySheet.getRange(currentRow, 1, 1, 3).merge();
            titleRange.setValue(`WEEK ${displayedWeekCounter} (${weekKey})`)
                .setFontWeight("bold")
                .setFontSize(11)
                .setBackground("#d0e0e3");

            currentRow++;

            const headerRange = weeklySheet.getRange(currentRow, 1, 1, 4);
            headerRange.setValues([["Driver", "Total Credit", "Total Cash", "Balance"]])
                .setFontWeight("bold")
                .setBackground("#3c78d8")
                .setFontColor("white")
                .setHorizontalAlignment("center");

            currentRow++;

            const drivers = Object.keys(weekBlock);
            const tableData = [];

            drivers.forEach(d => {
                const info = weeklyData[weekKey][d];
                const cash = driverCashMap[d] || 0;
                const balance = roundToTwo(info.credit - cash);
                tableData.push([d, info.credit, cash, balance]);
            });

            // Write Main Table
            const dataRange = weeklySheet.getRange(currentRow, 1, tableData.length, 4);
            dataRange.setValues(tableData);
            dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            weeklySheet.getRange(currentRow, 2, tableData.length, 3).setNumberFormat("$#,##0.00");

            const totalCredit = tableData.reduce((a, b) => a + b[1], 0);
            const totalCash = tableData.reduce((a, b) => a + b[2], 0);
            const totalBalance = tableData.reduce((a, b) => a + b[3], 0);

            const totalRowRange = weeklySheet.getRange(currentRow + tableData.length, 1, 1, 4);
            totalRowRange.setValues([["TOTAL", totalCredit, totalCash, totalBalance]])
                .setFontWeight("bold")
                .setBackground("#eeeeee")
                .setBorder(true, true, true, true, true, true);

            weeklySheet.getRange(currentRow + tableData.length, 2, 1, 3).setNumberFormat("$#,##0.00");

            const tableEnd = currentRow + tableData.length;

            // Chart: Total Credit by Driver
            const chart = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 7, 0, 0)
                .addRange(weeklySheet.getRange(currentRow, 1, tableData.length, 2))
                .setOption("title", `WEEK ${displayedWeekCounter} - Total Credit`)
                .setOption("legend", { position: "none" })
                .setOption("hAxis", { title: "Driver" })
                .setOption("vAxis", { title: "Total Credit" })
                .build();
            weeklySheet.insertChart(chart);

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
 * Parses DD-MM-YY date format from Raw Data headers.
 * Example: "01-06-26" → June 1, 2026
 */
function parseDDMMYY(value) {
    if (!value) return null;
    if (value instanceof Date) return value;
    const str = String(value).trim();
    const parts = str.split("-");
    if (parts.length !== 3) return null;
    const day = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // zero-based
    let year = parseInt(parts[2], 10);
    if (year < 100) year += 2000; // 26 → 2026
    if (isNaN(day) || isNaN(month) || isNaN(year)) return null;
    const d = new Date(year, month, day);
    return isNaN(d.getTime()) ? null : d;
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