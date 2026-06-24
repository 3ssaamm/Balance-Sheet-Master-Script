// =================================================================
// MASTER SCRIPT — CENTRALIZED BALANCE SHEET CONTROLLER
// =================================================================
// Change this ID each time you switch to a new monthly balance file.
// This is the ONLY line you need to update.
const TARGET_SHEET_ID = "1xHIBxGjKy61suqKb1ly4WeVlnIRX1O0owv3KVWhYFok";

/**
 * Main entry point — runs all reporting functions in the correct order.
 * Attach your time-driven trigger to THIS function.
 */
function runDailyBalance() {
    const ss = SpreadsheetApp.openById(TARGET_SHEET_ID);

    // 1. Update the main Summary sheet.
    updateSummaryAndCharts(ss);

    // 2. Wait for all pending spreadsheet changes to apply.
    SpreadsheetApp.flush();

    // 3. Generate the weekly summary.
    generateWeeklySummary(ss);

    // 4. Generate Angel's specific report.
    generateAngelReport(ss);

    // 5. Generate the final Bonus report.
    generateBonusReport(ss);
}


// =================================================================
// SUMMARY BUILD — Aggregates driver data into the Summary sheet
// =================================================================
function updateSummaryAndCharts(ss) {
    const summarySheet = ss.getSheetByName("Summary");
    if (!summarySheet) throw new Error('No sheet named "Summary" was found.');

    // --- Read the editable start date for the custom report ---
    const startDateCell = summarySheet.getRange("V1");
    const startDateValue = startDateCell.getValue();
    let customStartDate = null;
    if (startDateValue instanceof Date && !isNaN(startDateValue)) {
        customStartDate = new Date(startDateValue);
        customStartDate.setHours(0, 0, 0, 0);
    }

    // --- Pull raw rows into Summary (A-L) ---
    const headers = [
        "Date", "Driver", "Credit", "Cash", "Balance",
        "Hours", "Hour Rate", "Trips", "REV", "Achieved", "Canceled",
    ];
    summarySheet.getRange(2, 1, 1, headers.length).setValues([headers]);

    let outputRows = [];
    const driverSheets = ss.getSheets().filter(sh => {
        const name = sh.getName();
        return !["Summary", "Weekly Summary", "Charts", "Dashboard", "Bonus", "Angel Summary", "Settings"].includes(name);
    });

    // row mapping (zero-based)
    const CASH_ROW = 21;
    const CREDIT_ROW = 22;
    const RATE_ROW = 23;    // The % Row
    const BALANCE_ROW = 24;
    const HOURS_ROW = 25;
    const HOURRATE_ROW = 26;
    const REV_ROW = 27;
    const ACH_ROW = 28;
    const CANCELED_ROW = 29;

    driverSheets.forEach(sheet => {
        const driver = sheet.getName();
        const data = sheet.getDataRange().getValues();
        if (!data || data.length === 0) return;
        const nCols = data[0].length;
        for (let col = 1; col < nCols; col += 2) {
            const dateCell = data[0][col] || data[0][col - 1];
            const creditCell = (data[CREDIT_ROW] ? data[CREDIT_ROW][col - 1] : null);
            if (!dateCell) continue;

            const dateVal = (dateCell instanceof Date) ? dateCell : new Date(dateCell);
            const cCredit = parseNumber(creditCell);
            const cash = parseNumber(data[CASH_ROW] ? data[CASH_ROW][col] : null);
            const cBalance = parseNumber(data[BALANCE_ROW] ? data[BALANCE_ROW][col] : null);
            const hours = parseNumber(data[HOURS_ROW] ? data[HOURS_ROW][col] : null);
            const hourRate = (hours && hours > 0) ? roundToTwo(cCredit / hours) : 0;

            let dailyRate = parseNumber(data[RATE_ROW] ? data[RATE_ROW][col - 1] : null);

            let tripsCount = 0;
            if (cCredit > 0) {
                for (let row = 3; row < CREDIT_ROW; row++) {
                    const tripCredit = parseNumber(data[row] ? data[row][col - 1] : null);
                    if (tripCredit !== null && tripCredit > 0) {
                        tripsCount++;
                    }
                }
            }
            const rev = parseNumber(data[REV_ROW] ? data[REV_ROW][col] : null);
            const achieved = parseNumber(data[ACH_ROW] ? data[ACH_ROW][col] : null);
            const canceled = parseNumber(data[CANCELED_ROW] ? data[CANCELED_ROW][col] : null);

            outputRows.push([
                dateVal, driver, cCredit, cash, cBalance,
                hours, hourRate, tripsCount, rev, achieved, canceled, dailyRate
            ]);
        }
    });

    if (outputRows.length === 0) {
        Logger.log("⚠️ No data collected.");
        return;
    }

    // SORT BY DATE CHRONOLOGICALLY
    outputRows.sort((a, b) => {
        const dateDiff = a[0] - b[0];
        if (dateDiff !== 0) return dateDiff;
        return a[1].localeCompare(b[1]);
    });

    const rowsToWrite = [];
    let previousDateString = "";
    let dayBlocks = [];
    let currentStartRow = 3;
    let currentGroupSize = 0;

    outputRows.forEach((row, index) => {
        let rowToPush = row.slice(0, 11);
        const currentDateString = row[0] ? row[0].toDateString() : "";

        if (currentDateString === previousDateString) {
            rowToPush[0] = "";
            currentGroupSize++;
        } else {
            if (currentGroupSize > 0) {
                dayBlocks.push({ row: currentStartRow, count: currentGroupSize });
            }
            previousDateString = currentDateString;
            currentStartRow = 3 + index;
            currentGroupSize = 1;
        }
        rowsToWrite.push(rowToPush);
    });

    if (currentGroupSize > 0) {
        dayBlocks.push({ row: currentStartRow, count: currentGroupSize });
    }

    summarySheet.getRange(3, 1, summarySheet.getMaxRows() - 2, 1).breakApart();
    const dataRange = summarySheet.getRange(3, 1, rowsToWrite.length, 11);
    dataRange.setValues(rowsToWrite);
    dataRange.setBorder(false, false, false, false, false, false);
    dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
    summarySheet.getRange(3, 1, rowsToWrite.length, 1).setVerticalAlignment("middle");

    dayBlocks.forEach(block => {
        summarySheet.getRange(block.row, 1, block.count, 11)
            .setBorder(true, null, true, null, null, null, "black", SpreadsheetApp.BorderStyle.SOLID_THICK);
        if (block.count > 1) {
            summarySheet.getRange(block.row, 1, block.count, 1).mergeVertically();
        }
    });

    // --- Aggregate per driver ---
    let driverStats = {};
    outputRows.forEach(r => {
        const dateVal = r[0];
        const driver = r[1];
        const credit = r[2] || 0;
        const hours = r[5] || 0;
        const trips = r[7] || 0;
        const achieved = r[9];
        const canceled = r[10] || 0;

        if (!driverStats[driver]) {
            driverStats[driver] = {
                totalCredit: 0, totalHours: 0, totalTrips: 0, totalCanceled: 0,
                sumAchieved: 0, achievedCount: 0, workingDaysSet: new Set(),
                firstWorkDate: null // To auto-detect start dates
            };
        }

        const s = driverStats[driver];
        s.totalCredit += credit;
        s.totalHours += (typeof hours === "number" ? hours : 0);
        s.totalTrips += trips;
        s.totalCanceled += canceled;

        if (achieved !== null && achieved !== undefined && achieved > 0) {
            s.sumAchieved += achieved;
            s.achievedCount++;
        }

        if (dateVal && credit > 0) {
            const d = new Date(dateVal);
            if (!isNaN(d.getTime())) {
                s.workingDaysSet.add(d.toDateString());
                // Log the earliest date
                if (!s.firstWorkDate || d < s.firstWorkDate) {
                    s.firstWorkDate = new Date(d);
                }
            }
        }
    });

    // =================================================================
    // --- AUTO-SYNC "SETTINGS" SHEET WITH PROMO LEDGER ---
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
                promoLedger: settingsData[i][4] || "" // Read the ledger
            };
        }
    }

    const aggHeader = ["Driver", "Total Credit", "Avg Achieved", "Hour Rate", "Total Hours", "Total Trips", "Total Canceled", "Avg Per Day"];
    const aggRows = [aggHeader];

    // 5-Column Header (Renamed column E)
    const newSettingsRows = [["Driver", "System", "Start Date", "Promo Weeks Taken", "Promo Ledger"]];

    for (let driver in driverStats) {
        const s = driverStats[driver];
        const avgAch = s.achievedCount > 0 ? roundToTwo(s.sumAchieved / s.achievedCount) : 0;
        const hourRateBased = s.totalHours > 0 ? roundToTwo(s.totalCredit / s.totalHours) : 0;
        const totalHours = roundToTwo(s.totalHours);
        const avgPerDay = s.workingDaysSet.size > 0 ? roundToTwo(s.totalCredit / s.workingDaysSet.size) : 0;
        aggRows.push([driver, roundToTwo(s.totalCredit), avgAch, hourRateBased, totalHours, s.totalTrips, s.totalCanceled, avgPerDay]);

        let dInfo = existingSettings[driver];
        let systemStatus = dInfo ? dInfo.system : "";
        let startDate = dInfo ? dInfo.startDate : "";
        let promoTaken = dInfo ? dInfo.promoTaken : 0;
        let promoLedger = dInfo ? dInfo.promoLedger : "";

        // Auto-detect New Driver
        if (!systemStatus || systemStatus.toString().trim() === "") {
            systemStatus = "New Driver";
        }

        // Auto-detect Start Date if it is currently blank
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
    // --- MEMORY SAVER: Keep inactive drivers in Settings so they aren't deleted ---
    for (let oldDriver in existingSettings) {
        if (!driverStats[oldDriver]) { // If they didn't drive this month
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

    const startRow = 3, startCol = 13;
    summarySheet.getRange(startRow, startCol, aggRows.length, aggRows[0].length).setValues(aggRows);

    let dailyStats = {};
    outputRows.forEach(r => {
        const dateVal = r[0];
        if (!dateVal) return;
        const dateKey = dateVal.getTime();
        if (!dailyStats[dateKey]) {
            dailyStats[dateKey] = { date: dateVal, credit: 0, hours: 0, trips: 0, canceled: 0 };
        }
        dailyStats[dateKey].credit += (r[2] || 0);
        dailyStats[dateKey].hours += (r[5] || 0);
        dailyStats[dateKey].trips += (r[7] || 0);
        dailyStats[dateKey].canceled += (r[10] || 0);
    });

    const dailyRows = Object.values(dailyStats)
        .sort((a, b) => a.date - b.date)
        .map(d => [d.date, roundToTwo(d.credit), roundToTwo(d.hours), d.trips, d.canceled]);

    const dailyHeader = ["Date", "Total Credit", "Total Hours", "Total Trips", "Total Canceled"];
    const dailyStartCol = 22;

    summarySheet.getRange(2, dailyStartCol, 100, 5).clearContent();
    summarySheet.getRange(2, dailyStartCol, 1, dailyHeader.length).setValues([dailyHeader]).setFontWeight("bold");

    if (dailyRows.length > 0) {
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, dailyRows[0].length).setValues(dailyRows);
        summarySheet.getRange(3, dailyStartCol, dailyRows.length, 1).setNumberFormat("dddd, dd");
    }

    const allCharts = summarySheet.getCharts();
    allCharts.forEach(c => summarySheet.removeChart(c));

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

    let chart1 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("Q" + dataStart + ":Q" + lastRow))
        .setOption("title", "Total Working Hours by Driver")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#9467bd"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Hours" }).setPosition(2, 13, 0, 0).build();
    summarySheet.insertChart(chart1);

    let chart2 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("P" + dataStart + ":P" + lastRow))
        .setOption("title", "Average Hour Rate")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#2ca02c"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Rate ($)" }).setPosition(21, 13, 0, 0).build();
    summarySheet.insertChart(chart2);

    let chart3 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("O" + dataStart + ":O" + lastRow))
        .setOption("title", "Average Achieved %")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#ff7f0e"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "%" }).setPosition(39, 13, 0, 0).build();
    summarySheet.insertChart(chart3);

    let chart4 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("N" + dataStart + ":N" + lastRow))
        .setOption("title", "Total Credit")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#1f77b4"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Credit" }).setPosition(57, 13, 0, 0).build();
    summarySheet.insertChart(chart4);

    let chart5 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("T" + dataStart + ":T" + lastRow))
        .setOption("title", "Average Per Day")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#d62728"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Daily Credit" }).setPosition(75, 13, 0, 0).build();
    summarySheet.insertChart(chart5);

    let chart6 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("R" + dataStart + ":R" + lastRow))
        .setOption("title", "Total Trips")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#8c564b"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Trips" }).setPosition(93, 13, 0, 0).build();
    summarySheet.insertChart(chart6);

    let chart7 = summarySheet.newChart().setChartType(Charts.ChartType.COLUMN)
        .addRange(summarySheet.getRange("M" + dataStart + ":M" + lastRow))
        .addRange(summarySheet.getRange("S" + dataStart + ":S" + lastRow))
        .setOption("title", "Total Canceled Trips")
        .setOption("titleTextStyle", { bold: true, fontSize: 24 })
        .setOption("colors", ["#7f7f7f"]).setOption("legend", { position: "none" })
        .setOption("vAxis", { title: "Canceled" }).setPosition(111, 13, 0, 0).build();
    summarySheet.insertChart(chart7);
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

    // --- Get Main Data ---
    const lastRowCurrent = summarySheet.getLastRow();
    let allRows = [];
    if (lastRowCurrent >= 3) {
        const rawRows = summarySheet.getRange("A3:L" + lastRowCurrent).getValues();
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
                    const readCols = oldLastCol >= 12 ? 12 : oldLastCol;
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
    const dateIdx = 0, driverIdx = 1, creditIdx = 2, cashIdx = 3, balanceIdx = 4;
    const hoursIdx = 5, tripsIdx = 7;

    allRows.forEach((row) => {
        if (!Array.isArray(row) || !row[dateIdx] || !row[driverIdx]) return;
        const date = parseDate(row[dateIdx]);
        if (!date) return;

        const driver = row[driverIdx];
        const credit = Number(row[creditIdx]) || 0;
        const cash = Number(row[cashIdx]) || 0;
        const balance = Number(row[balanceIdx]) || 0;
        const hours = Number(row[hoursIdx]) || 0;
        const trips = Number(row[tripsIdx]) || 0;

        const weekStart = getMonday(date);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekKey = `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;

        if (!weeklyData[weekKey]) weeklyData[weekKey] = {};
        if (!weeklyData[weekKey][driver]) {
            weeklyData[weekKey][driver] = { credit: 0, hours: 0, cash: 0, balance: 0, trips: 0 };
        }

        weeklyData[weekKey][driver].credit += credit;
        weeklyData[weekKey][driver].hours += hours;
        weeklyData[weekKey][driver].cash += cash;
        weeklyData[weekKey][driver].balance += balance;
        weeklyData[weekKey][driver].trips += trips;
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

            const headerRange = weeklySheet.getRange(currentRow, 1, 1, 7);
            headerRange.setValues([["Driver", "Total Credit", "Total Hours", "Total Trips", "Total Cash", "Calculated Pay", "Avg/ Hour"]])
                .setFontWeight("bold")
                .setBackground("#3c78d8")
                .setFontColor("white")
                .setHorizontalAlignment("center");

            currentRow++;

            const drivers = Object.keys(weekBlock);
            const tableData = [];

            drivers.forEach(d => {
                const info = weeklyData[weekKey][d];
                const avgCreditPerHour = info.hours > 0 ? (info.credit / info.hours) : 0;

                // STRICT NO-MATH: Pull the exact aggregated balance from the driver's sheet
                const exactSheetBalance = info.balance;
                tableData.push([d, info.credit, info.hours, info.trips, info.cash, exactSheetBalance, avgCreditPerHour]);
            });

            // Write Main Table
            const dataRange = weeklySheet.getRange(currentRow, 1, tableData.length, 7);
            dataRange.setValues(tableData);
            dataRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

            weeklySheet.getRange(currentRow, 2, tableData.length, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow, 5, tableData.length, 2).setNumberFormat("$#,##0.00");

            const totalCredit = tableData.reduce((a, b) => a + b[1], 0);
            const totalHours = tableData.reduce((a, b) => a + b[2], 0);
            const totalTrips = tableData.reduce((a, b) => a + b[3], 0);
            const totalCash = tableData.reduce((a, b) => a + b[4], 0);
            const totalBalance = tableData.reduce((a, b) => a + b[5], 0);
            const creditPerHourValues = tableData.map(r => r[6]).filter(v => v > 0);
            const totalAvgCreditPerHour = creditPerHourValues.length > 0 ? creditPerHourValues.reduce((a, b) => a + b, 0) / creditPerHourValues.length : 0;

            const totalRowRange = weeklySheet.getRange(currentRow + tableData.length, 1, 1, 7);
            totalRowRange.setValues([["TOTAL", totalCredit, totalHours, totalTrips, totalCash, totalBalance, totalAvgCreditPerHour]])
                .setFontWeight("bold")
                .setBackground("#eeeeee")
                .setBorder(true, true, true, true, true, true);

            weeklySheet.getRange(currentRow + tableData.length, 2, 1, 1).setNumberFormat("$#,##0.00");
            weeklySheet.getRange(currentRow + tableData.length, 5, 1, 2).setNumberFormat("$#,##0.00");

            const tableEnd = currentRow + tableData.length;

            // Charts
            const chart = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 9, 0, 0)
                .addRange(weeklySheet.getRange(currentRow, 1, tableData.length - 1, 2))
                .setOption("title", `WEEK ${displayedWeekCounter} - Total Credit`)
                .setOption("legend", { position: "none" }).setOption("hAxis", { title: "Driver" }).setOption("vAxis", { title: "Total Credit" })
                .build();
            weeklySheet.insertChart(chart);

            const chart2 = weeklySheet.newChart().asColumnChart()
                .setPosition(startRow - 1, 15, 0, 0)
                .addRange(weeklySheet.getRange(currentRow, 1, tableData.length - 1, 1))
                .addRange(weeklySheet.getRange(currentRow, 7, tableData.length - 1, 1))
                .setOption("colors", ["#ff7f0e"])
                .setOption("title", `WEEK ${displayedWeekCounter} - Avg Credit per Hour`)
                .setOption("legend", { position: "none" }).setOption("hAxis", { title: "Driver" }).setOption("vAxis", { title: "Credit / Hour" })
                .build();
            weeklySheet.insertChart(chart2);

            currentRow = Math.max(tableEnd + 2, startRow + 22);
        }
    });
}


// =================================================================
// ANGEL REPORT — Progressive tier payout calculation
// =================================================================
/**
 * Creates a "Real-Time Progressive" report for Angel.
 * FIXED: Added date "fill-down" memory to handle the merged date cells in the Summary sheet.
 */
function generateAngelReport(ss) {
    const summarySheet = ss.getSheetByName("Summary");
    const targetDriver = "Angel";
    const reportSheetName = "Angel Summary";

    if (!summarySheet) throw new Error('No sheet named "Summary" was found.');

    // 1. --- Detect Month and Year ---
    const currentMonthName = ss.getName().split(" ")[0];
    const lastRow = summarySheet.getLastRow();
    let currentYear = new Date().getFullYear();
    if (lastRow >= 3) {
        const firstDateVal = summarySheet.getRange("A3").getValue();
        if (firstDateVal instanceof Date) {
            currentYear = firstDateVal.getFullYear();
        }
    }

    const firstOfMonth = new Date(Date.parse(`${currentMonthName} 1, ${currentYear}`));
    const endDate = new Date(currentYear, firstOfMonth.getMonth() + 1, 0);
    const reportStartDate = getMonday(firstOfMonth);

    // 2. --- Initialize Data Map ---
    let angelDataMap = {};

    function mapRowToAngelData(row) {
        if (row[1] === targetDriver && row[0] instanceof Date) {
            let dateKey = row[0].toDateString();
            angelDataMap[dateKey] = {
                credit: Number(row[2]) || 0,
                cash: Number(row[3]) || 0,
                hours: Number(row[5]) || 0,
                trips: Number(row[7]) || 0
            };
        }
    }

    // 3. --- Fetch Previous Month Data (Crossover Week) (WITH DATE FILL-DOWN) ---
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
                        // Memory logic: remember the date if it's there, otherwise use the last seen
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
        } catch (e) { Logger.log("Error fetching previous month for Angel: " + e.toString()); }
    }

    // 4. --- Fetch Current Month Data (WITH DATE FILL-DOWN) ---
    if (lastRow >= 3) {
        const currentData = summarySheet.getRange(3, 1, lastRow - 2, 11).getValues();

        let currentLastSeenDate = null;
        currentData.forEach(row => {
            // Memory logic: remember the date if it's there, otherwise use the last seen
            if (row[0] && row[0] !== "") currentLastSeenDate = row[0];
            else row[0] = currentLastSeenDate;

            if (row[0]) {
                const d = parseDate(row[0]);
                if (d) {
                    row[0] = d;
                    mapRowToAngelData(row);
                }
            }
        });
    }

    // 5. --- Build the Report Rows (Real-Time Daily Calculation) ---
    let reportRows = [];
    let chartDataRows = [];
    let weekCounter = 1;

    // Weekly Trackers
    let weeklyCredit = 0, weeklyHours = 0, weeklyTrips = 0, weeklyCash = 0;
    let weeklyGrossBal = 0, weeklyFee = 0, weeklyNet = 0;

    // Grand Trackers
    let grandCredit = 0, grandHours = 0, grandTrips = 0, grandCash = 0;
    let grandGrossBal = 0, grandFee = 0, grandNet = 0;

    let weekTotalRows = [];

    // PROGRESSIVE TIER VARIABLES
    const threshold = 1000;
    const baseRate = 0.80;
    const topRate = 0.90;
    const transferFeeRate = 0.03; // 3%

    for (let d = new Date(reportStartDate); d <= endDate; d.setDate(d.getDate() + 1)) {
        const dateKey = d.toDateString();
        const dayData = angelDataMap[dateKey] || { credit: 0, cash: 0, hours: 0, trips: 0 };

        const todayCredit = dayData.credit;
        const todayCash = dayData.cash;
        let todayGrossPayout = 0;
        let appliedRate = 0;

        // --- PROGRESSIVE MATH ENGINE ---
        if (todayCredit > 0) {
            if (weeklyCredit >= threshold) {
                // He started the day over $1000. All of today is 90%.
                todayGrossPayout = todayCredit * topRate;
            }
            else if ((weeklyCredit + todayCredit) <= threshold) {
                // Even with today's money, still under $1000. All of today is 80%.
                todayGrossPayout = todayCredit * baseRate;
            }
            else {
                // He crossed the $1000 mark TODAY. Split the math.
                const creditAtBaseRate = threshold - weeklyCredit;
                const creditAtTopRate = todayCredit - creditAtBaseRate;
                todayGrossPayout = (creditAtBaseRate * baseRate) + (creditAtTopRate * topRate);
            }

            // Calculate the visual applied rate to show him
            appliedRate = todayGrossPayout / todayCredit;
        }

        // Update the running weekly credit AFTER doing today's math
        weeklyCredit += todayCredit;

        // --- TRANSFER FEE MATH ---
        const grossBalance = todayGrossPayout - todayCash;
        // Only apply transfer fee if there is positive money to transfer
        const fee = grossBalance > 0 ? (grossBalance * transferFeeRate) : 0;
        const netPayout = grossBalance - fee;

        // Add row to report (10 Columns now)
        reportRows.push([
            new Date(d),
            todayCredit,
            weeklyCredit,
            appliedRate, // Shows exactly what % he got for today's money
            dayData.hours,
            dayData.trips,
            todayCash,
            grossBalance,
            fee,
            netPayout
        ]);

        chartDataRows.push([new Date(d), todayCredit]);

        // Accumulate other totals
        weeklyHours += dayData.hours;
        weeklyTrips += dayData.trips;
        weeklyCash += todayCash;
        weeklyGrossBal += grossBalance;
        weeklyFee += fee;
        weeklyNet += netPayout;

        // End of Week Check
        const isSunday = d.getDay() === 0;
        const isLastDayOfReport = d.toDateString() === endDate.toDateString();

        if (isSunday || isLastDayOfReport) {

            // Add Grand Totals
            grandCredit += weeklyCredit;
            grandHours += weeklyHours;
            grandTrips += weeklyTrips;
            grandCash += weeklyCash;
            grandGrossBal += weeklyGrossBal;
            grandFee += weeklyFee;
            grandNet += weeklyNet;

            let weekLabel = `Week ${weekCounter} Total`;

            // Determine row color based on where he finished
            let rowColor = "#e6f4ea"; // Green (Standard)
            if (weeklyCredit > threshold) {
                rowColor = "#fff2cc"; // Gold (He hit the tier!)
            }

            // Add cash back to get the true gross commission payout before dividing
            let weeklyEffectiveRate = weeklyCredit > 0 ? ((weeklyGrossBal + weeklyCash) / weeklyCredit) : 0;;

            reportRows.push([
                weekLabel,
                weeklyCredit,
                "",                   // Cum. Credit stays blank
                weeklyEffectiveRate,  // The weekly rate goes here!
                weeklyHours,
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

            // RESET EVERYTHING FOR NEXT MONDAY
            weeklyCredit = 0; weeklyHours = 0; weeklyTrips = 0; weeklyCash = 0;
            weeklyGrossBal = 0; weeklyFee = 0; weeklyNet = 0;

            if (!isLastDayOfReport) {
                weekCounter++;
                reportRows.push(["", "", "", "", "", "", "", "", "", ""]); // Spacer
            }
        }
    }

    // Add Grand Total Row
    reportRows.push(["", "", "", "", "", "", "", "", "", ""]);

    // Calculate overall effective rate for the month
    const grandRate = grandCredit > 0 ? ((grandGrossBal + grandCash) / grandCredit) : 0;

    reportRows.push([
        "GRAND TOTAL", grandCredit, "", grandRate, grandHours, grandTrips, grandCash, grandGrossBal, grandFee, grandNet
    ]);
    const grandTotalRowIndex = reportRows.length + 1;

    // 6. --- Write to Sheet ---
    let reportSheet = ss.getSheetByName(reportSheetName);
    if (!reportSheet) {
        reportSheet = ss.insertSheet(reportSheetName);
    } else {
        //reportSheet.clear();
        const charts = reportSheet.getCharts();
        charts.forEach(c => reportSheet.removeChart(c));
        try { reportSheet.showColumns(1, 20); } catch (e) { }
    }

    // UPDATED HEADERS (10 Columns)
    const headers = [["Date", "Daily Credit", "Weekly Cum.", "Applied Rate", "Hours", "Trips", "Cash", "Gross Bal.", "Fee (3%)", "Net Payout"]];

    reportSheet.getRange("A1:J1").setValues(headers)
        .setFontWeight("bold")
        .setFontSize(11)
        .setHorizontalAlignment("center")
        .setBackground("#f3f3f3")
        .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

    if (reportRows.length > 0) {
        const range = reportSheet.getRange(2, 1, reportRows.length, 10);
        range.setValues(reportRows);

        // Formatting
        reportSheet.getRange(2, 1, reportRows.length, 1).setNumberFormat("dddd, dd"); // Col A
        reportSheet.getRange(2, 2, reportRows.length, 2).setNumberFormat("$#,##0.00"); // Col B, C (Credit, Cum)
        reportSheet.getRange(2, 4, reportRows.length, 1).setNumberFormat("0.0%"); // Col D (Rate)
        reportSheet.getRange(2, 5, reportRows.length, 1).setNumberFormat("0.0"); // Col E (Hours)
        reportSheet.getRange(2, 6, reportRows.length, 1).setNumberFormat("0");   // Col F (Trips)
        reportSheet.getRange(2, 7, reportRows.length, 4).setNumberFormat("$#,##0.00"); // Col G, H, I, J (Cash, Bal, Fee, Net)

        range.setHorizontalAlignment("center");

        // Styling Subtotals
        weekTotalRows.forEach(item => {
            const rowRange = reportSheet.getRange(item.index, 1, 1, 10);
            rowRange.setFontWeight("bold");
            rowRange.setBackground(item.color);
            rowRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);
            reportSheet.getRange(item.index, 1).setHorizontalAlignment("center");
        });

        // Styling Grand Total
        const grandRowRange = reportSheet.getRange(grandTotalRowIndex, 1, 1, 10);
        grandRowRange.setFontWeight("bold");
        grandRowRange.setFontSize(12);
        grandRowRange.setBackground("#d9d2e9");
        grandRowRange.setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

        // --- FIX: Dynamic Column Sizing with Padding ---
        reportSheet.autoResizeColumns(1, 10);
        // Add 15 pixels of padding to every column so headers aren't suffocating
        for (let i = 1; i <= 10; i++) {
            let currentWidth = reportSheet.getColumnWidth(i);
            reportSheet.setColumnWidth(i, currentWidth + 15);
        }
        // Make sure Date column is wide enough
        if (reportSheet.getColumnWidth(1) < 120) reportSheet.setColumnWidth(1, 120);

        // --- Write Hidden Data for Chart ---
        if (chartDataRows.length > 0) {
            const chartDataRange = reportSheet.getRange(2, 12, chartDataRows.length, 2); // Hidden data in L & M
            chartDataRange.setValues(chartDataRows);
            reportSheet.getRange(2, 12, chartDataRows.length, 1).setNumberFormat("dddd, dd");
            reportSheet.getRange(2, 13, chartDataRows.length, 1).setNumberFormat("$#,##0");
            reportSheet.hideColumns(12, 2);
        }

        // --- Add Line Chart ---
        let angelChart = reportSheet.newChart()
            .setChartType(Charts.ChartType.LINE)
            .addRange(reportSheet.getRange(2, 12, chartDataRows.length, 1)) // Date (Col L)
            .addRange(reportSheet.getRange(2, 13, chartDataRows.length, 1)) // Total Credit (Col M)
            .setPosition(2, 12, 0, 0) // Chart sits at Col N
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


// =================================================================
// BONUS REPORT — Identifies drivers who hit $1500 weekly threshold
// =================================================================
/**
 * Creates the "Bonus" sheet with professional styling.
 * FIXED: Added date "fill-down" memory to handle the merged date cells in the Summary sheet.
 */
function generateBonusReport(ss) {
    Logger.log("--- Starting generateBonusReport (Styled) ---");

    const summarySheet = ss.getSheetByName("Summary");
    const bonusSheetName = "Bonus";

    // 1. --- Get Main Data & Driver Profiles (WITH DATE FILL-DOWN) ---
    const lastRowCurrent = summarySheet.getLastRow();
    let allRows = [];
    if (lastRowCurrent >= 3) {
        const rawRows = summarySheet.getRange("A3:L" + lastRowCurrent).getValues();

        let lastSeenDate = null;
        allRows = rawRows.map(r => {
            // Memory logic: remember the date if it's there, otherwise use the last seen
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

            // Matches your Weekly script logic for Non-Dispatch
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
                    const readCols = oldLastCol >= 12 ? 12 : oldLastCol;
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

    // 4. --- Process Data & Find Qualifiers based on NEW RULE ---
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

    // =================================================================
    // --- CONFIGURATION ---
    // Set to true: Non-Dispatch (90%) drivers ARE included in the Bonus sheet if they hit $1500.
    // Set to false: Non-Dispatch drivers are completely excluded from the Bonus sheet.
    // =================================================================
    const includeNonDispatchInBonus = true;

    let allQualifiedDrivers = [];

    Object.keys(weeklyData).forEach(weekKey => {
        const weekEndDate = parseDate(weekKey.split(" - ")[1]);

        if (weekEndDate.getMonth() === currentMonthIndex && weekEndDate.getFullYear() === currentYear) {

            const weekBlock = weeklyData[weekKey];
            for (const driver in weekBlock) {

                const info = weekBlock[driver];

                // --- Handle Non-Dispatch Drivers ---
                if (nonDispatchDrivers.has(driver)) {
                    // Check if the global toggle is ON and they hit the $1500 target
                    if (includeNonDispatchInBonus === true && info.credit >= 1500) {
                        allQualifiedDrivers.push([driver, info.credit, weekKey]);
                    }
                    continue; // Skip the standard check below for all Non-Dispatch drivers
                }

                // --- Handle Standard Drivers ---
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

    // Clear everything including previous borders and colors
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

        // --- 1. Header Styling ---
        const headerRange = bonusSheet.getRange("A1:C1");
        headerRange.setValues([["Driver", "Total Credit", "Week Period"]])
            .setFontWeight("bold")
            .setFontColor("white")
            .setBackground("#3c78d8") // Dark Blue
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setFontSize(11);

        // --- 2. Data Writing and Styling ---
        const dataRange = bonusSheet.getRange(2, 1, allQualifiedDrivers.length, 3);
        dataRange.setValues(allQualifiedDrivers)
            .setHorizontalAlignment("center")
            .setVerticalAlignment("middle")
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

        // Apply Zebra Stripes (Alternating Row Colors)
        for (let i = 0; i < allQualifiedDrivers.length; i++) {
            if (i % 2 === 1) {
                bonusSheet.getRange(2 + i, 1, 1, 3).setBackground("#f3f3f3"); // Very light grey
            }
        }

        bonusSheet.getRange(2, 2, allQualifiedDrivers.length, 1).setNumberFormat("$#,##0.00");

        // --- 3. Totals Calculation and Styling ---
        const totalBonusCredit = allQualifiedDrivers.reduce((sum, row) => sum + row[1], 0);
        const halfPercentBonus = totalBonusCredit * 0.005;
        const totalRow = bonusSheet.getLastRow() + 2;

        // Define the Totals Block
        const totalsBlock = bonusSheet.getRange(totalRow, 1, 2, 2);

        // Values
        bonusSheet.getRange(totalRow, 1, 2, 1).setValues([["Grand Total"], ["0.5% Bonus"]]);
        bonusSheet.getRange(totalRow, 2, 2, 1).setValues([[totalBonusCredit], [halfPercentBonus]]);

        // Styling Totals
        totalsBlock.setFontWeight("bold")
            .setBackground("#cfe2f3") // Light Blue theme
            .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)
            .setHorizontalAlignment("center");

        bonusSheet.getRange(totalRow, 2, 2, 1).setNumberFormat("$#,##0.00");

    } else {
        bonusSheet.getRange("A1").setValue("No Bonuses for this period.").setFontWeight("bold");
    }

    // --- 4. Auto-resize columns (Final step to ensure fit) ---
    bonusSheet.autoResizeColumns(1, 3);
    // Add a little extra padding to columns so they aren't tight
    bonusSheet.setColumnWidth(1, bonusSheet.getColumnWidth(1) + 20);
    bonusSheet.setColumnWidth(2, bonusSheet.getColumnWidth(2) + 20);
    bonusSheet.setColumnWidth(3, bonusSheet.getColumnWidth(3) + 20);
}


// =================================================================
// SHARED HELPERS — Deduplicated utility functions
// =================================================================

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