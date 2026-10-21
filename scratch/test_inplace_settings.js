const { deepStrictEqual, strictEqual } = require('assert');

function roundToTwo(num) {
    return Math.round((num + Number.EPSILON) * 100) / 100;
}

function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
}

function parseNumber(val) {
    if (val === null || val === undefined || val === '') return 0;
    if (typeof val === 'number') return val;
    const clean = String(val).replace(/[$,]/g, '').trim();
    const num = Number(clean);
    return isNaN(num) ? 0 : num;
}

function getMonday(d) {
    const dt = new Date(d);
    const day = dt.getDay();
    const diff = dt.getDate() - day + (day === 0 ? -6 : 1);
    dt.setDate(diff);
    dt.setHours(0, 0, 0, 0);
    return dt;
}

// Mock Settings data
const mockSettings = {
    "Angel Yoy": "80%-90%",
    "Ablaye Diop": 0.90,
    "Mamadou Diallo": 0.85
};

// Mock Summary rows: Col A: Date, Col B: Driver, Col C: Gross Fare, Col D: Hours, Col E: DEPH, Col F: CPH, Col G: Trips, Col H: Cash, Col I: Tolls, Col J: Credit/Debit, Col K: Balance
const mockSummaryRows = [
    // Mon 10/05/2026
    ["10/05/2026", "Angel Yoy", 600, 8, 0, 0, 15, 0, 50, 0, 0], // feeBase = 550 (< 1000) -> 80% = 440 + 50 tolls = 490
    ["", "Ablaye Diop", 200, 6, 0, 0, 8, 20, 10, 5, 0],       // formula
    // Tue 10/06/2026 (same week)
    ["10/06/2026", "Angel Yoy", 700, 8, 0, 0, 18, 0, 50, 0, 0], // feeBase = 650. Prev = 550. Base portion = 450 @ 80% = 360. Top = 200 @ 90% = 180. Payout = 540 + 50 = 590
    ["", "Mamadou Diallo", 300, 7, 0, 0, 10, 0, 20, 0, 0]      // formula
];

// Run syncSettingsInPlace logic
const progressiveTrackers = {};
let lastSeenDate = null;
const colKResults = [];

for (let i = 0; i < mockSummaryRows.length; i++) {
    const sheetRow = 3 + i;
    const row = mockSummaryRows[i];
    const rawDate = row[0];
    if (rawDate && rawDate !== "") {
        lastSeenDate = parseDate(rawDate);
    }

    const driverName = (row[1] || "").toString().trim();
    const fare = mockSettings[driverName];
    const isProgressive = String(fare || "").trim() === "80%-90%";

    const grossFare = parseNumber(row[2]) || 0;
    const cash = parseNumber(row[7]) || 0;
    const tolls = parseNumber(row[8]) || 0;
    const creditDebit = parseNumber(row[9]) || 0;

    let targetValOrFormula = "";
    if (isProgressive) {
        const rowDate = lastSeenDate || new Date();
        const weekStart = getMonday(rowDate).getTime();
        if (!progressiveTrackers[driverName]) {
            progressiveTrackers[driverName] = { weekStart: 0, weeklyFeeBase: 0 };
        }
        if (progressiveTrackers[driverName].weekStart !== weekStart) {
            progressiveTrackers[driverName] = { weekStart: weekStart, weeklyFeeBase: 0 };
        }

        const feeBase = Math.max(0, grossFare - tolls);
        let todayGrossPayout = 0;
        if (feeBase > 0) {
            const currentWeekFeeBase = progressiveTrackers[driverName].weeklyFeeBase;
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
        progressiveTrackers[driverName].weeklyFeeBase += feeBase;
        targetValOrFormula = roundToTwo(todayGrossPayout + tolls - cash + creditDebit);
    } else {
        targetValOrFormula = `=ROUND((C${sheetRow}-I${sheetRow})*IFERROR(VLOOKUP(B${sheetRow}, Settings!$A:$B, 2, FALSE), 0.9) + I${sheetRow} - H${sheetRow} + J${sheetRow}, 2)`;
    }

    colKResults.push({
        row: sheetRow,
        driver: driverName,
        isProgressive,
        valOrFormula: targetValOrFormula
    });
}

console.log("Col K results:", colKResults);

strictEqual(colKResults[0].valOrFormula, 490);
strictEqual(colKResults[1].valOrFormula, "=ROUND((C4-I4)*IFERROR(VLOOKUP(B4, Settings!$A:$B, 2, FALSE), 0.9) + I4 - H4 + J4, 2)");
strictEqual(colKResults[2].valOrFormula, 590);
strictEqual(colKResults[3].valOrFormula, "=ROUND((C6-I6)*IFERROR(VLOOKUP(B6, Settings!$A:$B, 2, FALSE), 0.9) + I6 - H6 + J6, 2)");

// Test Weekly Summary formulas
const weeklyTestDrivers = ["Angel Yoy", "Ablaye Diop"];
const weeklyFormulas = weeklyTestDrivers.map((d, idx) => {
    const r = 5 + idx;
    const isProg = String(mockSettings[d] || "").trim() === "80%-90%";
    return isProg
        ? `=ROUND(IF((B${r}-C${r})<=1000, (B${r}-C${r})*0.8, 800 + (B${r}-C${r}-1000)*0.9) + C${r} - E${r} + F${r}, 2)`
        : `=ROUND((B${r}-C${r})*IFERROR(VLOOKUP(A${r}, Settings!$A:$B, 2, FALSE), 0.9) + C${r} - E${r} + F${r}, 2)`;
});

strictEqual(weeklyFormulas[0], "=ROUND(IF((B5-C5)<=1000, (B5-C5)*0.8, 800 + (B5-C5-1000)*0.9) + C5 - E5 + F5, 2)");
strictEqual(weeklyFormulas[1], "=ROUND((B6-C6)*IFERROR(VLOOKUP(A6, Settings!$A:$B, 2, FALSE), 0.9) + C6 - E6 + F6, 2)");

console.log("✅ All settings in-place sync tests (Summary & Weekly) passed!");
