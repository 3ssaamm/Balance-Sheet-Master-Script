function roundToTwo(num) {
    return Math.round((num + Number.EPSILON) * 100) / 100;
}

function parseDate(val) {
    if (!val) return null;
    if (val instanceof Date) return val;
    const d = new Date(val);
    return isNaN(d.getTime()) ? null : d;
}

// Sample weekly sheet grid
const mockWeeklyGrid = [
    ["WEEK 1 (10/01/2026 - 10/04/2026)", "", "", "", "", "", "", "", ""],
    ["Driver", "Gross Fare", "Tolls", "Trips", "Total Cash", "Credit/Debit", "Total Hours", "Avg Per Hour", "Balance"],
    ["Ablaye Diop", 140.70, 3.42, 6, 15.00, 0, 4.0, "", ""],
    ["Benjamin Douglass", 0, 0, 0, 0, 7.46, 0.0, "", ""],
    ["Total", "", "", "", "", "", "", "", ""],
    ["", "", "", "", "", "", "", "", ""],
    ["WEEK 2 (10/05/2026 - 10/11/2026)", "", "", "", "", "", "", "", ""],
    ["Driver", "Gross Fare", "Tolls", "Trips", "Total Cash", "Credit/Debit", "Total Hours", "Avg Per Hour", "Balance"],
    ["Ablaye Diop", 250.00, 5.00, 10, 20.00, 0, 7.0, "", ""],
    ["Total", "", "", "", "", "", "", "", ""]
];

const mockHoursMap = {
    "Ablaye Diop": {
        [new Date(2026, 9, 1).getTime()]: 4.5, // changed from 4.0 to 4.5
        [new Date(2026, 9, 6).getTime()]: 8.0  // changed from 7.0 to 8.0
    },
    "Benjamin Douglass": {
        [new Date(2026, 9, 1).getTime()]: 0
    }
};

let currentWeekStart = null;
let currentWeekEnd = null;
const hoursUpdates = [];

for (let i = 0; i < mockWeeklyGrid.length; i++) {
    const rowNum = 3 + i;
    const cellA = (mockWeeklyGrid[i][0] || "").toString().trim();

    if (cellA.toUpperCase().startsWith("WEEK ")) {
        const match = cellA.match(/\((.*?)\)/);
        if (match) {
            const parts = match[1].split("-").map(s => s.trim());
            if (parts.length === 2) {
                currentWeekStart = parseDate(parts[0]);
                currentWeekEnd = parseDate(parts[1]);
            }
        }
        continue;
    }

    if (cellA === "Driver" || cellA.toLowerCase().startsWith("total") || !cellA) {
        continue;
    }

    if (currentWeekStart && currentWeekEnd) {
        const driverName = cellA;
        let sumHours = 0;
        if (mockHoursMap[driverName]) {
            for (let d = new Date(currentWeekStart); d <= currentWeekEnd; d.setDate(d.getDate() + 1)) {
                const dateKey = new Date(d.getFullYear(), d.getMonth(), d.getDate()).getTime();
                sumHours += (mockHoursMap[driverName][dateKey] || 0);
            }
        }
        hoursUpdates.push({
            row: rowNum,
            driver: driverName,
            hours: roundToTwo(sumHours)
        });
    }
}

console.log("Weekly updates:", hoursUpdates);
if (hoursUpdates[0].hours !== 4.5 || hoursUpdates[1].hours !== 0 || hoursUpdates[2].hours !== 8.0) {
    throw new Error("Weekly updates failed!");
}
console.log("✅ Weekly hours in-place sync test passed!");
