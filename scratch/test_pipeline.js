// Test the end-to-end extraction and formula generation logic
const fs = require('fs');

function roundToTwo(num) {
    return Math.round((num + Number.EPSILON) * 100) / 100;
}

// Test parsing of period lines
function parsePeriodLine(text) {
    if (!text) return null;
    const clean = String(text).replace(/[\u2010-\u2015\u2212]/g, '-').trim();
    const match = clean.match(/period:\s*([0-9\/\.\-]+)\s*[\-–—to]+\s*([0-9\/\.\-]+)/i);
    if (match) {
        return { startStr: match[1], endStr: match[2] };
    }
    return null;
}

console.log("Period match test 1:", parsePeriodLine("Period: 10/1/2026 – 10/1/2026"));
console.log("Period match test 2 (multi-day):", parsePeriodLine("Period: 10/1/2026 - 10/7/2026"));
console.log("Period match test 3 (hyphen):", parsePeriodLine("Period: 10-01-2026 - 10-01-2026"));

// Test column mapping from sample row
const sampleHeaders = ["Driver ID", "Name", "Trips", "Driver Net", "Fees", "Tollways", "Credits", "Debits", "Final Pay", "Cash"];
const colMap = {};
sampleHeaders.forEach((h, c) => {
    const lower = h.toLowerCase().trim();
    if (lower === "name" || lower.includes("driver name") || (lower === "driver" && !lower.includes("id"))) colMap.name = c;
    else if (lower.includes("driver id") || lower === "id") colMap.driverId = c;
    else if (lower === "trips" || lower === "trip" || lower.includes("trips")) colMap.trips = c;
    else if (lower.includes("driver net") || lower === "net") colMap.driverNet = c;
    else if (lower.includes("fee")) colMap.fees = c;
    else if (lower.includes("toll")) colMap.tolls = c;
    else if (lower.includes("final pay") || lower === "pay") colMap.finalPay = c;
    else if (lower.includes("cash")) colMap.cash = c;
});

console.log("Column Mapping:", colMap);
// Verify Hassan row:
const hassanRow = ["104111", "Hassan En Nejjari", 15, 386.35, -38.63, 7.46, 0.00, 0.00, 355.18, 33.00];
const cash = hassanRow[colMap.cash];
const finalPay = hassanRow[colMap.finalPay];
const fees = hassanRow[colMap.fees];
const tolls = hassanRow[colMap.tolls];
const credit = roundToTwo(cash + finalPay + Math.abs(fees));
console.log("Hassan computed credit:", credit, "Expected: 426.81. Tolls:", tolls);
