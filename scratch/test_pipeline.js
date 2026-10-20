function roundToTwo(num) {
    return Math.round((num + Number.EPSILON) * 100) / 100;
}

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
    else if (lower.includes("credit")) colMap.credits = c;
    else if (lower.includes("debit")) colMap.debits = c;
    else if (lower.includes("final pay") || lower === "pay") colMap.finalPay = c;
    else if (lower.includes("cash")) colMap.cash = c;
});

console.log("Column Mapping:", colMap);

// Test 1: Hassan En Nejjari
const hassanRow = ["104111", "Hassan En Nejjari", 15, 386.35, -38.63, 7.46, 0.00, 0.00, 355.18, 33.00];
const hassanCash = hassanRow[colMap.cash];
const hassanFinalPay = hassanRow[colMap.finalPay];
const hassanFees = hassanRow[colMap.fees];
const hassanTolls = hassanRow[colMap.tolls];
const hassanRawCredits = hassanRow[colMap.credits];
const hassanRawDebits = hassanRow[colMap.debits];

const hassanCD = roundToTwo(hassanRawCredits - Math.abs(hassanRawDebits));
const hassanTotalFromFile = roundToTwo(hassanCash + hassanFinalPay + Math.abs(hassanFees));
const hassanGrossFare = roundToTwo(hassanTotalFromFile - hassanCD);
const hassanBalance = roundToTwo((hassanGrossFare - hassanTolls) * 0.9 + hassanTolls - hassanCash + hassanCD);

console.log("Hassan: Gross Fare:", hassanGrossFare, "Credit/Debit:", hassanCD, "Tolls:", hassanTolls, "Balance:", hassanBalance);

// Test 2: Benjamin Douglass (Oct 03 - 0 trips, 0 cash, 0 tolls, 0 fees, 7.46 credit adjustment, 7.46 final pay)
const benRow = ["109999", "Benjamin Douglass", 0, 0.00, 0.00, 0.00, 7.46, 0.00, 7.46, 0.00];
const benCash = benRow[colMap.cash];
const benFinalPay = benRow[colMap.finalPay];
const benFees = benRow[colMap.fees];
const benTolls = benRow[colMap.tolls];
const benRawCredits = benRow[colMap.credits];
const benRawDebits = benRow[colMap.debits];

const benCD = roundToTwo(benRawCredits - Math.abs(benRawDebits));
const benTotalFromFile = roundToTwo(benCash + benFinalPay + Math.abs(benFees));
const benGrossFare = roundToTwo(benTotalFromFile - benCD);
const benBalance = roundToTwo((benGrossFare - benTolls) * 0.9 + benTolls - benCash + benCD);

console.log("Benjamin: Gross Fare:", benGrossFare, "Credit/Debit:", benCD, "Balance:", benBalance);
if (benGrossFare !== 0 || benCD !== 7.46 || benBalance !== 7.46) {
    throw new Error("Benjamin Douglass test failed!");
}
console.log("✅ All pipeline checks passed successfully!");
