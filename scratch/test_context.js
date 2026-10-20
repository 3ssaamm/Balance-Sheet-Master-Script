// Test month/year detection with "Copy of October - Drivers Daily Balance"
const ssName = "Copy of October - Drivers Daily Balance";
const monthNames = [
    "January", "February", "March", "April", "May", "June",
    "July", "August", "September", "October", "November", "December"
];
let foundMonthIndex = -1;
for (let i = 0; i < monthNames.length; i++) {
    if (new RegExp("\\b" + monthNames[i] + "\\b", "i").test(ssName)) {
        foundMonthIndex = i;
        break;
    }
}
let foundYear = -1;
const yearMatch = ssName.match(/\b(20\d\d)\b/);
if (yearMatch) foundYear = parseInt(yearMatch[1], 10);

const finalYear = foundYear > 2000 ? foundYear : 2026;
console.log("Month detected:", monthNames[foundMonthIndex], "MonthNum:", foundMonthIndex + 1);
console.log("Year detected:", finalYear);
