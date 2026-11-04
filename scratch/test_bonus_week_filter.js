const assert = require('assert');

function getMonday(d) {
    const dt = new Date(d);
    const day = dt.getDay();
    const diff = dt.getDate() - day + (day === 0 ? -6 : 1);
    dt.setDate(diff);
    dt.setHours(0, 0, 0, 0);
    return dt;
}

function formatDate(date) {
    const m = String(date.getMonth() + 1).padStart(2, '0');
    const d = String(date.getDate()).padStart(2, '0');
    const y = date.getFullYear();
    return `${m}/${d}/${y}`;
}

// Test weeks
const weeks = [
    new Date(2026, 8, 28), // 09/28/2026 (Mon) -> ends 10/04/2026 (Sun)
    new Date(2026, 9, 5),  // 10/05/2026 (Mon) -> ends 10/11/2026 (Sun)
    new Date(2026, 9, 12), // 10/12/2026 (Mon) -> ends 10/18/2026 (Sun)
    new Date(2026, 9, 19), // 10/19/2026 (Mon) -> ends 10/25/2026 (Sun)
    new Date(2026, 9, 26), // 10/26/2026 (Mon) -> ends 11/01/2026 (Sun)
    new Date(2026, 10, 2), // 11/02/2026 (Mon) -> ends 11/08/2026 (Sun)
];

function filterWeeksForMonth(currentMonthIndex, currentYear) {
    const included = [];
    weeks.forEach(startDate => {
        const weekStart = getMonday(startDate);
        const weekEnd = new Date(weekStart);
        weekEnd.setDate(weekEnd.getDate() + 6);
        const weekKey = `${formatDate(weekStart)} - ${formatDate(weekEnd)}`;

        const isCurrentMonthWeek = weekEnd &&
            weekEnd.getMonth() === currentMonthIndex &&
            weekEnd.getFullYear() === currentYear;

        if (isCurrentMonthWeek) {
            included.push(weekKey);
        }
    });
    return included;
}

// Test October (monthIndex = 9)
const octWeeks = filterWeeksForMonth(9, 2026);
console.log("October bonus weeks:", octWeeks);

// Test November (monthIndex = 10)
const novWeeks = filterWeeksForMonth(10, 2026);
console.log("November bonus weeks:", novWeeks);

// Assertions:
// 10/26/2026 - 11/01/2026 MUST NOT be in October!
assert.strictEqual(octWeeks.includes("10/26/2026 - 11/01/2026"), false, "10/26 - 11/01 should NOT be in October");
// 10/26/2026 - 11/01/2026 MUST be in November!
assert.strictEqual(novWeeks.includes("10/26/2026 - 11/01/2026"), true, "10/26 - 11/01 MUST be in November");

// 09/28/2026 - 10/04/2026 ends in October, so it IS in October:
assert.strictEqual(octWeeks.includes("09/28/2026 - 10/04/2026"), true, "09/28 - 10/04 ends in Oct, should be in Oct");

console.log("✅ All bonus week filtering tests passed successfully!");

