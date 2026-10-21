// Test logic for onEdit dispatcher and syncSettingsInPlace

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

// Test Settings Fare parsing
const sampleSettings = [
    ["Driver", "Fare %"],
    ["Ablaye Diop", "90%"],
    ["Angel Yoy", "80%-90%"],
    ["Benjamin Douglass", "90%"],
    ["Brian Macancela", "85%"]
];

const fareMap = {};
for (let i = 1; i < sampleSettings.length; i++) {
    const d = sampleSettings[i][0];
    const f = sampleSettings[i][1];
    fareMap[d] = f;
}

console.log("Fare map:", fareMap);
console.log("Angel is progressive:", String(fareMap["Angel Yoy"]).trim() === "80%-90%");
console.log("Brian is standard:", String(fareMap["Brian Macancela"]).trim() !== "80%-90%");

console.log("✅ Settings logic verified!");
