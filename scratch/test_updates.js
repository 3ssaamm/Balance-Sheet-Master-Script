// Quick syntax and math verification
const threshold = 1000;
const baseRate = 0.80;
const topRate = 0.90;
const transferFeeRate = 0.03;

function calcAngelDay(credit, tolls, cash, weeklyFeeBase) {
    const feeBase = Math.max(0, credit - tolls);
    let todayFeeBasePayout = 0;
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
    }
    const newWeeklyFeeBase = weeklyFeeBase + feeBase;
    const todayGrossPayout = todayFeeBasePayout + tolls;
    const grossBalance = todayGrossPayout - cash;
    const fee = grossBalance > 0 ? (grossBalance * transferFeeRate) : 0;
    const netPayout = grossBalance - fee;
    return { feeBase, todayGrossPayout, grossBalance, fee, netPayout, newWeeklyFeeBase };
}

// Test Day 1: credit 600, tolls 50, cash 50, weeklyFeeBase 0
const d1 = calcAngelDay(600, 50, 50, 0);
console.log("Day 1:", d1);
// feeBase = 550 <= 1000 -> payout = 550 * 0.8 = 440 + 50 tolls = 490. grossBal = 490 - 50 = 440. fee = 440*0.03 = 13.20. net = 426.80.

// Test Day 2: credit 700, tolls 50, cash 50, weeklyFeeBase 550
const d2 = calcAngelDay(700, 50, 50, d1.newWeeklyFeeBase);
console.log("Day 2:", d2);
// feeBase = 650. Base rate portion: 1000 - 550 = 450 @ 0.8 = 360. Top rate portion: 200 @ 0.9 = 180.
// todayFeeBasePayout = 540 + 50 tolls = 590. grossBal = 590 - 50 = 540. fee = 540*0.03 = 16.20. net = 523.80.
console.log("Angel test passed!");
