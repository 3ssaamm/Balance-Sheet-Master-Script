// Scratch test for calculations
function roundToTwo(num) {
    return Math.round((num + Number.EPSILON) * 100) / 100;
}

// Test 1: User's example
// "a driver made $1050 total credit including $50 tolls and $100 cash. we will subtract them first and then calculate our fees over the $1000 which will be $100 then will subtract the $100 cash which he already collected so the driver balance will be $850."
function calcBalance(credit, tolls, cash, fareRate) {
    const feeBase = credit - tolls;
    const driverShare = feeBase * fareRate;
    return roundToTwo(driverShare + tolls - cash);
}

const ex1 = calcBalance(1050, 50, 100, 0.90);
console.log("Test 1 (User example):", ex1, "Expected: 850. Match:", ex1 === 850);

// Test 2: Screenshot Hassan En Nejjari
// Cash: 33.00, Final Pay: 355.18, Fees: -38.63, Tollways: 7.46
const hassanCash = 33.00;
const hassanFinalPay = 355.18;
const hassanFees = 38.63;
const hassanTolls = 7.46;
const hassanCredit = roundToTwo(hassanCash + hassanFinalPay + hassanFees);
console.log("Test 2 (Hassan credit):", hassanCredit, "Expected: 426.81. Match:", hassanCredit === 426.81);
const hassanBalance = calcBalance(hassanCredit, hassanTolls, hassanCash, 0.90);
console.log("Test 2 (Hassan balance @ 90%):", hassanBalance);

// Test 3: Angel 80%-90% progressive with tolls
function calcAngelDay(priorFeeBaseWeekly, dailyCredit, dailyTolls, dailyCash) {
    const feeBase = dailyCredit - dailyTolls;
    let grossPayout = 0;
    const threshold = 1000;
    if (feeBase > 0) {
        if (priorFeeBaseWeekly >= threshold) {
            grossPayout = feeBase * 0.90;
        } else if (priorFeeBaseWeekly + feeBase <= threshold) {
            grossPayout = feeBase * 0.80;
        } else {
            const at80 = threshold - priorFeeBaseWeekly;
            const at90 = feeBase - at80;
            grossPayout = (at80 * 0.80) + (at90 * 0.90);
        }
    }
    const grossBalance = roundToTwo(grossPayout + dailyTolls - dailyCash);
    const fee = grossBalance > 0 ? roundToTwo(grossBalance * 0.03) : 0;
    const netPayout = roundToTwo(grossBalance - fee);
    return {
        feeBase,
        newWeeklyFeeBase: priorFeeBaseWeekly + feeBase,
        grossPayout,
        grossBalance,
        fee,
        netPayout
    };
}

const day1 = calcAngelDay(0, 600, 50, 50); // feeBase: 550 (< 1000 -> 80% payout = 440)
// grossBalance = 440 + 50 - 50 = 440. 3% fee = 13.20. netPayout = 426.80
console.log("Test 3 (Angel Day 1):", day1);

const day2 = calcAngelDay(550, 600, 50, 50); // feeBase: 550. 450 @ 80% = 360, 100 @ 90% = 90 -> grossPayout = 450.
// grossBalance = 450 + 50 - 50 = 450. 3% fee = 13.50. netPayout = 436.50
console.log("Test 3 (Angel Day 2):", day2);
