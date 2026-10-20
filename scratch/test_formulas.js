// Verify formula strings syntax
const r = 5;
const deph = `=IF(D${r}>0, ROUND((H${r}+K${r})/D${r}, 2), 0)`;
const cph = `=IF(D${r}>0, ROUND(C${r}/D${r}, 2), 0)`;
const bal = `=ROUND((C${r}-I${r})*IFERROR(VLOOKUP(B${r}, Settings!$A:$B, 2, FALSE), 0.9) + I${r} - H${r}, 2)`;
console.log("DEPH formula:", deph);
console.log("CPH formula:", cph);
console.log("Balance formula:", bal);

const aggRow = 4;
const creditFormula = `=SUMIF($B$3:$B, M${aggRow}, $C$3:$C)`;
const netFormula = `=ROUND(N${aggRow} - P${aggRow}, 2)`;
console.log("Agg Credit formula:", creditFormula);
console.log("Agg Net formula:", netFormula);
