/*
Performance benchmarking script for Range.js
 */

const XLSX = require("xlsx");
const XLSX_CALC = require("../src");
const workbook = XLSX.readFile(`${__dirname}/vlookup_large_range.xlsx`);

const times = []
const n = 100;

for (let i = 0; i < n; i++) {
    const t0 = performance.now();
    XLSX_CALC(workbook);
    const t1 = performance.now();
    const previous = t1 - t0;
    times.push(previous)
}

const average = (times.reduce((sum, t) => sum + t, 0)) / n;
console.log(`Average time for ${n} executions: ${average.toFixed(2)} ms`)
