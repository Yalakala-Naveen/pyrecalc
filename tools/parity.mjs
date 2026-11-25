// Parity check: the browser port (docs/pyrecalc.js) vs the Python package.
// Usage: node tools/parity.mjs  (expects tools/parity_expected.json, produced
// by tools/parity_expected.py). Prints PASS/FAIL per case.
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
const require = createRequire(import.meta.url);
const P = require("../docs/pyrecalc.js");

const GRID = { A1: 10, A2: 20, A3: 30, B1: 4, B2: 5, B3: 6, C1: "apple", C2: "banana", C3: "apple", D1: 2.5, D2: -3, D3: 0 };
const resolve = (sheet, ref) => (ref in GRID ? GRID[ref] : null);
const resolveRange = (sheet, a, b) => P.iterRange(a, b).map((r) => resolve(sheet, r));
const fmt = (v) => (P.isErr(v) ? v.__excelError : typeof v === "boolean" ? (v ? "TRUE" : "FALSE") : v === null ? "" : String(v));

const expected = JSON.parse(readFileSync(new URL("./parity_expected.json", import.meta.url), "utf8"));
let fail = 0;
for (const { formula, python } of expected) {
  const js = fmt(P.evalFormula(formula, resolve, resolveRange));
  const ok = js === python;
  if (!ok) fail++;
  console.log(`${ok ? "PASS" : "FAIL"}  ${formula}  python=${python} js=${js}`);
}
console.log(fail === 0 ? `PARITY OK (${expected.length}/${expected.length})` : `PARITY FAILED (${fail} mismatches)`);
process.exit(fail ? 1 : 0);
