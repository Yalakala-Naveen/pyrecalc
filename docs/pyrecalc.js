"use strict";
/* pyrecalc.js — browser port of the pyrecalc formula evaluator.
 *
 * Mirrors src/pyrecalc (parser.py + evaluator.py) for the shipped
 * function set. Excel errors are represented as { err: "#DIV/0!" }
 * style objects via the ERR() factory; check with isErr().
 * Parity with the Python package is verified by tools/parity.mjs.
 */
(function (root, factory) {
  const api = factory();
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  root.Pyrecalc = api;
})(typeof self !== "undefined" ? self : this, function () {

  const ERR = (code) => ({ __excelError: code });
  const isErr = (v) => v !== null && typeof v === "object" && "__excelError" in v;
  const DIV0 = ERR("#DIV/0!"), NA = ERR("#N/A"), NAME = ERR("#NAME?"),
        REF = ERR("#REF!"), VALUE = ERR("#VALUE!"), NUM = ERR("#NUM!");

  // ---- tokenizer / parser (mirrors parser.py) -------------------------
  const TOKEN_RE = /\s*(?:(\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?|("(?:[^"]|"")*")|('(?:[^']|'')*')|(\$?[A-Za-z]{1,3}\$?\d+)|(<>|<=|>=|[+\-*/^&=<>(),:!%])|([A-Za-z_\\][A-Za-z0-9_.]*))/y;
  const CELL_RE = /^\$?[A-Za-z]{1,3}\$?\d+$/;

  function tokenize(src) {
    const toks = []; let pos = 0;
    while (pos < src.length) {
      TOKEN_RE.lastIndex = pos;
      const m = TOKEN_RE.exec(src);
      if (!m || m.index !== pos) throw new Error("cannot tokenize at " + JSON.stringify(src.slice(pos, pos + 10)));
      pos = TOKEN_RE.lastIndex;
      const kinds = ["number", "string", "quoted_sheet", "cell", "op", "ident"];
      for (let k = 0; k < kinds.length; k++) if (m[k + 1] !== undefined) { toks.push({ kind: kinds[k], text: m[k + 1] }); break; }
    }
    toks.push({ kind: "eof", text: "" });
    return toks;
  }
  const normRef = (t) => t.replace(/\$/g, "").toUpperCase();

  function parse(src) {
    const toks = tokenize(src.trim()); let i = 0;
    const peek = (a = 0) => toks[Math.min(i + a, toks.length - 1)];
    const next = () => toks[i++];
    const isOp = (t, s) => t.kind === "op" && t.text === s;
    function expectOp(s) { const t = next(); if (!isOp(t, s)) throw new Error("expected " + s + ", got " + t.text); }
    function parseExpr() { return parseComparison(); }
    function parseComparison() {
      let node = parseConcat();
      while (peek().kind === "op" && ["=", "<>", "<", "<=", ">", ">="].includes(peek().text)) { const op = next().text; node = ["binop", op, node, parseConcat()]; }
      return node;
    }
    function parseConcat() { let n = parseAddSub(); while (isOp(peek(), "&")) { next(); n = ["binop", "&", n, parseAddSub()]; } return n; }
    function parseAddSub() { let n = parseMulDiv(); while (peek().kind === "op" && ["+", "-"].includes(peek().text)) { const op = next().text; n = ["binop", op, n, parseMulDiv()]; } return n; }
    function parseMulDiv() { let n = parsePower(); while (peek().kind === "op" && ["*", "/"].includes(peek().text)) { const op = next().text; n = ["binop", op, n, parsePower()]; } return n; }
    function parsePower() { let n = parseUnary(); while (isOp(peek(), "^")) { next(); n = ["binop", "^", n, parseUnary()]; } return n; }
    function parseUnary() { const t = peek(); if (t.kind === "op" && ["+", "-"].includes(t.text)) { next(); return ["unary", t.text, parseUnary()]; } return parsePostfix(); }
    function parsePostfix() { let n = parsePrimary(); while (isOp(peek(), "%")) { next(); n = ["percent", n]; } return n; }
    function maybeRange(sheet, start) {
      if (isOp(peek(), ":")) { next(); const e = next(); if (!["ident", "cell"].includes(e.kind) || !CELL_RE.test(e.text)) throw new Error("bad range end"); return ["range", sheet, start, normRef(e.text)]; }
      return ["ref", sheet, start];
    }
    function parseIdentLike(tok) {
      if (isOp(peek(), "!")) {
        next();
        const sheet = tok.kind === "quoted_sheet" ? tok.text.slice(1, -1).replace(/''/g, "'") : tok.text;
        const rt = next();
        if (!["ident", "cell"].includes(rt.kind) || !CELL_RE.test(rt.text)) throw new Error("bad ref after !");
        return maybeRange(sheet, normRef(rt.text));
      }
      if (tok.kind === "quoted_sheet") throw new Error("quoted sheet must be followed by !");
      const upper = tok.text.toUpperCase();
      if (upper === "TRUE") return ["bool", true];
      if (upper === "FALSE") return ["bool", false];
      if (isOp(peek(), "(")) {
        next(); const args = [];
        if (!isOp(peek(), ")")) { args.push(parseExpr()); while (isOp(peek(), ",")) { next(); args.push(parseExpr()); } }
        expectOp(")"); return ["call", upper, args];
      }
      if (CELL_RE.test(tok.text)) return maybeRange(null, normRef(tok.text));
      return ["name", upper];
    }
    function parsePrimary() {
      const tok = next();
      if (tok.kind === "number") return ["num", parseFloat(tok.text)];
      if (tok.kind === "string") return ["str", tok.text.slice(1, -1).replace(/""/g, '"')];
      if (isOp(tok, "(")) { const n = parseExpr(); expectOp(")"); return n; }
      if (["ident", "quoted_sheet", "cell"].includes(tok.kind)) return parseIdentLike(tok);
      throw new Error("unexpected token " + tok.text);
    }
    const node = parseExpr();
    if (peek().kind !== "eof") throw new Error("trailing input at " + peek().text);
    return node;
  }

  // ---- evaluator (mirrors evaluator.py) -------------------------------
  const isNum = (v) => typeof v === "number";
  function toNum(v) {
    if (isErr(v)) return v;
    if (v === null || v === undefined) return 0;
    if (typeof v === "boolean") return v ? 1 : 0;
    if (isNum(v)) return v;
    if (typeof v === "string") { const n = parseFloat(v.trim()); return v.trim() !== "" && !Number.isNaN(n) && String(n) === v.trim().replace(/^\+/, "") || (!Number.isNaN(n) && v.trim() !== "") ? (Number.isNaN(n) ? VALUE : n) : VALUE; }
    return VALUE;
  }
  function toText(v) {
    if (isErr(v)) return v;
    if (v === null || v === undefined) return "";
    if (typeof v === "boolean") return v ? "TRUE" : FALSE0();
    if (isNum(v)) return Number.isInteger(v) ? String(v) : String(v);
    return String(v);
  }
  function FALSE0() { return "FALSE"; }

  function colToNum(col) { let n = 0; for (const ch of col) n = n * 26 + (ch.charCodeAt(0) - 64); return n; }
  function numToCol(n) { let out = ""; while (n) { const r = (n - 1) % 26; n = Math.floor((n - 1) / 26); out = String.fromCharCode(65 + r) + out; } return out; }
  function splitRef(ref) { const m = /^([A-Z]+)(\d+)$/.exec(ref); return [colToNum(m[1]), parseInt(m[2], 10)]; }
  function iterRange(a, b) {
    const [c1, r1] = splitRef(a), [c2, r2] = splitRef(b); const out = [];
    for (let r = Math.min(r1, r2); r <= Math.max(r1, r2); r++) for (let c = Math.min(c1, c2); c <= Math.max(c1, c2); c++) out.push(numToCol(c) + r);
    return out;
  }

  function criteriaMatch(value, criteria) {
    if (isErr(criteria)) return false;
    if (typeof criteria === "string") {
      const m = /^(>=|<=|<>|>|<|=)?(.*)$/.exec(criteria);
      const op = m[1] || "=", rest = m[2];
      const numCrit = rest !== "" && !Number.isNaN(Number(rest)) ? Number(rest) : null;
      if (numCrit !== null && isNum(value)) {
        return { "=": value === numCrit, "<>": value !== numCrit, ">": value > numCrit, "<": value < numCrit, ">=": value >= numCrit, "<=": value <= numCrit }[op];
      }
      const a = value === null || value === undefined ? "" : String(value).toLowerCase();
      const b = rest.toLowerCase();
      return { "=": a === b, "<>": a !== b, ">": a > b, "<": a < b, ">=": a >= b, "<=": a <= b }[op];
    }
    if (isNum(criteria)) return isNum(value) && value === criteria;
    if (typeof criteria === "boolean") return value === criteria;
    return false;
  }

  function flatten(args) { const out = []; for (const a of args) { if (Array.isArray(a)) out.push(...a); else out.push(a); } return out; }

  function evaluate(node, resolve, resolveRange) {
    const kind = node[0];
    if (kind === "num" || kind === "str" || kind === "bool") return node[1];
    if (kind === "name") return NAME;
    if (kind === "ref") return resolve(node[1], node[2]);
    if (kind === "range") return resolveRange(node[1], node[2], node[3]);
    if (kind === "percent") { const v = toNum(evaluate(node[1], resolve, resolveRange)); return isErr(v) ? v : v / 100; }
    if (kind === "unary") { const v = toNum(evaluate(node[2], resolve, resolveRange)); if (isErr(v)) return v; return node[1] === "-" ? -v : v; }
    if (kind === "binop") return evalBinop(node[1], node[2], node[3], resolve, resolveRange);
    if (kind === "call") return evalCall(node[1], node[2], resolve, resolveRange);
    return VALUE;
  }

  function evalBinop(op, ln, rn, resolve, resolveRange) {
    const left = evaluate(ln, resolve, resolveRange), right = evaluate(rn, resolve, resolveRange);
    if (isErr(left)) return left; if (isErr(right)) return right;
    if (op === "&") { const a = toText(left), b = toText(right); return isErr(a) ? a : isErr(b) ? b : a + b; }
    if (["+", "-", "*", "/", "^"].includes(op)) {
      const a = toNum(left), b = toNum(right);
      if (isErr(a)) return a; if (isErr(b)) return b;
      if (op === "+") return a + b; if (op === "-") return a - b; if (op === "*") return a * b;
      if (op === "/") return b === 0 ? DIV0 : a / b;
      const p = Math.pow(a, b); return Number.isNaN(p) ? NUM : p;
    }
    // comparisons
    let a, b;
    if (isNum(left) && isNum(right)) { a = left; b = right; }
    else if (typeof left === "string" && typeof right === "string") { a = left.toLowerCase(); b = right.toLowerCase(); }
    else if (typeof left === "boolean" && typeof right === "boolean") { a = left; b = right; }
    else if ((left === null || left === undefined) && isNum(right)) { a = 0; b = right; }
    else if ((right === null || right === undefined) && isNum(left)) { a = left; b = 0; }
    else return VALUE;
    return { "=": a === b, "<>": a !== b, "<": a < b, "<=": a <= b, ">": a > b, ">=": a >= b }[op];
  }

  function evalCall(name, argNodes, resolve, resolveRange) {
    const ev = (n) => evaluate(n, resolve, resolveRange);
    if (name === "IF") {
      if (![2, 3].includes(argNodes.length)) return VALUE;
      const cond = ev(argNodes[0]); if (isErr(cond)) return cond;
      const truthy = typeof cond === "boolean" ? cond : (isNum(toNum(cond)) ? toNum(cond) !== 0 : false);
      return truthy ? ev(argNodes[1]) : (argNodes.length === 3 ? ev(argNodes[2]) : false);
    }
    if (name === "IFERROR") { if (argNodes.length !== 2) return VALUE; const v = ev(argNodes[0]); return isErr(v) ? ev(argNodes[1]) : v; }
    if (["SUMIF", "COUNTIF", "AVERAGEIF"].includes(name)) {
      if (![2, 3].includes(argNodes.length)) return VALUE;
      const critRange = ev(argNodes[0]), criteria = ev(argNodes[1]);
      if (isErr(criteria)) return criteria;
      const critVals = Array.isArray(critRange) ? critRange : [critRange];
      for (const v of critVals) if (isErr(v)) return v;
      if (name === "COUNTIF") { if (argNodes.length !== 2) return VALUE; return critVals.filter((v) => criteriaMatch(v, criteria)).length; }
      const sumRange = argNodes.length === 3 ? ev(argNodes[2]) : critRange;
      if (isErr(sumRange)) return sumRange;
      const sumVals = Array.isArray(sumRange) ? sumRange : [sumRange];
      for (const v of sumVals) if (isErr(v)) return v;
      const matched = [];
      critVals.forEach((cv, idx) => { const sv = sumVals[idx]; if (criteriaMatch(cv, criteria) && isNum(sv)) matched.push(sv); });
      if (name === "SUMIF") return matched.reduce((s, x) => s + x, 0);
      return matched.length ? matched.reduce((s, x) => s + x, 0) / matched.length : DIV0;
    }
    if (name === "SUMPRODUCT") {
      const arrays = [];
      for (const a of argNodes) { const v = ev(a); const vals = Array.isArray(v) ? v : [v]; for (const it of vals) if (isErr(it)) return it; arrays.push(vals.map((x) => (isNum(x) ? x : 0))); }
      if (!arrays.length || new Set(arrays.map((a) => a.length)).size !== 1) return VALUE;
      let total = 0; for (let r = 0; r < arrays[0].length; r++) { let p = 1; for (const a of arrays) p *= a[r]; total += p; }
      return total;
    }
    if (["AND", "OR", "NOT"].includes(name)) {
      const vals = argNodes.map(ev); const flat = flatten(vals);
      for (const v of flat) if (isErr(v)) return v;
      const bools = flat.filter((v) => v !== null && v !== undefined).map((v) => typeof v === "boolean" ? v : isNum(v) ? v !== 0 : null);
      if (bools.some((b) => b === null)) return VALUE;
      if (name === "AND") return bools.every(Boolean);
      if (name === "OR") return bools.some(Boolean);
      return bools.length === 1 ? !bools[0] : VALUE;
    }
    const args = argNodes.map(ev); const flat = flatten(args);
    for (const v of flat) if (isErr(v)) return v;
    const nums = flat.filter(isNum);
    switch (name) {
      case "SUM": return nums.reduce((s, x) => s + x, 0);
      case "AVERAGE": return nums.length ? nums.reduce((s, x) => s + x, 0) / nums.length : DIV0;
      case "MIN": return nums.length ? Math.min(...nums) : 0;
      case "MAX": return nums.length ? Math.max(...nums) : 0;
      case "COUNT": return nums.length;
      case "COUNTA": return flat.filter((v) => v !== null && v !== undefined && v !== "").length;
      case "ROUND": case "ROUNDUP": case "ROUNDDOWN": {
        if (args.length !== 2) return VALUE;
        const x = toNum(args[0]), d = toNum(args[1]); if (isErr(x) || isErr(d)) return isErr(x) ? x : d;
        const f = Math.pow(10, d);
        if (name === "ROUND") return Math.sign(x) * Math.round(Math.abs(x) * f) / f;
        if (name === "ROUNDUP") return Math.sign(x) * Math.ceil(Math.abs(x) * f) / f;
        return Math.sign(x) * Math.floor(Math.abs(x) * f) / f;
      }
      case "INT": { const x = toNum(args[0]); return isErr(x) ? x : Math.floor(x); }
      case "MOD": { const a = toNum(args[0]), b = toNum(args[1]); if (isErr(a) || isErr(b)) return VALUE; return b === 0 ? DIV0 : a - b * Math.floor(a / b); }
      case "ABS": { const x = toNum(args[0]); return isErr(x) ? x : Math.abs(x); }
      case "SQRT": { const x = toNum(args[0]); return isErr(x) ? x : x < 0 ? NUM : Math.sqrt(x); }
      case "POWER": { const a = toNum(args[0]), b = toNum(args[1]); if (isErr(a) || isErr(b)) return VALUE; const p = Math.pow(a, b); return Number.isNaN(p) ? NUM : p; }
      case "LEN": { const t = toText(args[0]); return isErr(t) ? t : t.length; }
      case "LEFT": case "RIGHT": {
        const t = toText(args[0]); if (isErr(t)) return t;
        const n = args.length > 1 ? toNum(args[1]) : 1; if (isErr(n)) return n;
        return name === "LEFT" ? t.slice(0, n) : t.slice(Math.max(0, t.length - n));
      }
      case "MID": { const t = toText(args[0]); if (isErr(t)) return t; const s = toNum(args[1]), n = toNum(args[2]); if (isErr(s) || isErr(n)) return VALUE; return t.slice(Math.max(0, s - 1), Math.max(0, s - 1) + n); }
      case "TRIM": { const t = toText(args[0]); return isErr(t) ? t : t.trim().replace(/ +/g, " "); }
      case "PROPER": { const t = toText(args[0]); return isErr(t) ? t : t.toLowerCase().replace(/(^|[^A-Za-z])([a-z])/g, (m, p, c) => p + c.toUpperCase()); }
      case "NA": return NA;
      default: return NAME;
    }
  }

  function evalFormula(formula, resolve, resolveRange) {
    try { return evaluate(parse(formula), resolve, resolveRange); }
    catch (e) { return NAME; }
  }

  return { parse, evaluate, evalFormula, iterRange, isErr, ERR };
});
