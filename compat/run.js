'use strict';

// Runs ExcelJS 4.4.0's own spec suite (fetched into compat/upstream by
// fetch-upstream.js) against excel-rs, or against real ExcelJS with
// EXCELRS_TARGET=exceljs.
//
// Only specs that test ExcelJS through its public API (verquire('exceljs'))
// are in the suite. Specs that test ExcelJS's internal classes (xforms,
// utils, doc/cell, Row/Column on a mock sheet, ...) have no counterpart in a
// Rust engine; suite.js lists each with the reason.
//
// Each spec file runs in its own mocha process, so a crash in one file cannot
// hide the others. Tests that never ran (the file crashed before reaching
// them) count as failures, using baseline.json (per-file test counts from
// real ExcelJS, written by `node run.js --baseline`).
//
//   node run.js                    run the suite, print a summary
//   node run.js <spec files...>    run only these (paths relative to upstream/)
//   node run.js --grep <text>      pass --grep to mocha
//   node run.js --json out.json    also write per-test results
//   node run.js --baseline         run against real ExcelJS, write baseline.json

const {spawnSync} = require('child_process');
const fs = require('fs');
const path = require('path');
const {suiteSpecs} = require('./suite');

const UPSTREAM = path.join(__dirname, 'upstream');
const BASELINE = path.join(__dirname, 'baseline.json');
if (!fs.existsSync(path.join(UPSTREAM, 'spec'))) {
  console.error('compat/upstream missing: run node fetch-upstream.js first');
  process.exit(2);
}

const args = process.argv.slice(2);
let jsonOut = null;
let writeBaseline = false;
const passthrough = [];
const files = [];
for (let i = 0; i < args.length; i++) {
  if (args[i] === '--json') jsonOut = path.resolve(args[++i]);
  else if (args[i] === '--grep') passthrough.push('--grep', args[++i]);
  else if (args[i] === '--baseline') writeBaseline = true;
  else files.push(args[i]);
}
const target = writeBaseline ? 'exceljs' : process.env.EXCELRS_TARGET || 'excel-rs';
const selected = files.length ? files : suiteSpecs();
const baseline = fs.existsSync(BASELINE) ? JSON.parse(fs.readFileSync(BASELINE, 'utf8')) : {};

const reportFile = path.join(UPSTREAM, 'spec', 'out', 'mocha-report.json');
fs.mkdirSync(path.dirname(reportFile), {recursive: true});

function runFile(file) {
  if (fs.existsSync(reportFile)) fs.unlinkSync(reportFile);
  const result = spawnSync(
    process.execPath,
    [
      require.resolve('mocha/bin/mocha.js'),
      '--require', 'spec/config/setup',
      '--require', 'spec/config/setup-unit',
      '--reporter', path.join(__dirname, 'reporter.js'),
      '--timeout', '20000',
      ...passthrough,
      file,
    ],
    {
      cwd: UPSTREAM,
      env: {
        ...process.env,
        EXCELRS_TARGET: target,
        NODE_PATH: path.join(__dirname, 'node_modules'),
        EXCELRS_REPORT: reportFile,
      },
      encoding: 'utf8',
      timeout: 300000,
    }
  );
  if (!fs.existsSync(reportFile)) {
    const why = (result.stderr || result.stdout || '').trim().split('\n').slice(-3).join(' | ');
    return {passes: [], failures: [], pending: [], crashed: `exit ${result.status}: ${why}`};
  }
  return JSON.parse(fs.readFileSync(reportFile, 'utf8'));
}

const results = {};
const totals = {passed: 0, failed: 0, notRun: 0, pending: 0};
for (const file of selected) {
  const r = runFile(file);
  const expected = baseline[file];
  const ran = r.passes.length + r.failures.length + r.pending.length;
  const notRun = expected !== undefined ? Math.max(0, expected - ran) : 0;
  results[file] = r;
  totals.passed += r.passes.length;
  totals.failed += r.failures.length;
  totals.pending += r.pending.length;
  totals.notRun += notRun;
  const bad = r.failures.length || notRun || r.crashed;
  const extra = [notRun ? `${notRun} not run` : '', r.crashed ? `CRASH ${r.crashed}` : '']
    .filter(Boolean)
    .join('; ');
  console.log(
    `${bad ? 'FAIL' : 'ok  '} ${String(r.passes.length).padStart(3)} pass ${String(r.failures.length).padStart(3)} fail  ${file}${extra ? `  (${extra})` : ''}`
  );
}

console.log(
  `\ntarget=${target}  passing=${totals.passed} failing=${totals.failed} not-run=${totals.notRun} pending=${totals.pending}`
);

if (writeBaseline) {
  const counts = {};
  for (const [file, r] of Object.entries(results)) {
    counts[file] = r.passes.length + r.failures.length + r.pending.length;
  }
  fs.writeFileSync(BASELINE, `${JSON.stringify(counts, null, 2)}\n`);
  console.log(`wrote ${path.relative(process.cwd(), BASELINE)}`);
}
if (jsonOut) {
  const simplify = t => ({title: t.fullTitle, error: t.err && t.err.message});
  const out = {};
  for (const [file, r] of Object.entries(results)) {
    out[file] = {
      crashed: r.crashed,
      passes: r.passes.map(simplify),
      failures: r.failures.map(simplify),
    };
  }
  fs.writeFileSync(jsonOut, `${JSON.stringify({totals, files: out}, null, 2)}\n`);
}
process.exit(totals.failed || totals.notRun ? 1 : 0);
