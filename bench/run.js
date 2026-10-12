'use strict';

// Times the same workloads with ExcelJS 4.4.0 (bench/node_modules: npm ci) and
// with excel-rs (crates/excel-rs-node, release build: `npm run build`).
//
//   node bench/run.js [rows] [--only name]
//
// Each workload runs in a fresh child process per library, so neither warms
// up the other's heap or JIT; the best of 3 runs is reported. For workloads
// that build or iterate a sheet in user code, the library call alone
// (writeBuffer, readFile, csv.writeFile) is also timed.

const {execFileSync} = require('child_process');
const path = require('path');
const fs = require('fs');
const os = require('os');

const ROOT = path.join(__dirname, '..');
const LIBS = {
  exceljs: path.join(__dirname, 'node_modules', 'exceljs'),
  'excel-rs': path.join(ROOT, 'crates', 'excel-rs-node'),
};

const args = process.argv.slice(2);
const ROWS = Number(args.find(a => /^\d+$/.test(a)) || 100000);
const only = args.includes('--only') ? args[args.indexOf('--only') + 1] : null;
const TMP = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-rs-bench-'));

// Workload code, run in the child with `ExcelJS`, `ROWS` and `TMP` in scope.
// Inputs are prepared outside the timed part (with ExcelJS, so both
// libraries read identical files).
const WORKLOADS = {
  'xlsx write': {
    run: `
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Data');
      ws.columns = [{header: 'Id', key: 'id'}, {header: 'Name', key: 'name'}, {header: 'Date', key: 'date', style: {numFmt: 'yyyy-mm-dd'}}, {header: 'Amount', key: 'amount', style: {numFmt: '#,##0.00'}}, {header: 'Flag', key: 'flag'}, {header: 'Formula', key: 'f'}];
      for (let i = 1; i <= ROWS; i++) {
        ws.addRow({id: i, name: 'name ' + (i % 1000), date: new Date(Date.UTC(2020, 0, 1 + (i % 365))), amount: i * 1.25, flag: i % 2 === 0, f: {formula: 'A' + (i + 1) + '*2', result: i * 2}});
      }
      MARK();
      await wb.xlsx.writeBuffer();`,
  },
  'xlsx read': {
    prepare: 'book',
    run: `
      const wb = new ExcelJS.Workbook();
      MARK();
      await wb.xlsx.readFile(TMP + '/book.xlsx');
      UNMARK();
      let n = 0;
      wb.getWorksheet('Data').eachRow(row => { row.eachCell(cell => { if (cell.value !== undefined) n++; }); });
      if (n < ROWS * 6) throw new Error('short read ' + n);`,
  },
  'stream write': {
    run: `
      const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: TMP + '/stream-out.xlsx', useSharedStrings: true, useStyles: true});
      const ws = wb.addWorksheet('Data');
      for (let i = 1; i <= ROWS; i++) {
        ws.addRow([i, 'name ' + (i % 1000), new Date(Date.UTC(2020, 0, 1 + (i % 365))), i * 1.25, i % 2 === 0]).commit();
      }
      ws.commit();
      await wb.commit();`,
  },
  'stream read': {
    prepare: 'book',
    run: `
      const reader = new ExcelJS.stream.xlsx.WorkbookReader(TMP + '/book.xlsx', {sharedStrings: 'cache', styles: 'cache', worksheets: 'emit'});
      let n = 0;
      for await (const ws of reader) { for await (const row of ws) { row.eachCell(cell => { if (cell.value !== undefined) n++; }); } }
      if (n < ROWS * 6) throw new Error('short read ' + n);`,
  },
  'csv write': {
    run: `
      const wb = new ExcelJS.Workbook();
      const ws = wb.addWorksheet('Data');
      for (let i = 1; i <= ROWS; i++) ws.addRow([i, 'name, ' + (i % 1000), i * 1.25, i % 2 === 0, 'quote "' + i + '"']);
      MARK();
      await wb.csv.writeFile(TMP + '/out.csv');`,
  },
  'csv read': {
    prepare: 'csv',
    run: `
      const wb = new ExcelJS.Workbook();
      const ws = await wb.csv.readFile(TMP + '/data.csv');
      if (ws.rowCount < ROWS) throw new Error('short read ' + ws.rowCount);`,
  },
};

const PREPARE = {
  book: `
    const wb = new ExcelJS.Workbook();
    const ws = wb.addWorksheet('Data');
    for (let i = 1; i <= ROWS; i++) {
      ws.addRow([i, 'name ' + (i % 1000), new Date(Date.UTC(2020, 0, 1 + (i % 365))), i * 1.25, i % 2 === 0, {formula: 'A' + i + '*2', result: i * 2}]);
      ws.getCell('C' + i).numFmt = 'yyyy-mm-dd';
    }
    await wb.xlsx.writeFile(TMP + '/book.xlsx');`,
  csv: `
    const lines = [];
    for (let i = 1; i <= ROWS; i++) lines.push([i, '"name, ' + (i % 1000) + '"', i * 1.25, 'text ' + i, '2020-01-' + String(1 + (i % 28)).padStart(2, '0')].join(','));
    require('fs').writeFileSync(TMP + '/data.csv', lines.join('\\n'));`,
};

function child(lib, code) {
  const script = `
    const ExcelJS = require(${JSON.stringify(LIBS[lib])});
    const ROWS = ${ROWS}; const TMP = ${JSON.stringify(TMP)};
    (async () => {
      const t0 = process.hrtime.bigint();
      let tMark = null;
      let tUnmark = null;
      const UNMARK = () => { tUnmark = process.hrtime.bigint(); };
      const MARK = () => { tMark = process.hrtime.bigint(); };
      ${code}
      const t1 = process.hrtime.bigint();
      const ms = Number(t1 - t0) / 1e6;
      // io: the library call alone (from MARK() to the end, or MARK to UNMARK)
      const io = tMark === null ? ms : Number((tUnmark || t1) - tMark) / 1e6;
      // peak resident memory (Linux), else the current
      let rss = process.memoryUsage().rss;
      try {
        rss = Number(/VmHWM:\\s+(\\d+)/.exec(require('fs').readFileSync('/proc/self/status', 'utf8'))[1]) * 1024;
      } catch (e) {}
      process.stdout.write(JSON.stringify({ms, io, rss}));
    })().catch(e => { console.error(e); process.exit(1); });`;
  return JSON.parse(execFileSync(process.execPath, ['--max-old-space-size=8192', '-e', script], {encoding: 'utf8', maxBuffer: 1 << 26}));
}

const profile = require(path.join(LIBS['excel-rs'], 'native')).buildProfile();
if (profile !== 'release') {
  console.error(`the excel-rs addon is a ${profile} build: run \`npm run build\` in crates/excel-rs-node`);
  process.exit(1);
}

const prepared = new Set();
console.log(`rows: ${ROWS}  node ${process.version}  ${os.cpus()[0].model}`);
console.log('workload        exceljs ms  excel-rs ms  speedup |  library call: exceljs ms  excel-rs ms  speedup |  exceljs peak  excel-rs peak');
for (const [name, w] of Object.entries(WORKLOADS)) {
  if (only && name !== only) continue;
  if (w.prepare && !prepared.has(w.prepare)) {
    child('exceljs', PREPARE[w.prepare]);
    prepared.add(w.prepare);
  }
  const best = lib => {
    let b = null;
    for (let i = 0; i < 3; i++) {
      const r = child(lib, w.run);
      if (!b || r.ms < b.ms) b = {...r, io: b ? Math.min(b.io, r.io) : r.io};
      else b.io = Math.min(b.io, r.io);
    }
    return b;
  };
  const e = best('exceljs');
  const p = best('excel-rs');
  const mb = x => `${(x / 1048576).toFixed(0)} MB`;
  const callCols = e.io === e.ms ? ' '.repeat(42) : `${e.io.toFixed(0).padStart(26)} ${p.io.toFixed(0).padStart(9)} ${(e.io / p.io).toFixed(2).padStart(7)}x`;
  console.log(
    `${name.padEnd(14)} ${e.ms.toFixed(0).padStart(10)} ${p.ms.toFixed(0).padStart(10)} ${(e.ms / p.ms).toFixed(2).padStart(8)}x |${callCols} | ${mb(e.rss).padStart(12)} ${mb(p.rss).padStart(9)}`
  );
}
fs.rmSync(TMP, {recursive: true, force: true});
