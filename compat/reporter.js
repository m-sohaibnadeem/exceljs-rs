'use strict';

// Mocha reporter: prints dots like the "dot" reporter and writes a JSON
// summary to $EXCELRS_REPORT (mocha 7's json reporter only writes to stdout).

const fs = require('fs');
const Mocha = require('mocha');

const {EVENT_TEST_PASS, EVENT_TEST_FAIL, EVENT_TEST_PENDING, EVENT_RUN_END} = Mocha.Runner.constants;

class ExcelRsReporter extends Mocha.reporters.Base {
  constructor(runner) {
  super(runner);
  const out = {passes: [], failures: [], pending: []};
  const entry = (t, err) => ({
    file: t.file,
    fullTitle: t.fullTitle(),
    err: err ? {message: String(err.message), stack: String(err.stack || '')} : undefined,
  });
  runner.on(EVENT_TEST_PASS, t => out.passes.push(entry(t)));
  runner.on(EVENT_TEST_FAIL, (t, err) => out.failures.push(entry(t, err)));
  runner.on(EVENT_TEST_PENDING, t => out.pending.push(entry(t)));
  runner.once(EVENT_RUN_END, () => {
    out.stats = {
      tests: out.passes.length + out.failures.length + out.pending.length,
      passes: out.passes.length,
      failures: out.failures.length,
      pending: out.pending.length,
    };
    fs.writeFileSync(process.env.EXCELRS_REPORT, JSON.stringify(out));
  });
  }
}

module.exports = ExcelRsReporter;
