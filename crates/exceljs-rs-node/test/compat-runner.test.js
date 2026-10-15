'use strict';

const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');

// Run the real coordinator with a deterministic subprocess result. This
// requires neither downloaded upstream specs nor compat's Mocha dependencies.
function run(result, report) {
  let reportExists = false;
  const exits = [];
  const compat = path.resolve(__dirname, '../../../compat');
  const fakeFs = {
    existsSync(file) {
      if (file.endsWith('baseline.json')) return false;
      if (file.endsWith('mocha-report.json')) return reportExists;
      return true;
    },
    mkdirSync() {},
    unlinkSync() {reportExists = false;},
    readFileSync() {return JSON.stringify(report);},
  };
  const mockedRequire = name => {
    if (name === 'child_process') return {spawnSync() {reportExists = !!report; return result;}};
    if (name === 'fs') return fakeFs;
    if (name === 'path') return path;
    if (name === './suite') return {suiteSpecs: () => ['new-spec.js']};
    throw new Error(`Unexpected require: ${name}`);
  };
  mockedRequire.resolve = () => 'mocha.js';
  vm.runInNewContext(fs.readFileSync(path.join(compat, 'run.js'), 'utf8'), {
    require: mockedRequire, __dirname: compat, console: {log() {}, error() {}},
    process: {argv: ['node', 'run.js'], execPath: process.execPath, env: {},
      exit(code) {exits.push(code);}},
  });
  return exits[0];
}

test('compatibility runner rejects a spec crash without a baseline entry', () => {
  assert.equal(run({status: 1, stderr: 'crashed'}, null), 1);
});

test('compatibility runner rejects abnormal exit after a complete passing report', () => {
  assert.equal(run({status: 1, stderr: 'late crash'}, {passes: [{}], failures: [], pending: []}), 1);
});

test('compatibility runner accepts a successful passing report', () => {
  assert.equal(run({status: 0}, {passes: [{}], failures: [], pending: []}), 0);
});
