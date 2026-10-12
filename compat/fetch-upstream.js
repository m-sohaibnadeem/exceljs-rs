'use strict';

const fs = require('node:fs');
const path = require('node:path');
const os = require('node:os');
const {createHash} = require('node:crypto');
const {execFileSync} = require('node:child_process');
const commit = 'ac96f9a61e9799c7776bd940f05c4a51d7200209';
const checksum = '80f60b73791f8d040742b7a3170846191b1f1592408ea94f2c8c237df65c5d6e';

async function main() {
  const response = await fetch(`https://codeload.github.com/exceljs/exceljs/tar.gz/${commit}`);
  if (!response.ok) throw new Error(`Upstream download failed: ${response.status}`);
  const bytes = Buffer.from(await response.arrayBuffer());
  if (createHash('sha256').update(bytes).digest('hex') !== checksum) {
    throw new Error('Upstream archive checksum mismatch');
  }
  const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'excel-rs-compat-'));
  try {
    const archive = path.join(temp, 'upstream.tar.gz');
    fs.writeFileSync(archive, bytes);
    const out = path.join(__dirname, 'upstream');
    if (fs.existsSync(out)) throw new Error('compat/upstream already exists; remove this generated directory before fetching again');
    fs.mkdirSync(out);
    execFileSync('tar', ['-xzf', archive, '--strip-components=1', '-C', out], {stdio: 'inherit'});
    fs.copyFileSync(path.join(__dirname, 'verquire.js'), path.join(out, 'spec/utils/verquire.js'));
    fs.mkdirSync(path.join(out, 'spec/out'), {recursive: true});
    console.log(`ExcelJS 4.4.0 specs ready at ${commit}`);
  } finally {
    fs.rmSync(temp, {recursive: true, force: true});
  }
}
main().catch(error => { console.error(error); process.exitCode = 1; });
