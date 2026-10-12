'use strict';
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '..');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const suffix = `${process.platform}-${process.arch}${process.platform === 'linux' ? '-gnu' : process.platform === 'win32' ? '-msvc' : ''}`;
const cli = process.env.npm_execpath;
if (!cli) throw new Error('Run through npm run test:package');
const temp = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-rs-package-'));
try {
  const archives = [];
  if (process.env.EXCELJS_RS_RELEASE_DIR) {
    const dir = path.resolve(process.env.EXCELJS_RS_RELEASE_DIR);
    for (const name of [`${pkg.name}-${suffix}`, pkg.name]) {
      const archive = path.join(dir, `${name}-${pkg.version}.tgz`);
      assert.ok(fs.statSync(archive).size > 0, `Missing archive: ${archive}`);
      archives.push(archive);
    }
  } else for (const cwd of [path.join(root, 'npm', suffix), root]) {
    const result = JSON.parse(execFileSync(process.execPath,
      [cli, 'pack', '--ignore-scripts', '--json', '--pack-destination', temp], {cwd, encoding: 'utf8'}))[0];
    const names = result.files.map(file => file.path);
    for (const required of ['LICENSE', 'LICENSE-EXCELJS', 'LICENSE-THIRD-PARTY', 'LICENSE-RUST-DEPENDENCIES']) assert.ok(names.includes(required));
    assert.ok(!names.some(file => /(^|\/)(test|scripts|src|artifacts|node_modules)\//.test(file)));
    if (cwd === root) {
      assert.ok(names.includes('native.js'));
      assert.ok(!names.some(file => file.endsWith('.node')), 'Root tarball must not bundle every platform');
    }
    archives.push(path.join(temp, result.filename));
  }
  fs.writeFileSync(path.join(temp, 'package.json'), '{"private":true}');
  execFileSync(process.execPath, [cli, 'install', '--ignore-scripts', '--no-audit', '--no-fund', ...archives], {cwd: temp, stdio: 'inherit'});
  execFileSync(process.execPath, ['-e', `
    const assert = require('node:assert/strict');
    const {Workbook} = require(${JSON.stringify(pkg.name)});
    (async () => {
      const workbook = new Workbook();
      const sheet = workbook.addWorksheet('Sales');
      sheet.addRow(['Notebook', 12.5]);
      sheet.getCell('B1').numFmt = '#,##0.00';
      const buffer = await workbook.xlsx.writeBuffer();
      assert.ok(Buffer.isBuffer(buffer));
      const saved = new Workbook();
      await saved.xlsx.load(buffer);
      assert.equal(saved.getWorksheet('Sales').getCell('B1').value, 12.5);
      assert.equal(saved.getWorksheet('Sales').getCell('B1').numFmt, '#,##0.00');
    })().catch(error => {console.error(error); process.exitCode = 1;});
  `], {cwd: temp, stdio: 'inherit'});
  console.log(`Installed tarball round trip passed for ${suffix} only. This does not verify other release targets.`);
} finally {
  fs.rmSync(temp, {recursive: true, force: true});
}
