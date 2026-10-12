'use strict';

// Regenerate notices from the exact Cargo.lock dependencies. MIT is selected
// where offered as an alternative; additional mandatory licenses are retained.
const fs = require('node:fs');
const path = require('node:path');
const {execFileSync} = require('node:child_process');
const root = path.resolve(__dirname, '..');

async function main() {
  const metadata = JSON.parse(execFileSync('cargo', ['metadata', '--locked', '--format-version', '1'],
    {cwd: root, encoding: 'utf8', maxBuffer: 16 * 1024 * 1024}));
  const sections = ['Third-party Rust dependency notices\n\nGenerated from Cargo.lock by scripts/licenses.js.\nBuild-time dependencies are included as well as runtime dependencies.\n'];
  for (const pkg of metadata.packages.filter(pkg => pkg.source).sort((a, b) => a.name.localeCompare(b.name, 'en'))) {
    const dir = path.dirname(pkg.manifest_path);
    const files = fs.readdirSync(dir).filter(file => /^(license|copying|notices?)([.-]|$)/i.test(file)
      && fs.statSync(path.join(dir, file)).isFile()
      && !(pkg.license.includes('MIT') && /apache|^copying$/i.test(file)));
    let notices = files.map(file => `${file}\n\n${fs.readFileSync(path.join(dir, file), 'utf8').trim()}`);
    if (!notices.length && pkg.repository === 'https://github.com/napi-rs/napi-rs') {
      const vcs = JSON.parse(fs.readFileSync(path.join(dir, '.cargo_vcs_info.json'), 'utf8'));
      const url = `https://raw.githubusercontent.com/napi-rs/napi-rs/${vcs.git.sha1}/LICENSE`;
      const response = await fetch(url);
      if (!response.ok) throw new Error(`Missing upstream license for ${pkg.name}: ${response.status}`);
      notices = [`${url}\n\n${(await response.text()).trim()}`];
    }
    if (!notices.length) throw new Error(`No license text found for ${pkg.name}`);
    sections.push(`================================================================\n${pkg.name} ${pkg.version}\nDeclared license: ${pkg.license}\nSource: https://crates.io/crates/${pkg.name}/${pkg.version}\n================================================================\n\n${notices.join('\n\n')}`);
  }
  // Rust standard library code is statically linked into native binaries.
  const rust = execFileSync('rustc', ['-vV'], {encoding: 'utf8'});
  const commit = /^commit-hash: (\w+)$/m.exec(rust)?.[1];
  if (!commit) throw new Error('Cannot determine Rust compiler revision');
  const url = `https://raw.githubusercontent.com/rust-lang/rust/${commit}/LICENSE-MIT`;
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Cannot fetch Rust license: ${response.status}`);
  sections.push(`================================================================\nRust standard library\n${url}\n================================================================\n\n${(await response.text()).trim()}`);
  fs.writeFileSync(path.join(root, 'LICENSE-RUST-DEPENDENCIES'), sections.join('\n\n') + '\n');
  console.log(`Generated notices for ${metadata.packages.filter(pkg => pkg.source).length} Rust dependencies and the standard library.`);
}
main().catch(error => {console.error(error); process.exitCode = 1;});
