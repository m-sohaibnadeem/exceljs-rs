'use strict';
const fs = require('node:fs');
const path = require('node:path');
const assert = require('node:assert/strict');
const root = path.resolve(__dirname, '../docs/_site');
const pages = fs.readdirSync(root).filter(file => file.endsWith('.html'));
for (const file of pages) {
  const html = fs.readFileSync(path.join(root, file), 'utf8');
  assert.ok(!/[\u2014]|&mdash;|&#8212;/.test(html), `${file}: em dash`);
  assert.match(html, /<title>[^<]+<\/title>/, `${file}: title`);
  for (const [, href] of html.matchAll(/(?:href|src)="([^"]+)"/g)) {
    if (/^(https?:|mailto:)/.test(href)) continue;
    const [target, hash] = href.split('#');
    const dest = path.resolve(root, target === './' ? 'index.html' : target || file);
    assert.ok(dest.startsWith(root + path.sep), `Link leaves site: ${href}`);
    assert.ok(fs.existsSync(dest), `${file}: missing ${href}`);
    if (hash) {
      const body = fs.readFileSync(dest, 'utf8');
      assert.ok(body.includes(`id="${hash}"`), `${file}: missing anchor ${href}`);
    }
  }
  for (const [, json] of html.matchAll(/data-chart='([^']+)'/g)) JSON.parse(json);
}
console.log(`Documentation links and metadata verified (${pages.length} pages).`);
