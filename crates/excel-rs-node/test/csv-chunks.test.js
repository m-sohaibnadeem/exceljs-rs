'use strict';

// workbook.csv.read parses its input chunk by chunk as it arrives (as
// fast-csv does). Any split of the input must give the rows of parsing it
// at once.

const test = require('node:test');
const assert = require('node:assert/strict');
const {createParser} = require('../lib/csv/fast-csv');

// The rows, or that it failed (the message quotes the text around the
// error, which depends on the chunks, as in fast-csv).
function rows(chunks, opts) {
  const out = [];
  try {
    const p = createParser(opts, row => out.push(row));
    chunks.forEach(c => p.write(c));
    p.end();
  } catch (e) {
    return {failed: /^Parse Error/.test(e.message)};
  }
  return out;
}

function random(seed) {
  let s = seed;
  return () => {
    s = (s * 1103515245 + 12345) % 2147483648;
    return s / 2147483648;
  };
}

test('chunked parse equals whole parse', () => {
  const pieces = ['a', 'é', '日本', '"q,uo""ted"', ',', '\n', '\r\n', ' ', '😀', '1.5', '""'];
  for (let seed = 1; seed <= 200; seed++) {
    const r = random(seed);
    let text = '';
    const n = 5 + Math.floor(r() * 60);
    for (let i = 0; i < n; i++) text += pieces[Math.floor(r() * pieces.length)];
    for (const opts of [{}, {headers: false, delimiter: ';'}]) {
      const whole = rows([Buffer.from(text)], opts);
      // split the bytes anywhere, also inside a character
      const bytes = Buffer.from(text);
      const chunks = [];
      for (let at = 0; at < bytes.length; ) {
        const len = 1 + Math.floor(r() * 7);
        chunks.push(bytes.subarray(at, at + len));
        at += len;
      }
      assert.deepEqual(rows(chunks, opts), whole, `seed ${seed}: ${JSON.stringify(text)}`);
    }
  }
});
