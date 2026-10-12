'use strict';

// Builds the documentation site into docs/_site (what GitHub Pages serves):
// each docs/pages/*.html gets the shared header and footer, and
// COMPATIBILITY.md becomes compatibility.html, so the two never disagree.
//
//   node docs/build.js        then open docs/_site/index.html
//
// A page starts with a comment giving its title and description:
//   <!-- title: ...
//        description: ... -->

const fs = require('fs');
const path = require('path');

const DOCS = __dirname;
const OUT = path.join(DOCS, '_site');
const NAV = [
  ['guide.html', 'Guide'],
  ['api.html', 'API'],
  ['benchmarks.html', 'Benchmarks'],
  ['compatibility.html', 'Compatibility'],
  ['security.html', 'Security'],
];

function escapeHtml(s) {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function layout(file, title, description, body) {
  const nav = NAV.map(
    ([href, label]) => `    <a href="${href}"${href === file ? ' aria-current="page"' : ''}>${label}</a>`
  ).join('\n');
  return `<!doctype html>
<html lang="en">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>${escapeHtml(title)}</title>
<meta name="description" content="${escapeHtml(description)}">
<link rel="stylesheet" href="style.css">
</head>
<body>
<a class="skip-link" href="#content">Skip to content</a>
<header class="top"><div class="top-inner">
  <a class="brand" href="./">excel<span>-rs</span></a>
  <nav class="nav" aria-label="Main">
${nav}
    <a class="gh" href="https://github.com/m-sohaibnadeem/excel-rs">GitHub</a>
  </nav>
</div></header>
${body.trim().replace('<main', '<main id="content"')}
<footer><div class="inner">
  excel-rs is MIT licensed. <a href="https://github.com/m-sohaibnadeem/excel-rs">Source on GitHub</a>.
</div></footer>
<script src="site.js"></script>
</body>
</html>
`;
}

// Inline Markdown: `code`, **bold**, [text](url).
function inline(text) {
  const codes = [];
  const held = text.replace(/`([^`]*)`/g, (m, code) => {
    codes.push(`<code>${escapeHtml(code)}</code>`);
    return `\u0000${codes.length - 1}\u0000`;
  });
  return escapeHtml(held)
    .replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>')
    .replace(/\[([^\]]+)\]\(([^)]+)\)/g, (m, label, href) => {
      // Repository files are linked on GitHub.
      const url = /^[a-z]+:/.test(href) ? href : `https://github.com/m-sohaibnadeem/excel-rs/blob/main/${href}`;
      return `<a href="${url}">${label}</a>`;
    })
    .replace(/\u0000(\d+)\u0000/g, (m, i) => codes[Number(i)]);
}

// The Markdown COMPATIBILITY.md uses: headings, paragraphs, and bullet lists
// whose items continue on indented lines.
function markdown(md) {
  const out = [];
  let para = [];
  let item = null;
  let list = false;
  const flushPara = () => {
    if (para.length) out.push(`<p>${inline(para.join(' '))}</p>`);
    para = [];
  };
  const flushItem = () => {
    if (item !== null) out.push(`<li>${inline(item)}</li>`);
    item = null;
  };
  const closeList = () => {
    flushItem();
    if (list) out.push('</ul>');
    list = false;
  };
  for (const line of md.split('\n')) {
    const heading = /^(#{1,3}) (.*)$/.exec(line);
    if (heading) {
      flushPara();
      closeList();
      const level = heading[1].length;
      const id = heading[2].toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
      out.push(`<h${level} id="${id}">${inline(heading[2])}</h${level}>`);
    } else if (line.startsWith('- ')) {
      flushPara();
      flushItem();
      if (!list) out.push('<ul>');
      list = true;
      item = line.slice(2);
    } else if (item !== null && /^\s+\S/.test(line)) {
      item += ` ${line.trim()}`;
    } else if (line.trim() === '') {
      flushPara();
      closeList();
    } else {
      closeList();
      para.push(line.trim());
    }
  }
  flushPara();
  closeList();
  return out.join('\n');
}

function build() {
  fs.rmSync(OUT, {recursive: true, force: true});
  fs.mkdirSync(OUT);
  for (const asset of ['style.css', 'site.js']) {
    fs.copyFileSync(path.join(DOCS, asset), path.join(OUT, asset));
  }
  fs.writeFileSync(path.join(OUT, '.nojekyll'), '');

  const pagesDir = path.join(DOCS, 'pages');
  for (const file of fs.readdirSync(pagesDir).filter(f => f.endsWith('.html'))) {
    const source = fs.readFileSync(path.join(pagesDir, file), 'utf8');
    const meta = /^<!--\s*title:\s*(.*?)\s*\n\s*description:\s*([\s\S]*?)\s*-->/.exec(source);
    if (!meta) throw new Error(`${file}: missing the title/description comment`);
    const body = source.slice(meta[0].length);
    fs.writeFileSync(path.join(OUT, file), layout(file, meta[1], meta[2], body));
  }

  const compat = fs.readFileSync(path.join(DOCS, '..', 'COMPATIBILITY.md'), 'utf8');
  fs.writeFileSync(
    path.join(OUT, 'compatibility.html'),
    layout(
      'compatibility.html',
      'Compatibility | excel-rs',
      'Every way excel-rs behaves differently from ExcelJS 4.4.0: the ExcelJS bugs it fixes and the engine differences.',
      `<main>\n${markdown(compat)}\n</main>`
    )
  );
  const security = fs.readFileSync(path.join(DOCS, '..', 'SECURITY.md'), 'utf8');
  fs.writeFileSync(path.join(OUT, 'security.html'), layout(
    'security.html', 'Security | excel-rs', 'Security reporting and safe handling of untrusted spreadsheet files.',
    `<main>\n${markdown(security)}\n</main>`
  ));
  console.log(`docs: built ${fs.readdirSync(OUT).length} files in ${path.relative(process.cwd(), OUT) || '.'}`);
}

build();
