'use strict';

// npm packs README.md and the licences from the package directory; the
// sources are at the repository root.
const fs = require('fs');
const path = require('path');

for (const file of ['README.md', 'LICENSE', 'LICENSE-EXCELJS', 'LICENSE-THIRD-PARTY', 'LICENSE-RUST-DEPENDENCIES']) {
  fs.copyFileSync(path.join(__dirname, '..', '..', '..', file), path.join(__dirname, '..', file));
}
