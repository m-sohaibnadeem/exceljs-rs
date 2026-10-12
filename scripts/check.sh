#!/usr/bin/env bash
# Runs every check a change must pass (CONTRIBUTING.md, "Checks"), stopping
# at the first failure. The ExcelJS spec suite needs compat/upstream
# (compat/fetch-upstream.sh) and compat/node_modules (npm ci in compat/).
set -euo pipefail
cd "$(dirname "$0")/.."

cargo fmt --all --check
cargo clippy --workspace --all-targets --locked -- -D warnings
cargo test --workspace --locked
(cd crates/excel-rs-node && npm run build:debug && npm test && npm run test:types)
node docs/build.js
node scripts/check-docs.js
(cd compat && node run.js)
echo "all checks passed"
