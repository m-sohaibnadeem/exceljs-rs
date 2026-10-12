# Contributing

Issues and pull requests are welcome. Include a minimal reproduction for bugs,
with confidential data removed. Report vulnerabilities through [SECURITY.md](SECURITY.md).

## Build

Install Node.js 22.14 or later, npm, Rust stable, and the native build tools for
your platform. On Windows, install Visual Studio Build Tools with the C++ workload
and Windows SDK. On macOS, install Xcode command line tools. On Linux, install
a C compiler and linker.

From the repository root:

```sh
npm ci --prefix crates/exceljs-rs-node
npm run build:debug --prefix crates/exceljs-rs-node
npm test --prefix crates/exceljs-rs-node
```

Use `npm run build --prefix crates/exceljs-rs-node` for an optimized addon.
Rebuild after changing Rust code. Generated loaders and native binaries are
ignored by Git. Source builds use `require('./crates/exceljs-rs-node')` from the
repository root. The Rust crates support the Node package and are not published
to crates.io as a stable Rust API.

## Checks

Run from the repository root:

```sh
cargo fmt --all --check
cargo clippy --workspace --all-targets --locked -- -D warnings
cargo test --workspace --locked
npm run build:debug --prefix crates/exceljs-rs-node
npm test --prefix crates/exceljs-rs-node
node docs/build.js
node scripts/check-docs.js
```

For the selected ExcelJS 4.4.0 public API specs:

```sh
npm ci --prefix compat
node compat/fetch-upstream.js
node compat/run.js
```

The upstream source is pinned to a commit. `compat/suite.js` records included
and excluded specs. Test dependencies are development tools and are not shipped
in the npm package. The test harness pins ExcelJS 4.4.0 and overrides its uuid dependency with a
patched version. Do not expose test servers to a network. Run npm audits in
all three package directories and `cargo audit` when updating dependencies.

Fixture snapshots cover models, cell values, styles, round trips, and malformed
input. Regenerate snapshots only for intended behavior changes and review each
difference. Add regressions for fixes and document API differences in
[COMPATIBILITY.md](COMPATIBILITY.md).

## Architecture

- `crates/exceljs-rs-node/lib/doc`: JavaScript workbook, worksheet, row, and cell objects.
- `crates/exceljs-rs-node/lib/xlsx`, `lib/csv`, and `lib/stream`: public I/O APIs.
- `crates/exceljs-rs/src`: Rust ZIP, XML, CSV, and model processing.
- `crates/exceljs-rs-node/src`: N-API bindings.
- `fixtures`: input workbooks and regression snapshots.
- `compat`: upstream public API compatibility tests.

Preserve third-party source notices. The JavaScript model and portions of the
engine and types include MIT-licensed code from ExcelJS and other projects;
see the root license notices. Compatibility is the public API target, not a
claim of independent authorship of all source code.

The engine forbids unsafe Rust. Handwritten addon code denies it, while generated
N-API bindings are exempt. Return errors for invalid input. Discuss new dependencies
and behavior changes in an issue or pull request.

## Documentation

Edit `docs/pages/*.html` and run `node docs/build.js`. Open
`docs/_site/index.html` for a local preview. Relative links support the GitHub
Pages project path. The build also renders compatibility and security guidance.
Keep examples runnable and avoid claims that tests or measurements do not support.

## Benchmarks

```sh
npm run build --prefix crates/exceljs-rs-node
npm ci --prefix bench
node bench/run.js 100000
```

Record hardware, OS, Node.js version, source revision, and raw output with new
measurements. Do not treat one workload as a general speed or memory guarantee.

## Releases

See [RELEASING.md](RELEASING.md). Publication is an explicit maintainer action.

Regenerate native dependency notices with `node scripts/licenses.js` after
changing Cargo.lock or the Rust toolchain. Review license changes before release.
