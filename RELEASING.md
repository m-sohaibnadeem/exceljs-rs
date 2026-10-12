# Releasing excel-rs

GitHub Actions builds, tests, and publishes all five native targets and the main
JavaScript package. The first publication can use a temporary npm bootstrap
token. Later releases use npm trusted publishing. Local publication of downloaded
artifacts is also supported. No credentials belong in this repository. The Rust
crates are not published to crates.io.

## Before the first release

Confirm ownership of the GitHub repository and npm package names. This project
uses `m-sohaibnadeem/excel-rs` and these npm packages:

- `excel-rs`
- `excel-rs-linux-x64-gnu`
- `excel-rs-linux-arm64-gnu`
- `excel-rs-darwin-x64`
- `excel-rs-darwin-arm64`
- `excel-rs-win32-x64-msvc`

An npm lookup returning 404 does not reserve a name or guarantee permission to
publish it. Use an npm account with two-factor authentication.

Keep the version in `Cargo.toml`, `Cargo.lock`, and
`crates/excel-rs-node/package.json` consistent. Update the npm lockfile after
changing package metadata. Add release notes in `CHANGELOG.md` and run the
checks in [CONTRIBUTING.md](CONTRIBUTING.md).

After committing and pushing the repository, wait for CI. Enable private
vulnerability reporting, secret scanning, push protection where available,
and branch protection requiring CI. Create an `npm` environment restricted
to release tags, with a required reviewer if supported by the repository plan.

## First publication through GitHub Actions

1. Push this source to the public repository named above. No Git initialization
   or remote creation is performed by the release scripts.
2. Create a GitHub environment named `npm`. Restrict its deployment tags to
   `v*` and add a required reviewer if available. Enable Actions for the repository.
3. Run **Release** manually on `main` with **publish = false**. Leave
   authentication as **trusted**. This builds all five platforms, runs CI, and
   tests the packed files on all five platforms with Node.js 22 and 24. It does
   not publish. Resolve any failures before continuing.
4. For packages that do not yet have trusted publishing, create a short-lived
   npm granular access token with **Read and write (publish and stage)** and
   **Bypass two-factor authentication**. It needs permission to create and
   publish all six package names. For these new unscoped names, that may require
   access to all packages on the publishing account; use this only for bootstrap.
   Store it as the `NPM_TOKEN` secret in the `npm` GitHub environment.
   Do not paste it into source files, shell commands, issue comments, or chat.
5. Commit the final version and lockfiles. Create and push the matching tag,
   initially `v0.1.0`. Automatic publication stays disabled until the repository
   variable `NPM_TRUSTED_PUBLISHING` is set to `true`.
6. Manually run **Release**, selecting the `v0.1.0` tag, **publish = true**, and
   **authentication = bootstrap-token**. Approve the environment deployment if
   configured. Publishing from a branch is rejected.
7. After success, verify all six npm package pages. Configure trusted publishing
   for each package using the settings below. Revoke the bootstrap token and
   remove the `NPM_TOKEN` environment secret when it is no longer needed.

The bootstrap token must permit direct publication. A stage-only token cannot
run this workflow's `npm publish` command. npm currently supports bypass-2FA
publication tokens, but plans to remove direct token publishing in January 2027;
use the local first-publication path if bootstrap tokens are no longer available.
See [npm token documentation](https://docs.npmjs.com/about-access-tokens/) and
[npm's publishing authentication requirements](https://docs.npmjs.com/requiring-2fa-for-package-publishing-and-settings-modification/).

A failed partial publication can be retried with the **same downloaded artifact**.
The publisher skips an existing version only when its registry integrity matches
that archive exactly. A new build can produce different ZIP or binary bytes;
if a retry fails the integrity check, use the original artifact or a new version.
The main package is published only after all platform packages are available.

## Optional first publication from Windows

A Windows build cannot supply the Linux and macOS binaries. Once the repository
is on GitHub, manually run the **Release** workflow on the exact commit to
publish. A manual run with **publish = false** builds and tests all five targets and
produces the `npm-release` artifact. It does not publish anything.

Download `npm-release` from that successful workflow run and extract the six
`.tgz` files and `release-manifest.json` to `crates/excel-rs-node/release/`. Do not combine files from different
runs or versions. Check the run's commit against your local checkout. The
workflow checks version consistency, licenses, native artifacts, package
contents, and a Linux installation of the packed package before uploading.

From PowerShell at the repository root, inspect the tarballs and confirm npm
authentication:

```powershell
npm login
npm whoami
$version = (Get-Content crates/excel-rs-node/package.json | ConvertFrom-Json).version
$releaseDir = Join-Path (Get-Location) 'crates/excel-rs-node/release'
$names = @('linux-x64-gnu', 'linux-arm64-gnu', 'darwin-x64', 'darwin-arm64', 'win32-x64-msvc')
foreach ($name in $names) {
  npm publish (Join-Path $releaseDir "excel-rs-$name-$version.tgz") --access public --ignore-scripts --dry-run
  if ($LASTEXITCODE -ne 0) { throw "Preview failed: $name" }
}
npm publish (Join-Path $releaseDir "excel-rs-$version.tgz") --access public --ignore-scripts --dry-run
```

Inspect the previews, including package names, versions, licenses, and files.
When ready to publish, repeat those commands without `--dry-run`, publishing
all five platform packages before `excel-rs`. This is an irreversible registry
operation. Local publishing may request two-factor authentication. Do not add
`--provenance` to this local release: the machine is not a supported CI identity.

Install the published version in a fresh directory and run the README example.
Use `npm view <package>@<version>` to verify all six packages. Record the first
release in GitHub without pushing a release tag that would publish it a second time.

## Later GitHub Actions releases

Configure a trusted publisher in npm for **each of the six packages**, using:

- GitHub owner: `m-sohaibnadeem`
- Repository: `excel-rs`
- Workflow filename: `release.yml`
- Environment: `npm`
- Allowed action: direct publication with `npm publish`

The workflow uses Node.js 24, npm 11.12.1, and a publishing job with
`id-token: write`. It publishes the checked tarballs with provenance. It does
not require an `NPM_TOKEN` secret in trusted mode. See [npm's trusted publishing documentation](https://docs.npmjs.com/trusted-publishers/).

Set the repository Actions variable `NPM_TRUSTED_PUBLISHING` to `true` only
after all six trusted publisher configurations are ready. New npm trusted
publisher configurations must complete a successful publish within two days;
configure them close to the next release.

Update the version and changelog, commit, and wait for CI. Push a matching
`v<version>` tag. The release workflow runs CI again, builds and tests every
native target, verifies all packages, and publishes through the `npm` environment.
Never move a release tag to a different commit.

If publication stops partway through, rerun only the failed publish job so it
uses the same checked artifact set. The publisher verifies existing versions
and skips exact matches. Conflicting bytes stop the release rather than being
silently accepted. Versions are immutable.

## Preparing artifacts locally

These commands collect existing builds. They do not build Linux or macOS
binaries on a Windows machine. An empty `artifacts/` directory is not enough.
Do not run this sequence until all five tested native binaries are available.
The package smoke test checks only the machine's current platform; passing it
does not establish that the full release is ready.

For release work outside the workflow, obtain all five platform binaries and
the generated `native.js` and `native.d.ts` from the same tested commit.
Place downloaded workflow artifacts under `crates/excel-rs-node/artifacts/`,
preserving their artifact folders. From `crates/excel-rs-node`:

```sh
npm ci
```

Copy `artifacts/js-binding/native.js` and `native.d.ts` into this directory, then:

```sh
npx --no-install napi create-npm-dirs
npx --no-install napi artifacts
npm run release:prepare
npm run release:publish -- --dry-run
npm run test:package
npm run release:pack
```

Preparation writes generated platform metadata and adds exact-version optional
dependencies to the main package manifest. It never publishes. Review generated
metadata before committing source changes. `npm run release:publish` without
`--dry-run` explicitly publishes all platforms followed by the main package.
The `prepublishOnly` hook only validates inputs and does not publish packages.

## GitHub Pages

In repository Settings, open Pages and select **GitHub Actions** as the source.
The Docs workflow builds `docs/_site` and deploys it using the `github-pages`
environment. It runs after documentation changes on `main` or by manual dispatch.
See [GitHub's Pages configuration guide](https://docs.github.com/en/pages/getting-started-with-github-pages/configuring-a-publishing-source-for-your-github-pages-site).

Check the deployed site at `https://m-sohaibnadeem.github.io/excel-rs/`, including
navigation, API examples, and mobile layout. No custom domain is configured.
