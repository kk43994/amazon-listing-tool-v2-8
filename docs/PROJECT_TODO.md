# Project Improvement TODO

Last reviewed: 2026-04-28
Scope: customer handoff package, Web UI, backend/API, Amazon SP-API flow, AI configuration, tests, release operations.

## Current Snapshot

- The app is already usable as a packaged local Web tool, with PyInstaller release, startup scripts, Doctor check, first-run guide, setup status API, and full test suite passing locally.
- Largest maintainability risks are concentrated in `web/app.py` (~5.5k lines) and `web/templates/_js.txt` (~7.8k lines).
- The customer-facing risk is not a single bug; it is the whole onboarding path: unzip -> launch -> configure AI -> configure Amazon -> import -> process -> preview -> submit.
- The next goal should be: a non-technical customer can finish first setup without reading developer docs or asking which button to press.

## P0 - Before Next Customer Release

### Release / Package Hardening

- [x] Build both macOS and Windows release artifacts from CI and verify both have `一键检测修复`, `环境检测`, `客户先看这里`, `.env.example`, template Excel, and docs in the root. (GitHub Actions matrix now builds/verifies macOS + Windows.)
- [ ] Run packaged `--doctor --quiet` on a clean machine/VM for macOS and Windows, not only on the developer machine.
- [x] Add a CI assertion that generated zip contains expected root files by exact logical name; current terminal zip listing may show garbled Chinese names, so verify extraction behavior on Windows Explorer and macOS Finder. (Added `tools/verify_release.py`; manual Finder/Explorer check still belongs in release checklist.)
- [x] Add a packaged smoke test that starts the app on a random free port, hits `/api/setup-status`, then exits cleanly.
- [x] Decide whether to ship Chinese filenames or duplicate ASCII fallback launchers, e.g. `Start-Amazon-2.8.command`, `Doctor.command`, `Read-Me-First.txt`, to avoid encoding and security-software issues.

### Customer Setup Wizard

- [x] Turn current setup cards into a true first-run wizard: Step 1 AI -> Step 2 Amazon account -> Step 3 import/sample -> Step 4 self-check.
- [x] Add a visible Simple Mode / Advanced Mode switch; hide endpoint/model fields by default and expose them only in Advanced Mode.
- [x] In Simple Mode, show only: text key, image key, one-click recommended config, save, test connection.
- [x] Add a dedicated Amazon account setup wizard with field-by-field examples and validation before saving.
- [x] Add “copy support info” button on setup failures that copies Doctor status, app version, OS, runtime dir, and masked config.

### Configuration Guardrails

- [x] Add backend validation for AI config, not only frontend hints: reject missing/legacy Base URL when user tries to save recommended customer mode.
- [x] Add backend validation for Amazon account format: Seller ID, LWA Client ID, Refresh Token, marketplace id mapping.
- [x] Add a one-click “reset API fields to recommended” endpoint so support can recover customers who edited advanced fields wrongly.
- [x] Persist whether the customer completed first-run setup; show the guide again if `.env` or `accounts.json` becomes incomplete.
- [x] Make self-check results actionable: each failed check should include a “fix now” button or exact next click.

### Submission Safety

- [x] Keep formal submit blocked unless selected SKUs have passed preview in the same account/marketplace recently.
- [x] Show a final “formal submit” confirmation with account, marketplace, SKU count, preview pass count, and irreversible warning.
- [x] Add a per-SKU submit eligibility reason list before submit, not just skipped counts.
- [x] Write a backup copy before any Excel mutation from validation, listing check, AI adoption, submit status writeback.
- [x] Add “open output folder” and “open latest backup” actions in the UI/package scripts.

## P1 - Customer Support / Doctor 2.0

### Doctor Program

- [x] Add Doctor checks for Windows Defender/quarantine symptoms: missing `_internal`, blocked exe, non-writable directory, running inside zip/download temp path.
- [x] Add Doctor check for launching from cloud-sync paths that often lock files: iCloud Drive, OneDrive, Dropbox, network share.
- [x] Add Doctor network mode with separate DNS resolution, TCP connect, TLS handshake, and HTTP status for `api.kk666.best`, Amazon LWA, SP-API regions.
- [x] Add Doctor support bundle export: zip `doctor-report.txt`, app logs, task history, setup status JSON, masked `.env`, masked `accounts.json`.
- [x] Add Doctor repair for common config mistakes: `api.kk666.online`, missing scheme, trailing `/v1` if using Gemini endpoint, bad endpoint without `{model}`.
- [x] Add offline dependency inventory JSON generated during build; Doctor should compare bundled modules against this manifest.

### Installer / Distribution

- [ ] Consider a real installer for Windows with Start Menu shortcut and clean data directory, instead of asking customers to run from arbitrary unzip paths.
- [ ] Consider macOS `.dmg` packaging and code signing/notarization to reduce Gatekeeper friction.
- [x] Generate SHA256 checksums for release zips and show them in Release notes.
- [x] Add version file and `/api/version` so support can tell exactly what customer is running.
- [ ] Add auto-update notice or at least “check latest version” link if network is available.

## P1 - Web UX Foolproofing

### Workspace Flow

- [x] Make “傻瓜流程” the primary workspace control: one big next-step button, with toolbar as secondary advanced controls.
- [x] Disable or visually lock risky toolbar buttons until prerequisites are met; current guards catch clicks, but disabled state would be clearer.
- [x] Add hover/tooltips to every locked step explaining why it is locked and what to do next.
- [x] Add a demo/sample Excel import button so customers can verify the tool works before using real data.
- [x] Add empty-state screens for no file/no account/no AI instead of showing an empty table.

### Settings Page

- [x] Split settings into tabs/cards: “AI”, “Amazon account”, “Self-check”, “Advanced”.
- [x] Add copy buttons next to recommended Base URL, endpoint, and model names.
- [x] Add inline field validation indicators while typing, before Save.
- [x] Add a “test text AI only” and “test image AI only”; current combined tests can make failures harder to isolate.
- [x] Add account test details: LWA token ok, marketplace ok, listings API ok, permission ok.

### Error Messages

- [ ] Normalize frontend/backed error messages to customer language: what happened, why, next click.
- [x] Redact all secrets in errors, logs, toasts, task details, and support bundles.
- [x] Add error codes such as `AI_BASE_BAD`, `AMAZON_TOKEN_FAIL`, `EXCEL_LOCKED`, so support can search quickly. (Started with AI/Amazon setup errors; continue normalizing runtime errors.)
- [x] For Excel file locked by WPS/Excel, show a dedicated message: close Excel and retry.

## P1 - Amazon SP-API Reliability

- [ ] Add official OAuth/token acquisition helper or at least a guided refresh-token checklist; manual Refresh Token is still too hard for customers.
- [x] Store account test timestamps and last failure reason, display them in account table.
- [x] Separate preview result by account and marketplace; do not treat old preview from another account as safe.
- [ ] Add rate limiting/backoff visualization for SP-API calls; show “waiting due to Amazon rate limit” instead of appearing stuck.
- [ ] Complete/verify fallback to Feeds API for very large batches; route should be explicit in task result.
- [ ] Add more tests around marketplace-specific required fields and product type schemas.

## P1 - Data Safety / Excel Integrity

- [x] Add atomic writes for all Excel changes: write temp file, verify, then replace.
- [x] Keep timestamped backups in `backups/` for every imported file mutation.
- [ ] Add a restore UI for backups.
- [x] Add file lock detection before writeback, especially Windows Excel/WPS lock behavior.
- [ ] Add a data diff preview before applying AI/adoption/bulk edits to many SKUs.
- [ ] Clean old temporary template folders under `output/tmp_template_*` with retention policy.

## P1 - Test Coverage / QA

- [ ] Add Playwright/browser E2E tests for first-run setup, AI config save, account modal validation, import, guard blocks, preview gate.
- [x] Add release artifact tests for both Windows and macOS in GitHub Actions.
- [ ] Add contract tests for each `/api/*` endpoint: schema, errors, auth/config prerequisites.
- [ ] Add tests for Doctor repair mode on a temp runtime directory.
- [x] Add tests for dangerous submit guard: cannot submit without preview; cannot submit with preview from wrong account.
- [x] Add snapshot tests for generated customer readme/quick start scripts so release packaging changes do not regress. (Covered by release root-file verifier; deeper text snapshots can still be added.)

## P2 - Codebase Maintainability

### Backend Split

- [ ] Split `web/app.py` into Flask blueprints: files, config, accounts, AI tasks, validation, templates, submit, diagnostics.
- [ ] Move Excel writeback helpers from `web/app.py` into `core/excel/` services.
- [ ] Move task history/status management into `core/tasks.py` with a small persistence layer.
- [ ] Move setup-status/self-check/doctor-related logic into a shared service so CLI Doctor and Web self-check do not drift.
- [ ] Add typed response models or lightweight dataclasses for key API payloads.

### Frontend Split

- [ ] Split `web/templates/_js.txt` into modules: state, api client, setup guide, table, drawer, tasks, submit, settings.
- [ ] Introduce a small frontend build step only if packaging remains reliable; otherwise use separate static files loaded directly.
- [ ] Centralize prerequisite guards, toast messages, and modal helpers.
- [ ] Add frontend unit tests for flow resolution and guard decisions.
- [ ] Remove inline `onclick` gradually in favor of event listeners for maintainability.

### Lint / Static Checks

- [ ] Add `ruff` config and run it in CI with an initial conservative rule set.
- [ ] Add `pyright` or `mypy` only after splitting modules; start with non-strict on core services.
- [x] Add `eslint`/`prettier` or at least `node --check` in CI for the generated frontend JS.
- [x] Add a script `tools/verify_release.py` that runs all local release checks with one command.

## P2 - Product Enhancements

- [ ] Add role-specific home screen: “运营使用” vs “技术配置”.
- [ ] Add guided product-type selection with examples and recent choices.
- [ ] Add AI cost/time estimate before batch AI processing.
- [ ] Add batch pause/resume for long AI/image jobs.
- [ ] Add per-SKU checklist drawer: AI done, local validation, listing check, preview, submit.
- [ ] Add exportable customer operation report: what was processed, what failed, next action.

## P2 - Observability / Support

- [ ] Add structured logs with request id/task id/SKU/account id but no secrets.
- [ ] Add UI log viewer for recent errors and Doctor results.
- [x] Add one-click “send support bundle” placeholder: create zip locally, user manually sends it.
- [ ] Add task stuck detection: if progress unchanged for N minutes, show likely cause and support bundle button.
- [ ] Add metrics counters locally: imports, AI jobs, preview attempts, submit success/fail.

## P3 - Documentation Cleanup

- [ ] Archive or update old project plan docs that still mention outdated domains/models or old workflow assumptions.
- [ ] Create one final customer-facing PDF/HTML quick guide with screenshots from current UI.
- [ ] Create one internal support runbook: common Doctor warnings, exact fixes, screenshot examples.
- [x] Create release checklist in `docs/客户部署说明.md` or a dedicated `docs/RELEASE_CHECKLIST.md`.
- [ ] Keep changelog and README aligned with actual release assets.

## Suggested Execution Order

1. Finish current release changes and cut `v0.4.5` after Windows/macOS CI succeeds.
2. Add Simple/Advanced settings mode and backend config validation.
3. Add support bundle and Doctor repair expansion.
4. Add Playwright E2E for first-run and submit guard.
5. Split `web/app.py` and `_js.txt` only after customer release is stable.

## Definition of Done for “Dummy-Proof”

- A fresh customer can unzip, launch, configure AI, configure Amazon, import a sample, run self-check, and understand next step without talking to support.
- Every blocked action explains the exact reason and provides one recommended next button.
- Every support request can start with `一键检测修复` and a single `doctor-report.txt`.
- Formal submit cannot happen accidentally or without a recent successful preview for the same account/marketplace.
- Release artifacts are smoke-tested in CI and manually verified on clean macOS + Windows machines.
