# Repository Guidelines

## Project Structure & Module Organization
- `src/`: VBA codebase. Patterns: `I*` (interfaces), `C*` (implementations), `E*` (entities/DTO), `mod*` (modules), tests `Test*.bas` and helpers `TI*.bas`.
- `front/Desarrollo/`: Development Access frontend (`CONDOR.accdb`).
- `ui/`: Form sources and assets (`ui/sources/*.accdb`, `ui/definitions/*.json`, `ui/templates`, `ui/assets`).
- `back/data/`: Access back-end databases (`*.accdb`).
- `db/migrations/`: SQL migration scripts.
- `scripts/`: CLI (`scripts/condor_cli.vbs`) and utilities.
- `docs/`: Architecture and process notes.

## Build, Test, and Development Commands
- Requires Windows with Microsoft Access and ACE OLEDB installed.
- List forms: `cscript scripts\condor_cli.vbs list-forms [--json] [--db <path>]`.
- Export form: `cscript scripts\condor_cli.vbs export-form <db> <FormName> --output .\out\Form.json`.
- Import form: `cscript scripts\condor_cli.vbs import-form <db> .\out\Form.json --target <FormName> --replace`.
- Roundtrip check: `cscript scripts\condor_cli.vbs roundtrip-form <db> <FormName> --temp .\temp`.
- List modules: `cscript scripts\condor_cli.vbs list-modules [--json] [--diff]`.
- Defaults: code/test commands resolve to `front\Desarrollo\CONDOR.accdb`; data commands to `back\data\CONDOR_datos.accdb` unless `--db` or `CONDOR_DB` is set.

## Coding Style & Naming Conventions
- Indentation: 4 spaces, no tabs; one statement per line.
- Naming: PascalCase for procedures; constants UPPER_SNAKE_CASE; follow `I*`, `C*`, `E*`, `mod*`, `Test*` patterns (e.g., `CWorkflowService.cls`, `modTestRunner.bas`).
- Error handling: use centralized handlers (see `modErrorHandler*`), log via `COperationLogger`.
- Comments in Spanish; keep functions small and single‑purpose.

## Testing Guidelines
- Place unit tests in `src/` using `Test*.bas`; mock via `CMock*` classes.
- Run tests with `cscript scripts\condor_cli.vbs test` when available; alternatively, execute `modTestRunner.RunAll` from Access.
- Aim to cover service logic and repository boundaries; prefer deterministic data.

## Commit & Pull Request Guidelines
- Use Conventional Commits: `feat:`, `fix:`, `refactor:`, `docs:`, `chore:` (Spanish summaries OK). Keep title ≤72 chars.
- PRs: include scope/impact, linked issues, screenshots or diffs for form changes (JSON export), and note any `db/migrations` updates.
- Before committing, set `ENTORNO_FORZADO` to `ForzarNinguno` in `src\modConfig.bas`. Do not commit local credentials or backup `*.accdb` files.

## Security & Configuration Tips
- Never commit passwords or network paths; examples use placeholders.
- Large binaries belong under `back/data/` only when necessary; prefer scripts/migrations for reproducibility.

## Master document
- Este documento siempre ha de estar alineado con el código /docs/CONDOR_MASTER_PLAN.md
