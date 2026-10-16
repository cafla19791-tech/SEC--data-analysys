# SEC--data-analysys

Python 3.12 project for analyzing company financial data from the SEC EDGAR
API (XBRL CompanyFacts), with CSV/offline fallbacks. Includes a Petrobras
net-income extractor and several BNDES cash-flow generators.

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` and this
  `AGENTS.md`). The actual applications live on feature branches, each with its
  own `requirements.txt`:
  - `analyze_finance.py` — generic financial-analyzer CLI (has an offline
    `--demo` mode). Branch: `cursor/analise-financeira-empresas-*`.
  - `extract_petrobras_net_income.py` — Petrobras net-income extractor with
    committed SEC/BCB data caches. Branch: `cursor/petrobras-lucro-liquido-*`.
  - `scripts/gerar_fluxos*.py` — BNDES cash-flow generators. Branches:
    `cursor/fluxos-*` and `cursor/financial-flows-processing-*`.
  When starting work, base your branch on `main` and bring in code from the
  relevant feature branch / PR (e.g. `git worktree add /tmp/app origin/<branch>`).
- Dependencies are installed with `pip --break-system-packages` (Ubuntu 24.04
  is PEP 668 "externally managed"). The startup update script runs
  `pip install --break-system-packages -r requirements.txt` only when a
  `requirements.txt` exists, so on a bare `main` checkout it is a no-op —
  after checking out a feature branch, run that branch's
  `pip install --break-system-packages -r requirements.txt` to pull its deps.
- Run tests with `python3 -m pytest -q` from the branch root. Some BNDES
  branches whose code lives under `scripts/` may need
  `PYTHONPATH=. python3 -m pytest tests/ -q`.
- Run the analyzer offline (no network): `python3 analyze_finance.py --demo`.
- Run the Petrobras extractor offline: `python3 extract_petrobras_net_income.py
  --years 10` (uses the committed `data/` caches; only `--refresh` hits the
  network).
- Live SEC EDGAR / BCB / BNDES calls require network access and an identifiable
  `User-Agent` (e.g. `--user-agent "Name email@domain"`). Prefer the offline
  demo mode or committed caches; the BNDES generators need `--download`
  (network) or a local `--excel` file to produce output.
