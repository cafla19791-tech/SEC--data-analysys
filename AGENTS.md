# SEC--data-analysys

Python 3.12 project for analyzing company financial data from the SEC EDGAR
API (XBRL CompanyFacts), with CSV/offline fallbacks.

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` and this
  `AGENTS.md`). The actual applications live on feature branches, each with its
  own `requirements.txt` (e.g. a generic financial-analyzer CLI
  `analyze_finance.py`, a Petrobras net-income extractor, and several BNDES
  cash-flow generators under `scripts/`). When starting work, base your branch
  on `main` and expect to bring code in from the relevant feature branch / PR.
- Dependencies are installed with `pip --break-system-packages` (Ubuntu 24.04
  is PEP 668 "externally managed", and no virtualenv system package is
  guaranteed in the base image). The startup update script runs
  `pip install --break-system-packages -r requirements.txt` only when a
  `requirements.txt` exists, so on a bare `main` checkout it is a no-op — after
  checking out a feature branch, re-run that install (or the update script)
  to pull in that branch's deps.
- Run tests with `python3 -m pytest -q` from the branch root (some BNDES
  branches need `PYTHONPATH=. python3 -m pytest tests/ -q`).
- Run the analyzer offline (no network) with `python3 analyze_finance.py --demo`.
- Live SEC EDGAR / BCB / BNDES calls require network access and an identifiable
  `User-Agent` (e.g. `--user-agent "Name email@domain"`). Network egress may be
  blocked in the cloud VM; prefer offline demo mode or the committed
  data caches (e.g. pass `--facts-cache`/`--fx-cache` to the Petrobras script,
  or `--input`/`--excel` to the BNDES scripts) to run end-to-end without
  hitting external APIs.
