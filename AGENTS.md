# SEC--data-analysys

Python 3.12 project(s) for analyzing company financial data (SEC EDGAR XBRL
CompanyFacts, BNDES/BCB open data), with CSV/offline fallbacks. Everything is a
command-line / script based Python workload — there is no long-running server on
`main`.

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` + this file). The
  actual applications live on feature branches, each a self-contained Python
  project with its own `requirements.txt`. Notable ones:
  - `cursor/analise-financeira-empresas-*` — generic financial analyzer CLI
    (`analyze_finance.py`), has an offline `--demo` mode.
  - `cursor/petrobras-lucro-liquido-*` — Petrobras net-income extractor.
  - `cursor/financial-flows-processing-*`, `cursor/fluxos-*`,
    `cursor/resumo-*` — BNDES cash-flow / fiscal-impact generators that emit
    CSV/XLSX; some also ship a Streamlit `app.py`.
  When starting work, base your branch on `main` and bring in code from the
  relevant feature branch / PR.
- Python 3.12 and `pip` are provided by the base image; common data libs
  (pandas, numpy, requests, openpyxl, pyarrow, python-dateutil, tabulate,
  pytest) are already available. Ubuntu 24.04 is PEP 668 "externally managed",
  so install into the environment with
  `pip install --break-system-packages -r requirements.txt`. The startup update
  script runs that install only when a `requirements.txt` exists, so on a bare
  `main` checkout it is a no-op — after checking out a feature branch, re-run it
  (or the update script) to pull that branch's deps (e.g. `streamlit`,
  `matplotlib`, `plotly`).
- Run tests with `python3 -m pytest -q` from the branch root.
- Run the analyzer offline (no network) with `python3 analyze_finance.py --demo`.
- Live SEC EDGAR / BCB / BNDES calls require network access and (for SEC) an
  identifiable `User-Agent` (e.g. `--user-agent "Name email@domain"`). Network
  egress may be blocked in the cloud VM; prefer offline demo mode, `--excel`/
  local CSV inputs, or committed data caches to run end-to-end without hitting
  external APIs.
- No linter is configured; `python3 -m py_compile <files>` is a reasonable
  smoke check.
