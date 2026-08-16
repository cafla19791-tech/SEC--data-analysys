# SEC--data-analysys

Python 3.12 project(s) for analyzing company financial data (SEC EDGAR XBRL
CompanyFacts, BNDES/BCB open data), with CSV/offline fallbacks. Everything is a
command-line / script based Python workload plus one optional Streamlit web UI —
there is no long-running backend server or database.

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` + this file). The
  actual applications live on feature branches, each a self-contained Python
  project with its own `requirements.txt`. Notable ones:
  - `cursor/analise-financeira-empresas-*` — generic financial analyzer CLI
    (`analyze_finance.py`), has an offline `--demo` mode.
  - `cursor/petrobras-lucro-liquido-*` — Petrobras net-income extractor.
  - `cursor/financial-flows-processing-*`, `cursor/fluxos-*`, `cursor/contagil-*`,
    `cursor/resumo-*` — BNDES cash-flow / fiscal-impact generators that emit
    CSV/XLSX; some also ship a Streamlit `app.py`.
  When starting work, base your branch on `main` and bring in code from the
  relevant feature branch / PR. To inspect a branch without leaving `main`, use
  `git worktree add /tmp/wt origin/<branch>`.
- Python 3.12 and `pip` are provided by the base image, but data libraries are
  NOT pre-installed. Ubuntu 24.04 is PEP 668 "externally managed", so install
  with `pip install --break-system-packages -r requirements.txt`. The startup
  update script runs that install only when a `requirements.txt` exists at the
  repo root, so on a bare `main` checkout it is a no-op — after bringing in a
  feature branch's code, re-run the update script (or the pip command above) to
  pull that branch's deps (e.g. `streamlit`, `matplotlib`, `plotly`).
- `pip` installs console scripts (e.g. `pytest`, `streamlit`) to
  `~/.local/bin`, which is not on `PATH`. Invoke them as modules instead:
  `python3 -m pytest`, `python3 -m streamlit run app.py`.
- Run tests with `python3 -m pytest -q` from the branch root. BNDES branches
  with a `scripts/` package need `PYTHONPATH=. python3 -m pytest -q`.
- Run the analyzer offline (no network): `python3 analyze_finance.py --demo`.
- Streamlit dashboard branches: `python3 -m streamlit run app.py
  --server.headless true --server.port 8501` (the `--server.headless true` flag
  avoids Streamlit's interactive email prompt). The UI reads committed
  `output/*.csv|xlsx`, so it renders immediately without regenerating data.
- Live SEC EDGAR / BCB / BNDES calls require network access and (for SEC) an
  identifiable `User-Agent` (e.g. `--user-agent "Name email@domain"`). Network
  egress may be blocked in the cloud VM; prefer offline demo mode, `--excel`/
  local CSV inputs, or committed data caches to run end-to-end without hitting
  external APIs.
- No linter is configured; `python3 -m py_compile <files>` is a reasonable
  smoke check.
