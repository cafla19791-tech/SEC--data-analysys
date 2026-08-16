# SEC--data-analysys

Python 3.12 project for analyzing company financial data (SEC EDGAR / Brazilian
BNDES & BCB open data), producing cash-flow projections, fiscal-impact summaries,
and rankings — with CLI scripts, CSV/Excel exports, and a Streamlit web UI.

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` and this
  `AGENTS.md`). The real applications live on feature branches, each with its own
  `requirements.txt` (e.g. the BNDES cash-flow generator + "resumo por agente"
  Streamlit app under `scripts/` + `app.py`, a generic financial-analyzer CLI
  `analyze_finance.py`, a Petrobras net-income extractor, etc.). Base your branch
  on `main` and bring in code from the relevant feature branch / PR before running.
- Dependencies install with `pip install --break-system-packages -r requirements.txt`
  (Ubuntu 24.04 is PEP 668 "externally managed"; there is no guaranteed venv). The
  startup update script runs this only when a `requirements.txt` exists, so on a
  bare `main` checkout it is a no-op — after checking out a feature branch, re-run
  the install (or the update script) to pull that branch's deps.
- pip installs console scripts (`streamlit`, `pytest`, …) into `~/.local/bin`,
  which is not on `PATH` by default. Invoke via the module form to avoid PATH
  issues: `python3 -m pytest`, `python3 -m streamlit run app.py`.
- Run tests from the branch root. Most branches need `PYTHONPATH=.`:
  `PYTHONPATH=. python3 -m pytest tests/ -q`.
- Run the BNDES cash-flow + ranking pipeline offline with the committed sample:
  `python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv`.
  This writes `output/resumo_por_agente.csv` which the web UI reads.
- Run the web UI (after generating the resumo above) with
  `python3 -m streamlit run app.py --server.port 8501 --server.headless true`.
  It serves on port 8501.
- Run the generic analyzer offline (no network) with
  `python3 analyze_finance.py --demo`.
- Live SEC EDGAR / BCB / BNDES calls require network egress and an identifiable
  `User-Agent` (e.g. `--user-agent "Name email@domain"`), which may be blocked in
  the cloud VM. Prefer offline demo mode or the committed sample/cache inputs
  (e.g. `--input`/`--excel` for BNDES, `--facts-cache`/`--fx-cache` for Petrobras).
