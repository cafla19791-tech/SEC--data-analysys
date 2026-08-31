# SEC--data-analysys

Python 3.12 project for financial/public-finance data analysis. It contains
several self-contained products, each on its own feature branch with its own
`requirements.txt`:

- `analyze_finance.py` — generic company financial-analyzer CLI over SEC EDGAR
  XBRL CompanyFacts, with an offline `--demo` mode.
  Branch: `cursor/analise-financeira-empresas-*`.
- `extract_petrobras_net_income.py` — Petrobras net-income extractor (SEC EDGAR
  + BCB USD/BRL) with committed data caches under `data/`.
  Branch: `cursor/petrobras-lucro-liquido-*`.
- `scripts/gerar_fluxos*.py` — BNDES cash-flow / subsidy generators.
  Branches: `cursor/fluxos-*` and `cursor/financial-flows-processing-*`.
- `scripts/gerar_fluxos.py` + `scripts/resumo_por_agente.py` + `app.py` — BNDES
  cash-flow generator with a per-agent ranking and a **Streamlit** web
  dashboard ("Resumo por Agente Financeiro").
  Branch: `cursor/resumo-agente-financeiro-*` (also `cursor/resumo-por-agente-excel-*`).

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` and this
  `AGENTS.md`). The actual applications live on the feature branches listed
  above, each with its own `requirements.txt`. When starting work, base your
  branch on `main` and bring in code from the relevant feature branch / PR. A
  clean, non-destructive way to run a branch's code is a worktree, e.g.
  `git worktree add /tmp/app origin/<branch>` and work there.
- Dependencies install with `pip install --break-system-packages -r
  requirements.txt` (Ubuntu 24.04 is PEP 668 "externally managed"; no
  virtualenv system package is guaranteed). The startup update script installs
  deps only when a `requirements.txt` exists at the repo root, so on a bare
  `main` checkout it is a no-op — after checking out / worktree-ing a feature
  branch, run that branch's install yourself.
- `pip` installs console scripts (`streamlit`, `pytest`, …) into `~/.local/bin`,
  which is not on `PATH` by default. Invoke via the module form to avoid PATH
  issues: `python3 -m pytest`, `python3 -m streamlit run app.py`.
- No lint/build tooling is configured anywhere (no ruff/flake8/mypy, no
  compile/bundle step, no CI). "Build" for these interpreted Python tools is
  just installing deps.
- Run tests from the branch root. The `scripts/`-based BNDES branches need
  `PYTHONPATH=. python3 -m pytest tests/ -q`.
- BNDES generator + Streamlit dashboard (offline sample): first generate the
  ranking the UI reads with
  `python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv`
  (writes `output/resumo_por_agente.csv`), then serve the UI with
  `python3 -m streamlit run app.py --server.port 8501 --server.headless true`
  (port 8501; `--server.headless true` avoids the interactive email prompt).
  A sample `output/resumo_por_agente.csv` is committed on the Streamlit branch,
  so the dashboard renders immediately. The four top metric cards intentionally
  show full-dataset totals and do NOT react to the sidebar search filter — only
  the chart/table update.
- Run the generic SEC analyzer offline: `python3 analyze_finance.py --demo`.
- Run the Petrobras extractor offline: `python3 extract_petrobras_net_income.py
  --years 10` (uses the committed `data/` caches; only `--refresh` hits the
  network).
- Live SEC EDGAR / BCB / BNDES calls require network egress and an identifiable
  `User-Agent` (e.g. `--user-agent "Name email@domain"`), which may be blocked
  in the cloud VM. Prefer offline demo mode / committed caches / the sample
  CSVs; the BNDES generators need `--download` (network) or a local
  `--excel`/`--input` file to produce output.
