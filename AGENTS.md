# SEC--data-analysys

Python 3.12 project for analyzing SEC EDGAR (XBRL CompanyFacts) and Brazilian
public-finance data (BNDES open data, BCB exchange rates), with CSV/offline
fallbacks. Includes a Petrobras net-income extractor and several BNDES
cash-flow generators (one with a Streamlit web UI).

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` and this
  `AGENTS.md`). The actual applications live on feature branches, each with its
  own `requirements.txt`. When starting work, base your branch on `main` and
  bring in code from the relevant feature branch / PR. A clean, non-destructive
  way to run a branch's code is a worktree, e.g.
  `git worktree add /tmp/app origin/<branch>`. Known product branches:
  - `cursor/analise-financeira-empresas-*` — `analyze_finance.py`, a generic
    financial-analyzer CLI with an offline `--demo` mode.
  - `cursor/petrobras-lucro-liquido-*` — `extract_petrobras_net_income.py`,
    Petrobras net-income extractor with committed SEC/BCB data caches in `data/`.
  - `cursor/fluxos-*` and `cursor/financial-flows-processing-*` — BNDES
    cash-flow generators under `scripts/gerar_fluxos*.py`.
  - `cursor/resumo-agente-financeiro-*` — BNDES flow generator + per-agent
    ranking, plus a **Streamlit** dashboard (`app.py`).
- Dependencies install with `pip install --break-system-packages -r
  requirements.txt` (Ubuntu 24.04 is PEP 668 "externally managed"; no
  virtualenv system package is guaranteed). The startup update script runs this
  only when a `requirements.txt` exists, so on a bare `main` checkout it is a
  no-op — after checking out / worktree-ing a feature branch, run that branch's
  install yourself.
- `pip` installs console scripts into `~/.local/bin`, which is not on `PATH`.
  Invoke tools via the module form instead, e.g. `python3 -m pytest`,
  `python3 -m streamlit run app.py`.
- No lint/build tooling is configured anywhere (no ruff/flake8/mypy/eslint, no
  compile/bundle step, no CI). "Build" for these interpreted Python tools is
  just installing deps.
- Run tests with `python3 -m pytest -q` from the branch root. The BNDES
  `scripts/`-based branches need `PYTHONPATH=. python3 -m pytest tests/ -q`.
- Run the analyzer offline (no network): `python3 analyze_finance.py --demo`.
- Run the Petrobras extractor offline: `python3 extract_petrobras_net_income.py
  --years 10` (uses committed `data/` caches; only `--refresh` hits the network).
- Streamlit dashboard (`cursor/resumo-agente-financeiro-*`): the app reads
  `output/resumo_por_agente.csv`, so generate it first with
  `python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv`
  (a committed offline sample), then
  `python3 -m streamlit run app.py --server.port 8501 --server.headless true`.
  The four top metric cards always show full-dataset totals and intentionally do
  not react to the sidebar search filter (only the chart/table do).
- Live SEC EDGAR / BCB / BNDES calls require network egress and an identifiable
  `User-Agent` (e.g. `--user-agent "Name email@domain"`), which may be blocked
  in the cloud VM. Prefer offline demo mode / committed caches / the sample
  CSVs; the BNDES generators need `--download` (network) or a local
  `--excel`/`--input` file to produce output.
