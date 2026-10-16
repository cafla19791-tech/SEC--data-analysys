# SEC--data-analysys

Python 3.12 project for Brazilian public-finance analysis. It contains several
self-contained CLI/data-pipeline products plus one Streamlit web UI:

- `analyze_finance.py` — generic company financial-analyzer CLI over SEC EDGAR
  XBRL CompanyFacts, with an offline `--demo` mode. Branch:
  `cursor/analise-financeira-empresas-*`.
- `extract_petrobras_net_income.py` — Petrobras net-income extractor (SEC EDGAR
  + BCB USD/BRL), with committed data caches. Branch:
  `cursor/petrobras-lucro-liquido-*`.
- `scripts/gerar_fluxos*.py` — BNDES cash-flow / subsidy generators. Branches:
  `cursor/fluxos-*` and `cursor/financial-flows-processing-*`.
- `app.py` — Streamlit dashboard ("Resumo por Agente Financeiro") on top of the
  BNDES generator output. Branch: `cursor/resumo-agente-financeiro-*`.

## Cursor Cloud specific instructions

- The `main` branch is intentionally minimal (only `README.md` and this
  `AGENTS.md`). The actual applications live on feature branches, each with its
  own `requirements.txt`. When starting work, base your branch on `main` and
  bring in code from the relevant feature branch / PR — e.g.
  `git worktree add /tmp/app origin/<branch>` and work there.
- Dependencies are installed with `pip --break-system-packages` (Ubuntu 24.04 is
  PEP 668 "externally managed"; no virtualenv system package is guaranteed). The
  startup update script runs
  `pip install --break-system-packages -r requirements.txt` only when a
  `requirements.txt` exists, so on a bare `main` checkout it is a no-op. After
  checking out a feature branch, run that branch's
  `pip install --break-system-packages -r requirements.txt` (or re-run the
  update script from that branch's root) to pull its deps.
- `pip --break-system-packages` installs console scripts (e.g. `streamlit`,
  `pytest`) into `~/.local/bin`, which is not on `PATH` by default. Either
  prefix with `export PATH="$HOME/.local/bin:$PATH"` or invoke via
  `python3 -m` (e.g. `python3 -m pytest`, `python3 -m streamlit`).
- Run tests with `python3 -m pytest -q` from the branch root. The BNDES branches
  keep code under `scripts/` and need `PYTHONPATH=. python3 -m pytest -q`.
- Run the SEC analyzer offline (no network): `python3 analyze_finance.py --demo`.
- Run the Petrobras extractor offline: `python3 extract_petrobras_net_income.py
  --years 10` (uses the committed `data/` caches; only `--refresh` hits the
  network).
- BNDES generator (offline sample):
  `python3 scripts/gerar_fluxos.py --input data/sample_operacoes_com_agente.csv`
  writes `output/resumo_por_agente.csv` etc. The Streamlit UI reads that output,
  so run the generator before `streamlit run app.py` (a sample output is also
  committed on the Streamlit branch, so the dashboard renders immediately).
- Streamlit dev server: `streamlit run app.py --server.headless true
  --server.port 8501` (default port 8501). `--server.headless true` avoids the
  interactive "enter your email" prompt in the cloud VM.
- Live SEC EDGAR / BCB / BNDES calls require network access and an identifiable
  `User-Agent` (e.g. `--user-agent "Name email@domain"`). Network egress may be
  blocked; prefer offline demo mode or committed caches to run end-to-end
  without hitting external APIs.
