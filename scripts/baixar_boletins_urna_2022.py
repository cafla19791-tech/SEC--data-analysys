#!/usr/bin/env python3
"""
Baixa os Boletins de Urna (BU) do 2º turno de 2022 nas 28 UFs
(27 estados + ZZ/exterior) e monta uma tabela nacional com
número de série e modelo de cada urna + votos para Presidente.

O TSE não publica um ZIP único de BU. Este script percorre as 28 UFs,
baixa cada arquivo oficial e junta o resultado.

Fontes oficiais (Dados Abertos do TSE):
  - Boletim de Urna 2º turno (por UF):
      https://dadosabertos.tse.jus.br/dataset/resultados-2022-boletim-de-urna
      cdn: .../buweb/bweb_2t_{UF}_311020221535.zip
  - Correspondência número interno × modelo:
      https://dadosabertos.tse.jus.br/dataset/correspondencia-entre-numero-interno-e-modelo-da-urna-1
      cdn: .../modelo_urna/modelourna_numerointerno.zip

Uso (repo):
  python3 scripts/baixar_boletins_urna_2022.py
  python3 scripts/baixar_boletins_urna_2022.py --ufs RR AC
  python3 scripts/baixar_boletins_urna_2022.py --somente-processar

Uso (ContAgil WinPython — duplo-clique):
  baixar_zips_urna_curl.bat
  python baixar_boletins_urna_2022.py --usar-curl --workers 1
  python baixar_boletins_urna_2022.py --somente-processar
"""

from __future__ import annotations

import argparse
import io
import os
import shutil
import subprocess
import sys
import time
import zipfile
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

import pandas as pd
import requests

UFS = (
    "AC",
    "AL",
    "AM",
    "AP",
    "BA",
    "CE",
    "DF",
    "ES",
    "GO",
    "MA",
    "MG",
    "MS",
    "MT",
    "PA",
    "PB",
    "PE",
    "PI",
    "PR",
    "RJ",
    "RN",
    "RO",
    "RR",
    "RS",
    "SC",
    "SE",
    "SP",
    "TO",
    "ZZ",
)

BWEB_2T_URL = (
    "https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/"
    "buweb/bweb_2t_{uf}_311020221535.zip"
)
MODELO_URL = (
    "https://cdn.tse.jus.br/estatistica/sead/odsele/modelo_urna/"
    "modelourna_numerointerno.zip"
)
# O CDN do TSE costuma devolver 403 para scripts. O Internet Archive
# guardou os ZIPs oficiais (captura 2022-11-08). 2023id_ resolve a cópia
# mais recente; o timestamp fixo evita o 302 que às vezes devolve 503.
WAYBACK_TS_PREFIX = "https://web.archive.org/web/20221108000702id_/"
WAYBACK_ID_PREFIX = "https://web.archive.org/web/2023id_/"
# Chrome UA no Archive.org devolve 503; no TSE ajuda a passar o WAF.
USER_AGENT_ARQUIVO = "ContAgil-TSE-BU/1.0"

# Faixas oficiais TSE (STI/COTEL): número interno → modelo.
# Mesmo conteúdo de modelourna_numerointerno.csv (publicado em 05/11/2022).
FAIXAS_MODELO_OFICIAL = (
    (2009, 999_500, 1_220_500),
    (2010, 1_220_501, 1_345_500),
    (2011, 1_368_501, 1_370_500),
    (2011, 1_600_000, 1_650_000),
    (2013, 1_650_001, 1_701_000),
    (2015, 1_750_000, 1_950_000),
    (2020, 2_000_000, 2_250_000),
)

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept": "application/zip,application/octet-stream,*/*",
    "Referer": "https://dadosabertos.tse.jus.br/",
}

BWEB_COLS = (
    "NR_TURNO",
    "SG_UF",
    "CD_MUNICIPIO",
    "NM_MUNICIPIO",
    "NR_ZONA",
    "NR_SECAO",
    "NR_LOCAL_VOTACAO",
    "CD_CARGO_PERGUNTA",
    "DS_CARGO_PERGUNTA",
    "CD_TIPO_VOTAVEL",
    "DS_TIPO_VOTAVEL",
    "NR_VOTAVEL",
    "NM_VOTAVEL",
    "QT_VOTOS",
    "NR_URNA_EFETIVADA",
    "QT_APTOS",
    "QT_COMPARECIMENTO",
    "QT_ABSTENCOES",
    "DT_ABERTURA",
    "DT_ENCERRAMENTO",
    "DS_TIPO_URNA",
)

CHAVES_URNA = (
    "SG_UF",
    "CD_MUNICIPIO",
    "NR_ZONA",
    "NR_SECAO",
    "NR_URNA_EFETIVADA",
)

REPO_ROOT = Path(__file__).resolve().parents[1]
CONTAGIL_WINPYTHON = Path(
    r"C:\Arquivos de Programas RFB\ContAgilAppBeta64\python_jep\winpython"
)


def _configure_stdio() -> None:
    for stream in (sys.stdout, sys.stderr):
        reconf = getattr(stream, "reconfigure", None)
        if reconf is None:
            continue
        try:
            reconf(encoding="utf-8", errors="replace")
        except Exception:
            pass


def _parece_winpython(pasta: Path) -> bool:
    try:
        pasta = pasta.resolve()
    except OSError:
        return False
    texto = str(pasta).upper().replace("/", "\\")
    return (pasta / "python.exe").exists() and (
        "WINPYTHON" in texto or "CONTAGIL" in texto
    )


def descobrir_winpython(*extras: Path) -> Path | None:
    """Localiza a pasta WinPython do ContAgil (python.exe + winpython/contagil)."""
    here = Path(__file__).resolve()
    candidatos = [
        *extras,
        Path.cwd(),
        here.parent,
        here.parent.parent,
        CONTAGIL_WINPYTHON,
    ]
    vistos: set[Path] = set()
    for cand in candidatos:
        if cand is None:
            continue
        for pasta in (Path(cand), Path(cand).parent):
            try:
                chave = pasta.resolve()
            except OSError:
                continue
            if chave in vistos:
                continue
            vistos.add(chave)
            if _parece_winpython(chave):
                return chave
    return None


def pastas_padrao(winpy: Path | None = None) -> tuple[Path, Path]:
    """ZIPs em dados/tse2022/raw e CSVs em saida/tse2022 no ContAgil."""
    raiz = winpy if winpy is not None else descobrir_winpython()
    if raiz is not None:
        return raiz / "dados" / "tse2022" / "raw", raiz / "saida" / "tse2022"
    return REPO_ROOT / "data" / "tse2022" / "raw", REPO_ROOT / "output" / "tse2022"


def _pasta_raw_de_massa(massa: Path) -> Path:
    nome = massa.name.lower()
    if nome == "raw":
        return massa
    if nome == "tse2022":
        return massa / "raw"
    return massa / "tse2022" / "raw"


def _pasta_saida_contagil(pasta: Path) -> Path:
    return pasta if pasta.name.lower() == "tse2022" else pasta / "tse2022"


def resolver_pastas(args: argparse.Namespace) -> tuple[Path, Path]:
    raw_default, saida_default = pastas_padrao()
    raw = args.raw_dir
    saida = args.saida
    if getattr(args, "massa_dados", None):
        raw = _pasta_raw_de_massa(Path(args.massa_dados))
    if getattr(args, "pasta_saida", None):
        saida = _pasta_saida_contagil(Path(args.pasta_saida))
    if raw is None:
        raw = raw_default
    if saida is None:
        saida = saida_default
    return Path(raw), Path(saida)


def parse_args(argv: list[str] | None = None) -> argparse.Namespace:
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument(
        "--ufs",
        nargs="+",
        default=list(UFS),
        help="UFs a baixar (padrão: as 28, incluindo ZZ).",
    )
    p.add_argument(
        "--raw-dir",
        type=Path,
        default=None,
        help="Pasta dos ZIPs baixados (default: ContAgil dados/tse2022/raw ou data/tse2022/raw).",
    )
    p.add_argument(
        "--saida",
        type=Path,
        default=None,
        help="Pasta dos CSVs gerados (default: ContAgil saida/tse2022 ou output/tse2022).",
    )
    p.add_argument(
        "--massa-dados",
        "--pasta-dados",
        type=Path,
        default=None,
        dest="massa_dados",
        help="ContAgil: pasta dados (ZIPs em dados/tse2022/raw).",
    )
    p.add_argument(
        "--pasta-saida",
        type=Path,
        default=None,
        help="ContAgil: pasta saida (CSVs em saida/tse2022).",
    )
    p.add_argument(
        "--modelo",
        type=Path,
        default=None,
        help="CSV/ZIP oficial de faixas de modelo (opcional).",
    )
    p.add_argument(
        "--somente-baixar",
        action="store_true",
        help="Só baixa os ZIPs, não processa.",
    )
    p.add_argument(
        "--somente-processar",
        action="store_true",
        help="Usa ZIPs já baixados em --raw-dir.",
    )
    p.add_argument(
        "--workers",
        type=int,
        default=1,
        help="Downloads em paralelo (padrão: 1; no ContAgil evite >1).",
    )
    p.add_argument(
        "--timeout",
        type=int,
        default=300,
        help="Timeout HTTP por arquivo, em segundos.",
    )
    p.add_argument(
        "--tentativas",
        type=int,
        default=2,
        help="Tentativas do Python por URL (padrão: 2). SSL/403 vai direto ao curl.",
    )
    p.add_argument(
        "--gerar-links",
        action="store_true",
        help="Só gera o HTML com links TSE + Archive.org e sai.",
    )
    p.add_argument(
        "--usar-curl",
        action="store_true",
        help="Baixa só com curl.exe (SChannel do Windows). Recomendado na RFB.",
    )
    return p.parse_args(argv)


def normalizar_ufs(ufs: list[str]) -> list[str]:
    out: list[str] = []
    seen: set[str] = set()
    for raw in ufs:
        uf = raw.strip().upper()
        if uf not in UFS:
            raise ValueError(f"UF inválida: {raw!r}. Use uma de: {', '.join(UFS)}")
        if uf not in seen:
            out.append(uf)
            seen.add(uf)
    return out


def url_bweb(uf: str) -> str:
    return BWEB_2T_URL.format(uf=uf.upper())


def urls_espelho(url: str) -> list[str]:
    """TSE oficial primeiro; Internet Archive se o CDN bloquear."""
    if url.startswith("https://web.archive.org/"):
        return [url]
    return [
        url,
        f"{WAYBACK_TS_PREFIX}{url}",
        f"{WAYBACK_ID_PREFIX}{url}",
    ]


def user_agent_para(url: str) -> str:
    if "web.archive.org" in url:
        return USER_AGENT_ARQUIVO
    return HEADERS["User-Agent"]


def cabecalhos_para(url: str) -> dict[str, str]:
    headers = dict(HEADERS)
    headers["User-Agent"] = user_agent_para(url)
    if "web.archive.org" in url:
        headers["Referer"] = "https://web.archive.org/"
    return headers


def escrever_pagina_links(destino: Path) -> Path:
    """HTML com os 28 ZIPs (TSE + Archive.org) para abrir no navegador."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    linhas = [
        "<!DOCTYPE html><html lang='pt-BR'><head><meta charset='utf-8'>",
        "<title>Boletins de Urna 2022 — download</title></head><body>",
        "<h1>Boletins de Urna 2022 (2º turno)</h1>",
        "<p>Na RFB o Python costuma falhar no TLS do Archive.org. "
        "Prefira o <code>baixar_zips_urna_curl.bat</code> (usa curl.exe do Windows). "
        "Ou baixe no navegador e salve em <code>dados\\tse2022\\raw</code>, "
        "depois: <code>python baixar_boletins_urna_2022.py --somente-processar</code>.</p><ol>",
    ]
    for uf in UFS:
        oficial = url_bweb(uf)
        archive = f"{WAYBACK_ID_PREFIX}{oficial}"
        nome = f"bweb_2t_{uf}_311020221535.zip"
        linhas.append(
            f"<li><b>{uf}</b> — "
            f"<a href='{oficial}'>{nome}</a> · "
            f"<a href='{archive}'>espelho Archive.org</a></li>"
        )
    linhas.append("</ol></body></html>")
    destino.write_text("\n".join(linhas), encoding="utf-8")
    return destino


def e_erro_ssl(exc: BaseException) -> bool:
    """WinPython/OpenSSL da RFB não fecha handshake com archive.org."""
    if isinstance(exc, requests.exceptions.SSLError):
        return True
    causa = getattr(exc, "__cause__", None)
    if causa is not None and "SSL" in type(causa).__name__:
        return True
    msg = str(exc).lower()
    return any(
        trecho in msg
        for trecho in (
            "sslv3",
            "ssl:",
            "ssl error",
            "handshake failure",
            "certificate verify",
            "sslerror",
        )
    )


def e_erro_irrecuperavel_python(exc: BaseException) -> bool:
    """403, TLS quebrado ou timeout: retry do requests não resolve."""
    status = getattr(getattr(exc, "response", None), "status_code", None)
    if status == 403:
        return True
    if e_erro_ssl(exc):
        return True
    msg = str(exc).lower()
    return "timed out" in msg or "timeout" in msg


def resumir_erro_download(exc: BaseException) -> str:
    status = getattr(getattr(exc, "response", None), "status_code", None)
    if status:
        return f"HTTP {status}"
    if e_erro_ssl(exc):
        return "falha TLS (Python/OpenSSL; use curl.exe)"
    msg = str(exc)
    if "timed out" in msg.lower() or "timeout" in msg.lower():
        return "timeout"
    primeira = msg.splitlines()[0].strip()
    return primeira[:160]


def preferir_curl(url: str, plataforma: str | None = None) -> bool:
    """No Windows, Archive.org via curl.exe (SChannel) evita o bug de TLS."""
    plat = plataforma if plataforma is not None else sys.platform
    return plat == "win32" and "web.archive.org" in url


def encontrar_curl() -> str | None:
    if sys.platform == "win32":
        system32 = (
            Path(os.environ.get("SystemRoot", r"C:\Windows")) / "System32" / "curl.exe"
        )
        if system32.is_file():
            return str(system32)
    for nome in ("curl.exe", "curl"):
        achado = shutil.which(nome)
        if achado:
            return achado
    return None


def montar_comando_curl(
    url: str,
    destino: Path,
    *,
    timeout: int,
    insecure: bool = False,
    curl: str = "curl",
    plataforma: str | None = None,
) -> list[str]:
    plat = plataforma if plataforma is not None else sys.platform
    cmd = [
        curl,
        "-L",
        "--fail",
        "--retry",
        "2",
        "--retry-delay",
        "4",
        "--connect-timeout",
        "45",
        "--max-time",
        str(max(60, timeout)),
        "-A",
        user_agent_para(url),
    ]
    if plat == "win32":
        cmd.append("--ssl-no-revoke")
    if insecure:
        cmd.append("-k")
    cmd.extend(["-o", str(destino), "--", url])
    return cmd


def baixar_com_curl(
    url: str,
    destino: Path,
    *,
    timeout: int = 300,
    curl: str | None = None,
) -> Path:
    """Baixa com curl.exe (SChannel no Windows). Evita o OpenSSL do WinPython."""
    exe = curl or encontrar_curl()
    if not exe:
        raise RuntimeError(
            "curl.exe nao encontrado. Na RFB use C:\\Windows\\System32\\curl.exe "
            "ou rode baixar_zips_urna_curl.bat."
        )
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists():
        destino.unlink()
    tentativas_cmd = [
        montar_comando_curl(url, destino, timeout=timeout, curl=exe),
    ]
    if sys.platform == "win32":
        tentativas_cmd.append(
            montar_comando_curl(url, destino, timeout=timeout, curl=exe, insecure=True)
        )
    last_err: Exception | None = None
    for i, cmd in enumerate(tentativas_cmd):
        extra = " (-k)" if i and "-k" in cmd else ""
        print(f"    CURL{extra} {url}", flush=True)
        try:
            proc = subprocess.run(
                cmd,
                capture_output=True,
                text=True,
                timeout=timeout + 60,
            )
        except subprocess.TimeoutExpired:
            last_err = RuntimeError(f"curl timeout: {url}")
            if destino.exists():
                destino.unlink(missing_ok=True)
            continue
        if proc.returncode == 0 and destino.exists() and destino.stat().st_size >= 100:
            return destino
        if destino.exists():
            destino.unlink(missing_ok=True)
        detalhe = (proc.stderr or proc.stdout or "").strip().replace("\n", " ")
        last_err = RuntimeError(
            f"curl {proc.returncode}: {detalhe[-240:] or 'sem saida'}"
        )
    raise last_err or RuntimeError(f"curl falhou: {url}")


def escrever_script_curl(destino: Path) -> Path:
    """Gera um .bat que baixa os 28 ZIPs com curl.exe do Windows."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    linhas = [
        "@echo off",
        "REM Gerado pelo script — usa curl.exe do Windows (SChannel).",
        "setlocal EnableExtensions",
        'cd /d "%~dp0"',
        'if not exist "%CD%\\dados" if exist "%CD%\\..\\..\\dados" cd /d "%CD%\\..\\.."',
        'if not exist "%CD%\\dados" if exist "%CD%\\..\\dados" cd /d "%CD%\\.."',
        'set "RAW=%CD%\\dados\\tse2022\\raw"',
        'if not exist "%RAW%" mkdir "%RAW%"',
        f'set "UA_TSE={HEADERS["User-Agent"]}"',
        f'set "UA_IA={USER_AGENT_ARQUIVO}"',
        'set "TSE=https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb"',
        f'set "IA={WAYBACK_TS_PREFIX}https://cdn.tse.jus.br/estatistica/sead/eleicoes/eleicoes2022/buweb"',
        'set "CURL=%SystemRoot%\\System32\\curl.exe"',
        'if not exist "%CURL%" set "CURL=curl.exe"',
        "echo Destino: %RAW%",
        "echo curl: %CURL%",
        "echo.",
    ]
    for uf in UFS:
        nome = f"bweb_2t_{uf}_311020221535.zip"
        linhas.append(f'call :BAIXA "{uf}" "{nome}"')
    linhas.extend(
        [
            "echo.",
            "echo Processando ZIPs...",
            "if exist baixar_boletins_urna_2022.py (",
            '  python baixar_boletins_urna_2022.py --somente-processar --massa-dados "%CD%\\dados" --pasta-saida "%CD%\\saida"',
            ") else (",
            "  echo Rode: python baixar_boletins_urna_2022.py --somente-processar",
            ")",
            "goto :FIM",
            "",
            ":BAIXA",
            'set "UF=%~1"',
            'set "NOME=%~2"',
            'set "DEST=%RAW%\\%NOME%"',
            'if exist "%DEST%" (',
            '  for %%A in ("%DEST%") do if %%~zA GTR 1000 (',
            "    echo [ok] %UF% ja existe",
            "    goto :EOF",
            "  )",
            ")",
            "echo [TSE] %UF%",
            '"%CURL%" -L --fail --retry 2 --connect-timeout 45 --max-time 600 --ssl-no-revoke -A "%UA_TSE%" -o "%DEST%.part" "%TSE%/%NOME%"',
            "if not errorlevel 1 (",
            '  move /Y "%DEST%.part" "%DEST%" >nul',
            "  echo [ok] %UF% via TSE",
            "  goto :EOF",
            ")",
            "echo [IA] %UF%",
            '"%CURL%" -L --fail --retry 2 --connect-timeout 45 --max-time 600 --ssl-no-revoke -k -A "%UA_IA%" -o "%DEST%.part" "%IA%/%NOME%"',
            "if not errorlevel 1 (",
            '  move /Y "%DEST%.part" "%DEST%" >nul',
            "  echo [ok] %UF% via Archive.org",
            "  goto :EOF",
            ")",
            "echo [ERRO] %UF%",
            'if exist "%DEST%.part" del "%DEST%.part"',
            "goto :EOF",
            "",
            ":FIM",
            "endlocal",
        ]
    )
    destino.write_text("\r\n".join(linhas) + "\r\n", encoding="utf-8")
    return destino


def baixar_arquivo(
    url: str,
    destino: Path,
    *,
    timeout: int = 180,
    tentativas: int = 2,
    session: requests.Session | None = None,
    usar_curl: bool = False,
) -> Path:
    """Baixa um arquivo oficial do TSE; se falhar, Archive.org; se TLS falhar, curl.exe."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    if destino.exists() and destino.stat().st_size > 0:
        return destino

    http = session or requests.Session()
    last_err: Exception | None = None
    tmp = destino.with_suffix(destino.suffix + ".part")

    def _via_curl(candidato: str) -> bool:
        nonlocal last_err
        try:
            baixar_com_curl(candidato, tmp, timeout=timeout)
            if tmp.stat().st_size < 100:
                raise RuntimeError(f"Download vazio: {candidato}")
            tmp.replace(destino)
            return True
        except Exception as exc:
            last_err = exc
            if tmp.exists():
                tmp.unlink(missing_ok=True)
            print(f"    curl falhou: {resumir_erro_download(exc)}", flush=True)
            return False

    for candidato in urls_espelho(url):
        if usar_curl or preferir_curl(candidato):
            if _via_curl(candidato):
                return destino
            if usar_curl:
                continue

        for tentativa in range(max(1, tentativas)):
            try:
                print(f"    GET {candidato}", flush=True)
                with http.get(
                    candidato,
                    headers=cabecalhos_para(candidato),
                    stream=True,
                    timeout=timeout,
                ) as resp:
                    resp.raise_for_status()
                    with tmp.open("wb") as fh:
                        for chunk in resp.iter_content(chunk_size=1024 * 256):
                            if chunk:
                                fh.write(chunk)
                if tmp.stat().st_size < 100:
                    raise RuntimeError(f"Download vazio: {candidato}")
                tmp.replace(destino)
                return destino
            except Exception as exc:
                last_err = exc
                if tmp.exists():
                    tmp.unlink(missing_ok=True)
                status = getattr(getattr(exc, "response", None), "status_code", None)
                if status == 403:
                    print(f"    bloqueado (403): {candidato}", flush=True)
                    break
                if e_erro_ssl(exc):
                    print(
                        "    TLS do Python falhou; tentando curl.exe do Windows...",
                        flush=True,
                    )
                    break
                if tentativa + 1 >= tentativas or e_erro_irrecuperavel_python(exc):
                    break
                time.sleep(4 * (2**tentativa))

        if _via_curl(candidato):
            return destino

    raise RuntimeError(
        f"Falha ao baixar {url}: {resumir_erro_download(last_err or RuntimeError('erro'))}"
    ) from last_err


def ler_csv_tse(fonte: Path | io.BytesIO | io.StringIO) -> pd.DataFrame:
    return pd.read_csv(fonte, sep=";", encoding="latin-1", dtype=str, low_memory=False)


def csv_dentro_do_zip(zip_path: Path) -> bytes:
    with zipfile.ZipFile(zip_path) as zf:
        nomes = [n for n in zf.namelist() if n.lower().endswith(".csv")]
        if not nomes:
            raise FileNotFoundError(f"Nenhum CSV em {zip_path}")
        # Prefere o arquivo de dados, não o leiame.
        nomes.sort(key=lambda n: (0 if "leiame" in n.lower() else 1, n))
        escolhido = nomes[-1]
        return zf.read(escolhido)


def carregar_faixas_modelo(path: Path | None = None) -> pd.DataFrame:
    """Lê o CSV/ZIP oficial ou usa as faixas publicadas pelo TSE."""
    if path is not None and path.exists():
        if path.suffix.lower() == ".zip":
            raw = csv_dentro_do_zip(path)
            df = ler_csv_tse(io.BytesIO(raw))
        else:
            df = ler_csv_tse(path)
        cols = {c.strip().upper(): c for c in df.columns}
        modelo_col = next(
            (cols[k] for k in ("DS_MODELO_URNA", "DS_MODELO", "MODELO") if k in cols),
            None,
        )
        ini_col = next(
            (
                cols[k]
                for k in (
                    "NR_FAIXA_INICIAL",
                    "NR_INICIAL",
                    "NR_FABRICACAO_INICIAL",
                    "NR_NUMERO_INTERNO_INICIAL",
                )
                if k in cols
            ),
            None,
        )
        fim_col = next(
            (
                cols[k]
                for k in (
                    "NR_FAIXA_FINAL",
                    "NR_FINAL",
                    "NR_FABRICACAO_FINAL",
                    "NR_NUMERO_INTERNO_FINAL",
                )
                if k in cols
            ),
            None,
        )
        if modelo_col and ini_col and fim_col:
            out = pd.DataFrame(
                {
                    "NR_MODELO": pd.to_numeric(df[modelo_col], errors="coerce"),
                    "NR_FAIXA_INICIAL": pd.to_numeric(df[ini_col], errors="coerce"),
                    "NR_FAIXA_FINAL": pd.to_numeric(df[fim_col], errors="coerce"),
                }
            ).dropna()
            out["NR_MODELO"] = out["NR_MODELO"].astype(int)
            out["NR_FAIXA_INICIAL"] = out["NR_FAIXA_INICIAL"].astype(int)
            out["NR_FAIXA_FINAL"] = out["NR_FAIXA_FINAL"].astype(int)
            return out.reset_index(drop=True)

    return pd.DataFrame(
        FAIXAS_MODELO_OFICIAL,
        columns=["NR_MODELO", "NR_FAIXA_INICIAL", "NR_FAIXA_FINAL"],
    )


def rotulo_modelo(nr_modelo: int | float | None) -> str:
    if nr_modelo is None or pd.isna(nr_modelo):
        return "sem_faixa"
    return f"UE{int(nr_modelo)}"


def classificar_modelo(numeros: pd.Series, faixas: pd.DataFrame) -> pd.Series:
    """Associa NR_URNA_EFETIVADA ao modelo pela faixa oficial do TSE."""
    nums = pd.to_numeric(numeros, errors="coerce")
    modelo = pd.Series(pd.NA, index=numeros.index, dtype="Int64")
    for row in faixas.itertuples(index=False):
        mask = nums.between(int(row.NR_FAIXA_INICIAL), int(row.NR_FAIXA_FINAL))
        modelo.loc[mask] = int(row.NR_MODELO)
    return modelo


def filtrar_presidente_2t(df: pd.DataFrame) -> pd.DataFrame:
    trabalho = df.copy()
    trabalho.columns = [c.strip() for c in trabalho.columns]
    if "NR_TURNO" in trabalho.columns:
        turno = pd.to_numeric(trabalho["NR_TURNO"], errors="coerce")
        trabalho = trabalho.loc[turno == 2]
    if "CD_CARGO_PERGUNTA" in trabalho.columns:
        cargo = pd.to_numeric(trabalho["CD_CARGO_PERGUNTA"], errors="coerce")
        por_codigo = cargo == 1
        if por_codigo.any():
            trabalho = trabalho.loc[por_codigo]
        elif "DS_CARGO_PERGUNTA" in trabalho.columns:
            trabalho = trabalho.loc[
                trabalho["DS_CARGO_PERGUNTA"].astype(str).str.contains(
                    "Presidente", case=False, na=False
                )
            ]
    elif "DS_CARGO_PERGUNTA" in trabalho.columns:
        trabalho = trabalho.loc[
            trabalho["DS_CARGO_PERGUNTA"].astype(str).str.contains(
                "Presidente", case=False, na=False
            )
        ]
    return trabalho


def consolidar_urnas(df: pd.DataFrame, faixas: pd.DataFrame) -> pd.DataFrame:
    """Uma linha por urna/seção, com série, modelo e votos de Presidente."""
    base = filtrar_presidente_2t(df)
    if base.empty:
        return pd.DataFrame()

    for col in (
        "QT_VOTOS",
        "QT_APTOS",
        "QT_COMPARECIMENTO",
        "QT_ABSTENCOES",
        "NR_URNA_EFETIVADA",
        "NR_VOTAVEL",
        "CD_TIPO_VOTAVEL",
    ):
        if col in base.columns:
            if col.startswith("QT_") or col == "NR_URNA_EFETIVADA":
                base[col] = pd.to_numeric(base[col], errors="coerce")

    tipo = base.get("DS_TIPO_VOTAVEL", pd.Series("", index=base.index)).astype(str)
    nr_vot = pd.to_numeric(base.get("NR_VOTAVEL"), errors="coerce")
    cd_tipo = pd.to_numeric(base.get("CD_TIPO_VOTAVEL"), errors="coerce")
    nm = base.get("NM_VOTAVEL", pd.Series("", index=base.index)).astype(str).str.upper()

    base["QT_VOTOS_LULA"] = 0
    base["QT_VOTOS_BOLSONARO"] = 0
    base["QT_VOTOS_BRANCO"] = 0
    base["QT_VOTOS_NULO"] = 0

    votos = pd.to_numeric(base["QT_VOTOS"], errors="coerce").fillna(0)
    lula = (nr_vot == 13) | nm.str.contains("LULA", na=False)
    bolso = (nr_vot == 22) | nm.str.contains("BOLSONARO", na=False)
    branco = (
        tipo.str.contains("BRANCO", case=False, na=False)
        | (cd_tipo == 2)
        | (nr_vot == 95)
        | nm.str.contains("BRANCO", na=False)
    )
    nulo = (
        tipo.str.contains("NULO", case=False, na=False)
        | (cd_tipo == 3)
        | (nr_vot == 96)
        | nm.str.contains("NULO", na=False)
    )
    base.loc[lula, "QT_VOTOS_LULA"] = votos.loc[lula]
    base.loc[bolso, "QT_VOTOS_BOLSONARO"] = votos.loc[bolso]
    base.loc[branco, "QT_VOTOS_BRANCO"] = votos.loc[branco]
    base.loc[nulo, "QT_VOTOS_NULO"] = votos.loc[nulo]

    chaves = [c for c in CHAVES_URNA if c in base.columns]
    extras = [
        c
        for c in (
            "NM_MUNICIPIO",
            "NR_LOCAL_VOTACAO",
            "QT_APTOS",
            "QT_COMPARECIMENTO",
            "QT_ABSTENCOES",
            "DT_ABERTURA",
            "DT_ENCERRAMENTO",
            "DS_TIPO_URNA",
        )
        if c in base.columns
    ]
    agrupado = (
        base.groupby(chaves, dropna=False)
        .agg(
            {
                **{c: "first" for c in extras},
                "QT_VOTOS_LULA": "sum",
                "QT_VOTOS_BOLSONARO": "sum",
                "QT_VOTOS_BRANCO": "sum",
                "QT_VOTOS_NULO": "sum",
            }
        )
        .reset_index()
    )

    agrupado["NR_MODELO"] = classificar_modelo(agrupado["NR_URNA_EFETIVADA"], faixas)
    agrupado["DS_MODELO_URNA"] = agrupado["NR_MODELO"].map(rotulo_modelo)
    agrupado["QT_VOTOS_VALIDOS"] = (
        agrupado["QT_VOTOS_LULA"] + agrupado["QT_VOTOS_BOLSONARO"]
    )
    ordem = [
        "SG_UF",
        "CD_MUNICIPIO",
        "NM_MUNICIPIO",
        "NR_ZONA",
        "NR_SECAO",
        "NR_LOCAL_VOTACAO",
        "NR_URNA_EFETIVADA",
        "NR_MODELO",
        "DS_MODELO_URNA",
        "QT_APTOS",
        "QT_COMPARECIMENTO",
        "QT_ABSTENCOES",
        "QT_VOTOS_LULA",
        "QT_VOTOS_BOLSONARO",
        "QT_VOTOS_BRANCO",
        "QT_VOTOS_NULO",
        "QT_VOTOS_VALIDOS",
        "DT_ABERTURA",
        "DT_ENCERRAMENTO",
        "DS_TIPO_URNA",
    ]
    return agrupado[[c for c in ordem if c in agrupado.columns]]


def processar_zip_bweb(zip_path: Path, faixas: pd.DataFrame) -> pd.DataFrame:
    raw = csv_dentro_do_zip(zip_path)
    df = pd.read_csv(
        io.BytesIO(raw),
        sep=";",
        encoding="latin-1",
        dtype=str,
        usecols=lambda c: c in BWEB_COLS,
        low_memory=False,
    )
    return consolidar_urnas(df, faixas)


def resumo_por_uf(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby("SG_UF", dropna=False)
        .agg(
            QT_URNAS=("NR_URNA_EFETIVADA", "nunique"),
            QT_SECOES=("NR_SECAO", "size"),
            QT_VOTOS_LULA=("QT_VOTOS_LULA", "sum"),
            QT_VOTOS_BOLSONARO=("QT_VOTOS_BOLSONARO", "sum"),
            QT_VOTOS_BRANCO=("QT_VOTOS_BRANCO", "sum"),
            QT_VOTOS_NULO=("QT_VOTOS_NULO", "sum"),
            QT_VOTOS_VALIDOS=("QT_VOTOS_VALIDOS", "sum"),
        )
        .reset_index()
        .sort_values("SG_UF")
    )
    return g


def resumo_por_modelo(df: pd.DataFrame) -> pd.DataFrame:
    g = (
        df.groupby(["NR_MODELO", "DS_MODELO_URNA"], dropna=False)
        .agg(
            QT_URNAS=("NR_URNA_EFETIVADA", "nunique"),
            QT_VOTOS_LULA=("QT_VOTOS_LULA", "sum"),
            QT_VOTOS_BOLSONARO=("QT_VOTOS_BOLSONARO", "sum"),
            QT_VOTOS_VALIDOS=("QT_VOTOS_VALIDOS", "sum"),
        )
        .reset_index()
        .sort_values("NR_MODELO")
    )
    g["PCT_LULA"] = (g["QT_VOTOS_LULA"] / g["QT_VOTOS_VALIDOS"] * 100).round(2)
    g["PCT_BOLSONARO"] = (g["QT_VOTOS_BOLSONARO"] / g["QT_VOTOS_VALIDOS"] * 100).round(2)
    return g


def baixar_todos(
    ufs: list[str],
    raw_dir: Path,
    *,
    workers: int,
    timeout: int,
    tentativas: int,
    baixar_modelo: bool = True,
    modelo_path: Path | None = None,
    usar_curl: bool = False,
) -> dict[str, Path]:
    raw_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    tarefas: list[tuple[str, str, Path]] = []
    if baixar_modelo:
        dest_modelo = modelo_path or (raw_dir / "modelourna_numerointerno.zip")
        tarefas.append(("MODELO", MODELO_URL, dest_modelo))
    for uf in ufs:
        tarefas.append((uf, url_bweb(uf), raw_dir / f"bweb_2t_{uf}_311020221535.zip"))

    ok: dict[str, Path] = {}
    erros: list[str] = []

    def _job(item: tuple[str, str, Path]) -> tuple[str, Path]:
        chave, url, dest = item
        print(f"  baixando {chave}: {url}", flush=True)
        path = baixar_arquivo(
            url,
            dest,
            timeout=timeout,
            tentativas=tentativas,
            session=session,
            usar_curl=usar_curl,
        )
        print(f"  ok {chave} ({path.stat().st_size:,} bytes)", flush=True)
        return chave, path

    with ThreadPoolExecutor(max_workers=max(1, workers)) as pool:
        futs = {pool.submit(_job, t): t[0] for t in tarefas}
        for fut in as_completed(futs):
            chave = futs[fut]
            try:
                k, path = fut.result()
                ok[k] = path
            except Exception as exc:
                if chave == "MODELO":
                    print(
                        "  aviso: não baixou o ZIP de modelos; "
                        "usando as faixas oficiais embutidas no script.",
                        flush=True,
                    )
                    print(f"         ({resumir_erro_download(exc)})", flush=True)
                    continue
                erros.append(f"{chave}: {resumir_erro_download(exc)}")
                print(f"  ERRO {chave}: {resumir_erro_download(exc)}", flush=True)

    if erros:
        raise RuntimeError(
            "Falha em um ou mais downloads do TSE: " + ", ".join(erros)
        )
    return ok


def main(argv: list[str] | None = None) -> int:
    _configure_stdio()
    args = parse_args(argv)
    ufs = normalizar_ufs(args.ufs)
    raw_dir, saida = resolver_pastas(args)
    raw_dir.mkdir(parents=True, exist_ok=True)
    saida.mkdir(parents=True, exist_ok=True)

    winpy = descobrir_winpython()
    if winpy is not None:
        print(f"ContAgil WinPython: {winpy}", flush=True)
    print(f"UFs ({len(ufs)}): {', '.join(ufs)}", flush=True)
    print(f"ZIPs : {raw_dir}", flush=True)
    print(f"Saida: {saida}", flush=True)

    if args.gerar_links:
        html = escrever_pagina_links(saida / "baixar_boletins_links.html")
        print(f"Abra no navegador: {html}")
        return 0

    if not args.somente_processar:
        try:
            baixar_todos(
                ufs,
                raw_dir,
                workers=args.workers,
                timeout=args.timeout,
                tentativas=args.tentativas,
                baixar_modelo=args.modelo is None,
                modelo_path=args.modelo
                if args.modelo and args.modelo.suffix == ".zip"
                else None,
                usar_curl=args.usar_curl,
            )
        except RuntimeError as exc:
            html = escrever_pagina_links(saida / "baixar_boletins_links.html")
            bat_curl = escrever_script_curl(saida / "baixar_zips_urna_curl.bat")
            print(
                "\nNao deu para baixar os ZIPs pelo Python.\n"
                "Na RFB o TSE devolve 403 e o OpenSSL do WinPython "
                "quebra o TLS com archive.org (handshake failure).\n"
                "Use o curl.exe do Windows (SChannel):\n"
                "  baixar_zips_urna_curl.bat\n"
                f"  (copia gerada em {bat_curl})\n"
                "Ou: python baixar_boletins_urna_2022.py --usar-curl --workers 1\n"
                f"Se o curl tambem falhar, abra no Edge:\n  {html}\n"
                f"Salve os ZIPs em:\n  {raw_dir}\n"
                "Depois: python baixar_boletins_urna_2022.py --somente-processar\n"
                f"Detalhe: {exc}\n",
                flush=True,
            )
            return 2

    if args.somente_baixar:
        print(f"Downloads em {raw_dir}")
        return 0

    faixas = carregar_faixas_modelo(args.modelo)
    print(
        f"Faixas de modelo: {len(faixas)} "
        f"({', '.join(sorted({rotulo_modelo(x) for x in faixas['NR_MODELO']}))})",
        flush=True,
    )

    partes: list[pd.DataFrame] = []
    faltando: list[str] = []
    for uf in ufs:
        zip_path = raw_dir / f"bweb_2t_{uf}_311020221535.zip"
        if not zip_path.exists():
            # aceita qualquer zip da UF já extraído/renomeado
            candidatos = sorted(raw_dir.glob(f"bweb_2t_{uf}_*.zip"))
            if not candidatos:
                faltando.append(uf)
                continue
            zip_path = candidatos[0]
        print(f"  processando {uf} ← {zip_path.name}", flush=True)
        partes.append(processar_zip_bweb(zip_path, faixas))

    if faltando:
        raise FileNotFoundError(
            "ZIP de BU ausente para: "
            + ", ".join(faltando)
            + f". Baixe com o script (sem --somente-processar) em {raw_dir}"
        )

    tabela = pd.concat(partes, ignore_index=True) if partes else pd.DataFrame()
    if tabela.empty:
        raise RuntimeError("Nenhuma urna processada. Confira os CSVs do TSE.")

    detalhe = saida / "urnas_2t_presidente.csv"
    por_uf = saida / "resumo_por_uf.csv"
    por_modelo = saida / "resumo_por_modelo.csv"
    tabela.to_csv(detalhe, index=False, encoding="utf-8")
    resumo_por_uf(tabela).to_csv(por_uf, index=False, encoding="utf-8")
    resumo_por_modelo(tabela).to_csv(por_modelo, index=False, encoding="utf-8")

    print()
    print(f"Urnas/seções: {len(tabela):,}".replace(",", "."))
    print(f"Série preenchida: {tabela['NR_URNA_EFETIVADA'].notna().sum():,}".replace(",", "."))
    print(f"Lula: {int(tabela['QT_VOTOS_LULA'].sum()):,}".replace(",", "."))
    print(f"Bolsonaro: {int(tabela['QT_VOTOS_BOLSONARO'].sum()):,}".replace(",", "."))
    print(f"Detalhe: {detalhe}")
    print(f"Por UF: {por_uf}")
    print(f"Por modelo: {por_modelo}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
