"""Helpers de formatacao Excel (largura de colunas)."""

from __future__ import annotations

from typing import Any, Iterable, Sequence


def _display_width(value: Any) -> int:
    if value is None:
        return 0
    try:
        import math

        if isinstance(value, float) and (math.isnan(value) or math.isinf(value)):
            return 0
    except Exception:
        pass
    s = str(value)
    # Excel width ~ ASCII chars; accented letters still count as 1 visually enough.
    return len(s)


def column_widths_for_frame(
    columns: Sequence[Any],
    rows: Iterable[Sequence[Any]] | None = None,
    *,
    min_width: float = 10,
    max_width: float = 60,
    padding: float = 3,
) -> list[float]:
    """Calcula larguras para caber rotulo + conteudo (aprox. chars Excel)."""
    widths = [max(min_width, _display_width(c) + padding) for c in columns]
    if rows is not None:
        for row in rows:
            for i, cell in enumerate(row):
                if i >= len(widths):
                    break
                widths[i] = max(widths[i], min(max_width, _display_width(cell) + padding))
    return [min(max_width, w) for w in widths]


def apply_column_widths(
    worksheet: Any,
    widths: Sequence[float],
    *,
    engine: str,
    formats: Sequence[Any] | None = None,
) -> None:
    """Aplica larguras no worksheet (xlsxwriter ou openpyxl)."""
    if engine == "xlsxwriter":
        for i, width in enumerate(widths):
            fmt = formats[i] if formats and i < len(formats) else None
            if fmt is not None:
                worksheet.set_column(i, i, width, fmt)
            else:
                worksheet.set_column(i, i, width)
        return

    # openpyxl
    from openpyxl.utils import get_column_letter

    for i, width in enumerate(widths, start=1):
        worksheet.column_dimensions[get_column_letter(i)].width = width


def autosize_dataframe_sheet(
    writer: Any,
    sheet_name: str,
    df: Any,
    *,
    engine: str,
    col_formats: Sequence[Any] | None = None,
    sample_rows: int = 200,
    min_width: float = 10,
    max_width: float = 60,
    padding: float = 3,
    extra_title_width: float | None = None,
) -> None:
    """Ajusta colunas de uma aba ja escrita com pandas.to_excel."""
    ws = writer.sheets[sheet_name]
    cols = list(df.columns)
    # Amostra para nao varrer milhoes de linhas
    sample = df.head(sample_rows)
    rows = sample.itertuples(index=False, name=None)
    widths = column_widths_for_frame(
        cols,
        rows,
        min_width=min_width,
        max_width=max_width,
        padding=padding,
    )
    if extra_title_width is not None and widths:
        widths[0] = max(widths[0], extra_title_width)
    apply_column_widths(ws, widths, engine=engine, formats=col_formats)
