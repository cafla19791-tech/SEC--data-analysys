# World Bank Featured Indicators

Excel workbook with one sheet per featured World Bank indicator (countries/aggregates × years).

## File

- `world_bank_featured_indicators.xlsx`
  - **Index**: category, full indicator name, and code for every series
  - **Summary**: download status, coverage (rows and year range)
  - **One sheet per indicator**: named by indicator code (e.g. `NY.GDP.MKTP.CD`)

Each indicator sheet contains:

| Column | Description |
| --- | --- |
| Country Name | Country or aggregate name |
| Country Code | ISO3 / WDI code |
| Indicator Name | Full indicator label |
| Indicator Code | WDI series code |
| 1960…2025 | Annual values (when available) |

Source: [World Bank World Development Indicators](https://data.worldbank.org/) via the indicator CSV download API.

## Regenerate

```bash
python3 fetch_world_bank_indicators.py
```
