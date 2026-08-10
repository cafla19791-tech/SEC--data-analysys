# World Bank Featured Indicators

Excel workbooks with one sheet per featured World Bank indicator (countries/aggregates × years).

## Files

### Macro / growth set
`world_bank_featured_indicators.xlsx`

Growth and economic structure, income and savings, balance of payments, prices and terms of trade, labor and productivity.

### Business / infrastructure set
`world_bank_featured_indicators_business_infra.xlsx`

Business environment, financial access and stability, stock markets, government finance and taxes, military and fragile situations, infrastructure and communications, science and innovation.

### Poverty / shared prosperity set
`world_bank_featured_indicators_poverty.xlsx`

Poverty rates at national and international poverty lines, income/consumption distribution, and shared prosperity.

## Workbook layout

Each file contains:

- **Index**: category, full indicator name, and code
- **One sheet per indicator**: named by indicator code (e.g. `IC.REG.DURS`)
- **Summary**: download status and coverage (rows and year range)

Each indicator sheet contains:

| Column | Description |
| --- | --- |
| Country Name | Country or aggregate name |
| Country Code | ISO3 / WDI code |
| Indicator Name | Full indicator label |
| Indicator Code | WDI series code |
| 1960… | Annual values (when available) |

Source: [World Bank World Development Indicators](https://data.worldbank.org/) via the indicator CSV download API.

## Regenerate

```bash
# Both workbooks
python3 fetch_world_bank_indicators.py

# Only one set
python3 fetch_world_bank_indicators.py --set growth-macro
python3 fetch_world_bank_indicators.py --set business-infra
python3 fetch_world_bank_indicators.py --set poverty
```
