# Canadian multifamily research pipeline

A reproducible Python workflow for Canadian rental-market, population, construction and interest-rate research. It produces source-traceable datasets, six charts, an Excel workbook, validation diagnostics and an evidence brief.

## Quick start

Python 3.11 or newer is required. From this repository's folder:

```sh
python -m pip install -r requirements.txt
python RE_Pipeline.py
```

The default output folder is `output/` beside the script. No API keys, personal files, machine-specific directories or paid data subscriptions are required. Source downloads require internet access on the first run. On systems without an IANA time-zone database, install `tzdata` as well.

```sh
# Reuse validated inputs after a successful online run
python RE_Pipeline.py --offline

# Redownload sources
python RE_Pipeline.py --refresh

# Set the report year, observation cutoff and destination
python RE_Pipeline.py --rental-year 2025 --as-of 2026-10-04 --output-dir output

# Run focused tests without downloading data
python -m unittest discover -s . -p 'test_*.py' -v
```

Offline use requires cached inputs matching the requested dates. To reproduce a past snapshot, retain its raw files and metadata and pass its original `--as-of` date. An online run with an old cutoff uses currently available revisions; it is not a historical release-vintage backtest. Archived rental report layouts are supported for 2019–2025; future layouts may require updates. The default rental year is the prior calendar year.

## Sources and methodology

- CMHC Rental Market Survey, archived Canada workbooks: rental vacancy, fixed-sample rent growth, turnover rents and rental universe.
- Statistics Canada 17-10-0148-01: annual CMA population; 17-10-0009-01: national quarterly population.
- Statistics Canada 34-10-0154-01: CMHC monthly housing construction.
- Bank of Canada Valet: policy target and five- and ten-year government yields.

CMA population is matched to CMA rental observations. Construction flows require twelve months; December gives the year-end stock. YTD compares identical months across years. National rental centres of 10,000+ and the construction CMA aggregate remain separate universes. CMHC quality flags are retained, and suppressed or statistically insignificant growth is not imputed as zero.

## Outputs

- `raw/`: original downloads and retrieval/hash metadata.
- `clean/`: analysis-ready CSVs with geographic and temporal definitions.
- `charts/`: six PNG/PDF figures.
- `tables/`: formatted Excel workbook and CSV tables.
- `summary/`: evidence brief and analytical limitations.
- `diagnostics/`: 21 data checks, coverage, logs and provenance manifests.

Each successful run validates the data and exported workbook before replacing published output folders. Previous results are archived. No regression or composite investment ranking is presented: this is descriptive research. Valuation scenarios use explicit NOI and cap-rate assumptions and measure capital-value change, not total return.

## Research limits

October rental surveys, July population and calendar-year construction differ in timing. Census boundaries and historical revisions affect comparisons. Construction includes all tenures and dwelling types. Turnover rents compare groups of units, not matched-unit rent changes. The analysis does not establish causation or identify a particular fund's portfolio exposure.

## Sharing and reproducibility

This repository contains code and documentation only. `.gitignore` excludes default generated outputs, caches, local environments and logs. If you choose a custom output directory, add it to `.gitignore` before committing. Raw data redistribution is subject to the original providers' terms. No ownership of those datasets is implied.

The workflow was developed with AI assistance. Public source data, explicit calculations, automated checks and visual review support the research; interpretation still requires analyst judgment. Version 2.0.1 was executed against the 4 October 2026 source snapshot, with all 21 built-in data checks passing. Clean datasets matched the validated research-note snapshot exactly.
