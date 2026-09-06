# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

Pre-processed election results for North Dakota, part of the OpenElections project. The repo is primarily *data*: CSV files organized by election year (`2012/`, `2016/`, `2018/`, `2020/`). The small amount of code here consists of one-off parser scripts that converted raw source files into the CSVs.

Two companion repositories matter when working here:

- `openelections-sources-nd` — the raw source files (state-provided Excel/PDF/etc.). Parsers in this repo read from a sibling checkout of it (e.g. `parse_excel.R` uses `../openelections-sources-nd`).
- `openelections/openelections-data-tests` — the validation test suite. It is *not* in this repo; CI checks it out at tag `v2.2.0`. To run tests locally, clone it and run against this repo:

```bash
git clone https://github.com/openelections/openelections-data-tests ../openelections-data-tests
python3 ../openelections-data-tests/run_tests.py duplicate_entries file_format missing_values vote_breakdown_totals .
# Run a single test:
python3 ../openelections-data-tests/run_tests.py file_format .
# Test only specific files (how the PR workflow runs it):
python3 ../openelections-data-tests/run_tests.py --files path/to/file.csv vote_breakdown_totals .
```

The four tests are `duplicate_entries`, `file_format`, `missing_values`, and `vote_breakdown_totals` — new/modified CSVs must pass all four. The suite requires Python 3 (see `.python-version`: 3.12, managed by uv via `pyproject.toml`).

## Data conventions

- **File naming**: `<YYYYMMDD>__nd__<election>__<level>.csv` inside the year directory, e.g. `2016/20161108__nd__general__precinct.csv`. `<election>` is typically `general` or `primary`; `<level>` is `county` or `precinct`.
- **CSV schema** (lowercase headers, in this order): `county, precinct, office, district, party, candidate, votes`. County-level files omit `precinct` (and possibly other columns); rows carry an empty `district` or `party` when not applicable.
- One row per county/precinct × office × candidate, with vote totals as integers.

## CI (GitHub Actions)

- `data_tests.yml` — on every push/PR, runs all four tests against the whole repo.
- `data_tests_changed_files.yml` — on PRs, runs the tests only against added/changed `*.csv` files.
- `pull_request_comment.yml` — on failure of the changed-files workflow, a bot posts the failure logs as a PR comment.

Failure logs are uploaded as workflow artifacts when tests fail.

## Parser scripts

- `parse_excel.R` — generated the 2016 general county/precinct CSVs from per-office Excel files in the sources repo (R + readxl/dplyr/tidyr; county sheets named per county).
- `parser.py` — one-off xlrd script that extracted a single legislative-district State House race (District 47, DEM) from `Legislative District Precinct Results.xlsx`.

These are historical, not a general pipeline: each new election year typically means writing a new parser (or hand-cleaning) against that year's source format, producing files that follow the conventions above.