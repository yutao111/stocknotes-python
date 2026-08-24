# AGENTS.md

Compact instructions for OpenCode sessions working on this repo.

## Overview

Single-file Flask app (`app.py`) with SQLite backend. Parses stock trading statements (券商交割单), calculates FIFO matches, and displays P&L analysis.

## Commands

```bash
# Install and run
python3 -m pip install -r requirements.txt
python3 app.py

# Run tests
python3 -m unittest -v

# Dev server
python3 app.py
# → http://127.0.0.1:5001
```

## Architecture

- **app.py** (652 lines): Flask routes, DB schema, parsing logic, FIFO calculation, SVG chart generation
- **templates/**: Jinja2 templates for UI
- **static/**: CSS and assets
- **stocknotes.db**: SQLite database (auto-created on first run)
- **test_app.py**: unittest suite with full import/FIFO coverage

## Key execution flow

1. User uploads `.xlsx`/`.xls`/`.csv`/`.txt` statement file via `/admin/import`
2. `read_upload()` → parses file into rows
3. `parse_rows()` → normalizes headers using `HEADER_ALIASES`, validates and cleans data, generates fingerprint for deduplication
4. Preview shown with validation errors
5. On confirm: insert into `executions`, then call `rebuild_fifo()`
6. `rebuild_fifo()` clears and recalculates all `fifo_matches`, `positions`, `unmatched_sells` from scratch

## Important patterns

- **Fingerprint deduplication**: `app.py:293-294` — SHA256 hash of `trade_date|trade_time|stock_code|action|quantity|deal_price|deal_id` prevents duplicate imports
- **FIFO calculation**: `app.py:305-351` — processes all executions chronologically, maintains per-stock buy queues, matches sells against oldest buys
- **Stock code normalization**: `app.py:184-194` — strips prefixes (SH/SZ), pads to 6 digits, handles float and string inputs
- **Chart rendering**: `app.py:413-503` — generates inline SVG with price line, buy/sell dots, profit labels with collision avoidance
- **Database**: SQLite with foreign keys enabled (`app.py:57`), schema created in `init_db()` (`app.py:61-144`)

## Testing quirks

- Tests use `STOCKNOTES_DB` env var to isolate temp database (`test_app.py:14`)
- Must import `app` module inside `setUpClass` after setting env var
- Test data uses simplified statements with integer dates like `20260801` that get normalized to `2026-08-01`
- Chart label placement tests verify collision avoidance logic works (`test_app.py:101-106`)

## Common gotchas

- All FIFO recalculation is full rebuild, not incremental — performance acceptable for typical personal trading volume
- Date/time normalization handles multiple formats: YYYYMMDD integers, `YYYY-MM-DD` strings, datetime objects, timestamps
- Quantity normalization: negative values indicate sells in some broker formats
- Fees always absolute value during parsing (`app.py:284-288`)
- DB path defaults to `./stocknotes.db` but respects `STOCKNOTES_DB` env var
- Templates expect specific dict keys from `analysis_data()` — match the contract when modifying

## Style conventions

- Single quotes avoided in Python (uses double quotes)
- Flask routes use `@app.get` / `@app.post` decorators
- Database connections use context managers
- Template filter names: `money`, `percent` (`app.py:506-513`)
- Chinese user-facing strings, English code/comments
