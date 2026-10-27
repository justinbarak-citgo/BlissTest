# bliss

Pull BLISS lab ad-hoc query results straight into a pandas DataFrame. No browser, no manual Excel export.

Log in once, run as many queries as you need, close. Default window is the **last 7 days through today**.

## Quick start

```python
from bliss import BlissClient, pull

df = pull()                                   # last 7 days, default endpoint
df = pull("2026-10-01", "2026-10-15")         # explicit range (end is inclusive)

with BlissClient() as bliss:                  # one login, many pulls
    wk  = bliss.query()
    sep = bliss.query(start="2026-09-01", end="2026-09-30")
    mo  = bliss.query(days_back=30, attributes=[11, 2, 6])
    alt = bliss.query(unit=1200, code=5)
```

Command line (returns data only; add `--save` to also keep the raw file):

```powershell
uv run bliss-pull
uv run bliss-pull --start 2026-10-01 --end 2026-10-15
uv run bliss-pull --days 14 --endpoint http://otherlab01 --debug
uv run bliss-pull --unit 1200 --code 5
uv run bliss-pull --save                 # also save to ./output
uv run bliss-pull --save C:\data\bliss    # also save to a specific folder
```

### CLI options

| Option | Default | Notes |
|---|---|---|
| `--start` | `end - days` | `YYYY-MM-DD` |
| `--end` | today | `YYYY-MM-DD`, inclusive |
| `--days` | `7` | Only used when `--start` is omitted |
| `--unit` / `--code` | `1000` / `0` | BLISS Unit / Code ids |
| `--endpoint` | `$BLISS_HOST` or `ccrlab05` | Name or URL |
| `--save [FOLDER]` | off | Also save the raw file; `--save` alone uses `./output` |
| `--debug` | off | Dump every HTTP response to `./_debug/` |

## Setup

**Requirements:** Python 3.13+, uv, network access to the BLISS server.

Work on the library itself:

```powershell
uv sync
uv run bliss-pull --debug
```

Use it from another uv project (editable, so library changes show up immediately):

```powershell
cd C:\dev\my-other-project
uv add --editable C:\dev\bliss
```

Or from a Git repo, pinned to a release tag:

```powershell
uv add "bliss @ git+https://github.com/<org>/bliss.git@v0.2.2"
```

Optional: if BLISS ever returns legacy `.xls` or an HTML table instead of `.xlsx`:

```powershell
uv add "bliss[legacy]"      # adds xlrd + lxml
```

## Configuration

Create a `.env` in the root of the project that **uses** the library. **Never commit it**: add `.env` to `.gitignore`.

```dotenv
BLISS_USER=citgo_network_id
BLISS_PASSWORD="your-password"
# BLISS_HOST=ccrlab05        # optional: endpoint name or URL, overrides the default
```

Lookup order (first value found wins):

1. Real environment variables
2. `.env` in the calling project (searched upward from the working directory)
3. `.env` next to `client.py` (fallback)

You can also pass `user=` / `password=` directly to `BlissClient`.

### Endpoints

| You pass | Result |
|---|---|
| `None` | `$BLISS_HOST`, else `ccrlab05` |
| `"ccrlab05"` | Looked up in `ENDPOINTS` (case-insensitive) |
| `"http://otherlab01"` / `"otherlab01.citgo.com"` | Ad-hoc endpoint, standard BLISS paths |
| `BlissEndpoint(...)` | Used as-is |

Add a permanent endpoint in `src/bliss/client.py`:

```python
ENDPOINTS: dict[str, BlissEndpoint] = {
    "ccrlab05": BlissEndpoint("http://ccrlab05", name="ccrlab05"),
    "lclab01":  BlissEndpoint("http://lclab01", name="lclab01"),
}
```

Non-standard install paths:

```python
from bliss import BlissEndpoint, pull

ep = BlissEndpoint("http://server", name="custom",
                   base_path="/LIMS/Query/Adhoc.NET/",
                   login_path="/LIMS/Pages/Security/Login.aspx")
df = pull(endpoint=ep)
```

One `.env` login is used for every endpoint. If servers need different accounts, pass `user=` / `password=` per client.

## API

### `BlissClient(user=None, password=None, *, endpoint=None, debug=False, debug_dir=None, poll_seconds=2.0, timeout_seconds=300.0)`

One logged-in session against one endpoint. Use as a context manager, or call `.open()` / `.close()`. `query()` logs in lazily if needed. After each query, `client.last_status` holds BLISS's final status (sample count, pages, etc.).

### `BlissClient.query(...) -> pd.DataFrame`

| Param | Default | Notes |
|---|---|---|
| `start` | `end - days_back` | date, datetime, Timestamp, or string (`"2026-10-01"`, `"10/1/2026"`) |
| `end` | today (local) | **Inclusive** |
| `days_back` | `7` | Only used when `start` is `None` |
| `unit` | `"1000"` | BLISS Unit id |
| `code` | `"0"` | BLISS Code id |
| `attributes` | `"11.2.6.7.3.10.9.49"` | List `[11, 2, 6]` or dotted string |
| `key_attributes` | `"205.206.200"` | Same format |
| `extra_params` | `None` | Raw overrides for any BLISS query param, e.g. `{"Lab": "3"}` |
| `sheet` | `0` | Sheet index or name to return |
| `save_to` | `None` | No file is saved by default. Pass a folder to also save the raw file |

### `pull(start=None, end=None, *, endpoint=None, debug=False, **query_kwargs) -> pd.DataFrame`

Login → one query → close. Takes the same keywords as `query()`.

### Helpers

- `get_endpoint(endpoint)` → resolved `BlissEndpoint`
- `resolve_dates(start, end, days_back)` → `(start_date, end_date)`, handy for logging what window you pulled

## Behavior notes

- **Sequential per client.** BLISS ties the export job to the session cookie. For parallel pulls, use one client per thread.
- **Auto re-login.** If the session expires mid-batch, the client logs in again once and retries.
- **1,500-sample page limit.** BLISS pages large results. If a query returns more than one page, a warning is logged and only page 1 is downloaded. Narrow the date range.
- **Logging, not print.** Uses the `bliss.client` logger. Enable with `logging.basicConfig(level=logging.INFO)`; the CLI does this for you.
- **Dates.** BLISS's `EndDate` is midnight-exclusive, so the client sends `end + 1 day`.

## Troubleshooting

Run with `debug=True` (or `--debug`). Every HTTP response is written to `./_debug/` as `.meta.json` + `.body` pairs, named by endpoint, query number and step (e.g. `ccrlab05_login_1_post`, `ccrlab05_q01_2_status_003`).

| Error | Likely cause |
|---|---|
| `BLISS_USER / BLISS_PASSWORD missing` | No `.env` found, or variables not set |
| `Login rejected: ...` | Wrong password, locked/expired account. The reason is BLISS's own message |
| `Token64 not found on login page` | Login page changed, or wrong endpoint/`login_path` |
| `Expected redirect to excel.aspx?wait=true` | Bad query params, or wrong `base_path` |
| `Job not finished after 300s` | Large date range. Raise `timeout_seconds` |
| `Unknown BLISS endpoint 'x'` | Not in `ENDPOINTS` and not a URL. Add it or pass `http://x` |
| `Missing optional dependency 'xlrd'` / `'lxml'` | BLISS returned `.xls` or HTML. Install `bliss[legacy]` |

## How it works

Reverse-engineered from the browser:

1. GET `Login.aspx` → RC4-encrypt the password with the page's one-time `Token64` → POST the form
2. GET `excel.aspx?Action=excelh&...` → 302 to `excel.aspx?wait=true`
3. Poll GET `excel.aspx?getstatus=1` (XML) until `isComplete`
4. GET `excel.aspx?getresults=1` → the file (`.xlsx`, `.xls`, or an HTML table), following up to 3 redirect pages
