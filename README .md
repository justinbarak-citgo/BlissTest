# bliss_lab

Pull BLISS lab ad-hoc query results straight into a pandas DataFrame — no browser, no manual Excel export.

Log in once, run as many queries as you need, close. Default window is the **last 7 days through today**.

---

## Quick start

```python
from bliss_lab import BlissClient, pull

df = pull()                                   # last 7 days, default endpoint
df = pull("2026-10-01", "2026-10-15")         # explicit range (end is inclusive)

with BlissClient() as bliss:                  # one login, many pulls
    wk  = bliss.query()
    sep = bliss.query(start="2026-09-01", end="2026-09-30")
    mo  = bliss.query(days_back=30, attributes=[11, 2, 6])
    alt = bliss.query(unit=1200, code=5)
```

Command line:

```
uv run bliss_lab.py
uv run bliss_lab.py --start 2026-10-01 --end 2026-10-15
uv run bliss_lab.py --days 14 --endpoint http://otherlab01 --debug
uv run bliss_lab.py --unit 1200 --code 5
```

---

## Setup (uv)

**Requirements:** Python 3.11+, network access to the BLISS server.

```
uv add pandas requests python-dotenv openpyxl lxml
uv add xlrd            # only if BLISS ever returns legacy .xls
uv add --dev ruff
```

**Use from another project** — package it once, then add it as an editable dependency:

```
uv init --lib bliss-lab              # move bliss_lab.py -> src/bliss_lab/__init__.py
cd ../my-other-project
uv add --editable ../bliss-lab
```

---

## Configuration

Create a `.env` (next to `bliss_lab.py` or in the folder you run from). **Never commit it** — add `.env` to `.gitignore`.

```
BLISS_USER=citgo_network_id
BLISS_PASSWORD="your-password"
# BLISS_HOST=ccrlab05        # optional: endpoint name or URL, overrides the default
```

Existing environment variables win over `.env`. You can also pass `user=` / `password=` directly to `BlissClient`.

### Endpoints

Endpoint resolution order:

1. `endpoint=` argument (name, URL, or `BlissEndpoint`)
2. `BLISS_HOST` environment variable
3. `ENDPOINTS[DEFAULT_ENDPOINT]` → `ccrlab05`

| You pass | Result |
|---|---|
| `None` | `$BLISS_HOST`, else `ccrlab05` |
| `"ccrlab05"` | Looked up in `ENDPOINTS` (case-insensitive) |
| `"http://otherlab01"` / `"otherlab01.citgo.com"` | Ad-hoc endpoint, standard BLISS paths |
| `BlissEndpoint(...)` | Used as-is |

**Add a permanent endpoint** in `bliss_lab.py`:

```python
ENDPOINTS: dict[str, BlissEndpoint] = {
    "ccrlab05": BlissEndpoint("http://ccrlab05", name="ccrlab05"),
    "lclab01":  BlissEndpoint("http://lclab01", name="lclab01"),
}
```

**Non-standard install paths:**

```python
ep = BlissEndpoint("http://server", name="custom",
                   base_path="/LIMS/Query/Adhoc.NET/",
                   login_path="/LIMS/Pages/Security/Login.aspx")
df = pull(endpoint=ep)
```

> One `.env` login is used for every endpoint. If servers need different accounts, pass `user=` / `password=` per client.

---

## API

### `BlissClient(user=None, password=None, *, endpoint=None, debug=False, debug_dir=None, poll_seconds=2.0, timeout_seconds=300.0)`

One logged-in session against one endpoint. Use as a context manager, or call `.open()` / `.close()`. `query()` logs in lazily if needed.

### `BlissClient.query(...) -> pd.DataFrame`

| Param | Default | Notes |
|---|---|---|
| `start` | `end - days_back` | `date`, `datetime`, `Timestamp`, or string (`"2026-10-01"`, `"10/1/2026"`) |
| `end` | today (local) | **Inclusive** |
| `days_back` | `7` | Only used when `start` is `None` |
| `unit` | `"1000"` | BLISS Unit id |
| `code` | `"0"` | BLISS Code id |
| `attributes` | `"11.2.6.7.3.10.9.49"` | List `[11, 2, 6]` or dotted string |
| `key_attributes` | `"205.206.200"` | Same format |
| `extra_params` | `None` | Raw overrides for any BLISS query param, e.g. `{"Lab": "3"}` |
| `sheet` | `0` | Sheet index or name to return |
| `save_to` | `None` | Folder to also save the raw file |

### `pull(start=None, end=None, *, endpoint=None, debug=False, **query_kwargs) -> pd.DataFrame`

Login → one query → close. Takes the same keywords as `query()`.

### Helpers

- `get_endpoint(endpoint)` → resolved `BlissEndpoint`
- `resolve_dates(start, end, days_back)` → `(start_date, end_date)` — handy for logging what window you pulled

---

## Behavior notes

- **Sequential per client.** BLISS ties the export job to the session cookie. For parallel pulls, use one client per thread.
- **Auto re-login.** If the session expires mid-batch, the client logs in again once and retries.
- **Logging, not print.** Uses the `bliss_lab` logger. Enable with `logging.basicConfig(level=logging.INFO)`; the CLI does this for you.
- **Dates.** BLISS's `EndDate` is midnight-exclusive, so the client sends `end + 1 day`.

---

## Troubleshooting

Run with `debug=True` (or `--debug`). Every HTTP response is written to `./_debug/` as `<endpoint>_<step>.meta.json` + `.body`.

| Error | Likely cause |
|---|---|
| `BLISS_USER / BLISS_PASSWORD missing` | No `.env` found, or variables not set |
| `Login rejected: ...` | Wrong password, locked/expired account — the reason is BLISS's own message |
| `Token64 not found on login page` | Login page changed, or wrong endpoint/`login_path` |
| `Expected redirect to excel.aspx?wait=true` | Bad query params, or wrong `base_path` |
| `Job not finished after 300s` | Large date range — raise `timeout_seconds` |
| `Unknown BLISS endpoint 'x'` | Not in `ENDPOINTS` and not a URL — add it or pass `http://x` |

---

## How it works

Reverse-engineered from the browser:

0. `GET Login.aspx` → RC4-encrypt the password with the page's one-time `Token64` → `POST` the form
1. `GET excel.aspx?Action=excelh&...` → 302 to `excel.aspx?wait=true`
2. Poll `GET excel.aspx?getstatus=1` (XML) until `isComplete`
3. `GET excel.aspx?getresults=1` → the file (`.xlsx`, `.xls`, or an HTML table), following up to 3 redirect pages
