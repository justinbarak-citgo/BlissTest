"""
BLISS lab ad-hoc query -> pandas DataFrame.

Library use
-----------
    from bliss_lab import BlissClient, pull

    # One-shot: login, pull, close
    df = pull()                                   # last 7 days through today
    df = pull("2026-10-01", "2026-10-15")         # explicit range (end is inclusive)

    # A different / ad-hoc server
    df = pull(endpoint="http://otherlab01")

    # Several pulls on one login
    with BlissClient() as bliss:
        last_week = bliss.query()
        september = bliss.query(start="2026-09-01", end="2026-09-30")
        month     = bliss.query(days_back=30, attributes=[11, 2, 6])
        other     = bliss.query(unit=1200, code=5)

CLI
---
    uv run bliss_lab.py
    uv run bliss_lab.py --start 2026-10-01 --end 2026-10-15 --debug
    uv run bliss_lab.py --endpoint http://otherlab01

Flow (reverse-engineered from the browser capture)
--------------------------------------------------
  0. GET Login.aspx, RC4-encrypt password with page Token64, POST the form
  1. GET excel.aspx?Action=excelh&...&type=horizontal -> 302 to excel.aspx?wait=true
  2. Poll GET excel.aspx?getstatus=1 (XML) until the job is done
  3. GET excel.aspx?getresults=1 -> the file, or a small page pointing to it

The job is tied to the ASP.NET_SessionId cookie, so one client = one session.
Queries on a client run sequentially; don't share one client across threads.

Endpoint: endpoint= accepts a name from ENDPOINTS, a URL, or a BlissEndpoint.
Default is $BLISS_HOST if set, otherwise ENDPOINTS[DEFAULT_ENDPOINT] (ccrlab05).

Credentials: pass user=/password=, or set BLISS_USER / BLISS_PASSWORD in the
environment or a .env file (next to this module, or in the working directory).
"""

from __future__ import annotations

import argparse
import base64
import io
import json
import logging
import os
import re
import time
import xml.etree.ElementTree as ET
from collections.abc import Iterable, Mapping
from dataclasses import dataclass
from datetime import date, datetime, timedelta
from html.parser import HTMLParser
from pathlib import Path
from types import TracebackType
from typing import Any, Final, TypeAlias
from urllib.parse import urlencode, urljoin, urlsplit

import pandas as pd
import requests
from dotenv import find_dotenv, load_dotenv

__all__ = [
    "DEFAULT_ENDPOINT",
    "ENDPOINTS",
    "BlissClient",
    "BlissEndpoint",
    "BlissError",
    "get_endpoint",
    "pull",
    "resolve_dates",
]

log: Final = logging.getLogger(__name__)


# ----------------------------------------------------------------------------- endpoints
@dataclass(frozen=True, slots=True)
class BlissEndpoint:
    """Where a BLISS instance lives. Paths are the stock BLISS layout."""

    host: str
    name: str = ""
    base_path: str = "/BLISS/Query/Adhoc.NET/"
    login_path: str = (
        "/BLISS/Pages/Security/Login.aspx?ReturnUrl=%2FBLISS%2FDefault.aspx"
    )

    def __post_init__(self) -> None:
        host = self.host.strip().rstrip("/")
        if not host.lower().startswith(("http://", "https://")):
            host = "http://" + host
        object.__setattr__(self, "host", host)
        if not self.name:
            object.__setattr__(self, "name", urlsplit(host).hostname or host)

    @property
    def base_url(self) -> str:
        return self.host + self.base_path

    @property
    def login_url(self) -> str:
        return self.host + self.login_path


# Known BLISS instances. Add more here when needed; key = short name.
ENDPOINTS: dict[str, BlissEndpoint] = {
    "ccrlab05": BlissEndpoint("http://ccrlab05", name="ccrlab05"),
}
DEFAULT_ENDPOINT: str = "ccrlab05"

EndpointLike: TypeAlias = BlissEndpoint | str | None


def get_endpoint(endpoint: EndpointLike = None) -> BlissEndpoint:
    """Resolve an endpoint.

    None            -> $BLISS_HOST if set, else ENDPOINTS[DEFAULT_ENDPOINT]
    BlissEndpoint   -> used as-is
    "ccrlab05"      -> looked up in ENDPOINTS
    "http://host"   -> ad-hoc endpoint with the standard BLISS paths
    """
    if isinstance(endpoint, BlissEndpoint):
        return endpoint
    if endpoint is None:
        _load_env()
        endpoint = os.environ.get("BLISS_HOST") or DEFAULT_ENDPOINT
    key = endpoint.strip()
    if key.lower() in ENDPOINTS:
        return ENDPOINTS[key.lower()]
    if "://" in key or "." in key or "/" in key:
        return BlissEndpoint(key)
    raise ValueError(
        f"Unknown BLISS endpoint {endpoint!r}. Known: {sorted(ENDPOINTS)}; "
        "or pass a full URL like 'http://server'."
    )


DEFAULT_DAYS_BACK: Final = 7
POLL_SECONDS: Final = 2.0
TIMEOUT_SECONDS: Final = 300.0
REQUEST_TIMEOUT: Final = 60.0
RESULTS_TIMEOUT: Final = 300.0

UA: Final = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/155.0.0.0 Safari/537.36 Edg/155.0.0.0"
)

# Query definition captured from the browser. Dates are filled per query;
# anything else can be overridden via query(..., extra_params={...}).
BASE_PARAMS: Final[dict[str, str]] = {
    "Action": "excelh",
    "StartDate": "",
    "EndDate": "",
    "Unit": "1000",
    "Code": "0",
    "TM": "0",
    "Det": "0",
    "rs": "",
    "re": "",
    "Attributes": "11.2.6.7.3.10.9.49",
    "KeyAttributes": "205.206.200",
    "KeySearch": "",
    "KST": "0",
    "Lab": "0",
    "Source": "0",
    "EntryType": "0",
    "ArrivedOnly": "0",
    "SpecType": "0",
    "TestsbyLab": "false",
    "FilterTests": "true",
    "IncludeBaseTestMethods": "0",
    "FQT": "false",
    "SQT": "00:00",
    "EQT": "24:00",
    "SearchType": "1",
    "Customer": "0",
    "Product": "0",
    "SampleType": "0,0",
    "IncludeCanceled": "false",
    "Page": "1",
    "new": "0",
    "myTabId": "tab_ext-gen126",  # UI artifact; try removing once it works
    "type": "horizontal",
}

XLSX_MAGIC: Final = b"PK\x03\x04"  # .xlsx (zip)
XLS_MAGIC: Final = b"\xd0\xcf\x11\xe0"  # legacy .xls (OLE)

DateLike: TypeAlias = date | datetime | str | pd.Timestamp
AttrList: TypeAlias = str | int | Iterable[int | str]


class BlissError(RuntimeError):
    pass


class _SessionExpired(BlissError):
    pass


# ----------------------------------------------------------------------------- helpers
def _load_env() -> None:
    """Load .env next to this module, then one found from the working dir.
    Existing environment variables are never overwritten."""
    load_dotenv(Path(__file__).with_name(".env"))
    if found := find_dotenv(usecwd=True):
        load_dotenv(found)


def _to_date(d: DateLike) -> date:
    """Accept date, datetime, Timestamp, or any string pandas can parse."""
    if isinstance(d, datetime):  # also covers pd.Timestamp
        return d.date()
    if isinstance(d, date):
        return d
    try:
        return pd.Timestamp(d).date()
    except (ValueError, TypeError) as e:
        raise ValueError(f"Can't parse date: {d!r}") from e


def resolve_dates(
    start: DateLike | None,
    end: DateLike | None,
    days_back: int = DEFAULT_DAYS_BACK,
) -> tuple[date, date]:
    """Defaults: end = today, start = end - days_back. End is inclusive."""
    end_d = _to_date(end) if end is not None else date.today()
    start_d = (
        _to_date(start) if start is not None else end_d - timedelta(days=days_back)
    )
    if start_d > end_d:
        raise ValueError(f"start ({start_d}) is after end ({end_d})")
    return start_d, end_d


def _fmt(d: date) -> str:
    return f"{d:%Y-%m-%d}T00:00:00.000Z"


def _dotted(v: AttrList) -> str:
    if isinstance(v, (str, int)):
        return str(v)
    return ".".join(str(x) for x in v)


def build_params(
    start: date,
    end: date,
    *,
    unit: str | int | None = None,
    code: str | int | None = None,
    attributes: AttrList | None = None,
    key_attributes: AttrList | None = None,
    extra_params: Mapping[str, Any] | None = None,
) -> dict[str, str]:
    p = dict(BASE_PARAMS)
    p["StartDate"] = _fmt(start)
    p["EndDate"] = _fmt(end + timedelta(days=1))  # BLISS end is exclusive at midnight
    if unit is not None:
        p["Unit"] = str(unit)
    if code is not None:
        p["Code"] = str(code)
    if attributes is not None:
        p["Attributes"] = _dotted(attributes)
    if key_attributes is not None:
        p["KeyAttributes"] = _dotted(key_attributes)
    if extra_params:
        unknown = set(extra_params) - set(BASE_PARAMS)
        if unknown:
            log.warning("Unknown BLISS params (sent anyway): %s", sorted(unknown))
        p.update({k: str(v) for k, v in extra_params.items()})
    return p


def _looks_like_login(resp: requests.Response) -> bool:
    url = resp.url.lower()
    body = (
        resp.text[:5000].lower()
        if "text" in resp.headers.get("Content-Type", "")
        else ""
    )
    return (
        "login" in url
        or "signin" in url
        or "account/" in url
        or 'type="password"' in body
        or resp.status_code in (401, 403)
    )


def _rc4(key: bytes, data: bytes) -> bytes:
    """BLISS login 'encryption' (Baytek.Utils.encrypt) is plain RC4."""
    S, j = list(range(256)), 0
    for i in range(256):
        j = (j + S[i] + key[i % len(key)]) % 256
        S[i], S[j] = S[j], S[i]
    i = j = 0
    out = bytearray()
    for c in data:
        i = (i + 1) % 256
        j = (j + S[i]) % 256
        S[i], S[j] = S[j], S[i]
        out.append(c ^ S[(S[i] + S[j]) % 256])
    return bytes(out)


class _Inputs(HTMLParser):
    """Collect every <input name=... value=...> on the login page."""

    def __init__(self) -> None:
        super().__init__()
        self.fields: dict[str, str] = {}

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag == "input":
            a = dict(attrs)
            if a.get("name"):
                self.fields[a["name"]] = a.get("value") or ""


def _is_done(xml_text: str) -> bool:
    """Parse BLISS <Status> XML. True = finished, False = still running."""
    try:
        root = ET.fromstring(xml_text.strip())
    except ET.ParseError as e:
        raise BlissError(f"Unreadable status reply:\n{xml_text[:500]}") from e

    def flag(tag: str) -> bool:
        return (root.findtext(tag) or "").strip().lower() == "true"

    if flag("isError"):
        raise BlissError(f"BLISS export failed:\n{xml_text[:500]}")
    return flag("isComplete")


def _to_frame(data: bytes, ext: str, sheet: int | str) -> pd.DataFrame:
    buf = io.BytesIO(data)
    if ext == ".html.xls":  # HTML table dressed up as Excel
        tables = pd.read_html(buf)
        if not tables:
            raise BlissError("Results page had no tables.")
        return tables[sheet if isinstance(sheet, int) else 0]
    return pd.read_excel(buf, sheet_name=sheet)


# ----------------------------------------------------------------------------- client
class BlissClient:
    """One logged-in BLISS session. Use as a context manager, or open()/close()."""

    def __init__(
        self,
        user: str | None = None,
        password: str | None = None,
        *,
        endpoint: EndpointLike = None,
        debug: bool = False,
        debug_dir: str | Path | None = None,
        poll_seconds: float = POLL_SECONDS,
        timeout_seconds: float = TIMEOUT_SECONDS,
    ) -> None:
        _load_env()
        self.user: str = user or os.environ.get("BLISS_USER", "")
        self._password: str = password or os.environ.get("BLISS_PASSWORD", "")
        self.endpoint: BlissEndpoint = get_endpoint(endpoint)
        self.host: str = self.endpoint.host
        self.base: str = self.endpoint.base_url
        self.login_url: str = self.endpoint.login_url
        self.debug: bool = debug
        self.debug_dir: Path = Path(debug_dir) if debug_dir else Path.cwd() / "_debug"
        self.poll_seconds: float = poll_seconds
        self.timeout_seconds: float = timeout_seconds
        self._s: requests.Session | None = None
        self._n: int = 0  # query counter, keeps debug dumps from overwriting

    # -- lifecycle -------------------------------------------------------------
    def __enter__(self) -> BlissClient:
        return self.open()

    def __exit__(
        self,
        exc_type: type[BaseException] | None,
        exc: BaseException | None,
        tb: TracebackType | None,
    ) -> None:
        self.close()

    def __repr__(self) -> str:
        state = "open" if self.is_open else "closed"
        return f"<BlissClient {self.user or '?'}@{self.endpoint.name} ({state})>"

    @property
    def is_open(self) -> bool:
        return self._s is not None

    def open(self) -> BlissClient:
        if self._s is None:
            self._s = self._login()
        return self

    def close(self) -> None:
        if self._s is not None:
            self._s.close()
            self._s = None
            log.info("BLISS session closed")

    # -- public ----------------------------------------------------------------
    def query(
        self,
        start: DateLike | None = None,
        end: DateLike | None = None,
        *,
        days_back: int = DEFAULT_DAYS_BACK,
        unit: str | int | None = None,
        code: str | int | None = None,
        attributes: AttrList | None = None,
        key_attributes: AttrList | None = None,
        extra_params: Mapping[str, Any] | None = None,
        sheet: int | str = 0,
        save_to: str | Path | None = None,
    ) -> pd.DataFrame:
        """Run one BLISS ad-hoc export and return it as a DataFrame.

        start / end   date, datetime, or string ("2026-10-01"). End is inclusive.
                      Defaults: end = today, start = end - days_back.
        days_back     only used when start is None (default 7).
        unit          BLISS Unit id (default "1000").
        code          BLISS Code id (default "0").
        attributes    list of ids [11, 2, 6] or dotted string "11.2.6".
        key_attributes  same format as attributes.
        extra_params  any other raw BLISS query params, e.g. {"Lab": "3"}.
        sheet         sheet index or name to return (default first sheet).
        save_to       folder to also save the raw file into (optional).
        """
        self.open()  # lazy login if needed
        start_d, end_d = resolve_dates(start, end, days_back)
        params = build_params(
            start_d,
            end_d,
            unit=unit,
            code=code,
            attributes=attributes,
            key_attributes=key_attributes,
            extra_params=extra_params,
        )
        qs = urlencode(params, safe=":,")  # keep ':' and ',' literal like the browser

        self._n += 1
        tag = f"{self.endpoint.name}_q{self._n:02d}"
        log.info("BLISS %s: %s -> %s", tag, start_d, end_d)

        try:
            self._submit(qs, tag)
        except _SessionExpired:
            log.info("Session expired - logging in again")
            self.close()
            self.open()
            self._submit(qs, tag)

        self._wait(tag)
        data, ext = self._fetch(tag)

        if save_to is not None:
            out = Path(save_to)
            out.mkdir(parents=True, exist_ok=True)
            path = (
                out
                / f"BLISS_{self.endpoint.name}_{start_d:%Y%m%d}_{end_d:%Y%m%d}_{datetime.now():%H%M%S}{ext}"
            )
            path.write_bytes(data)
            log.info("Saved %s bytes -> %s", f"{len(data):,}", path)

        return _to_frame(data, ext, sheet)

    # -- internals -------------------------------------------------------------
    @property
    def _session(self) -> requests.Session:
        if self._s is None:
            raise BlissError("Session not open.")
        return self._s

    def _dump(self, name: str, resp: requests.Response) -> None:
        if not self.debug:
            return
        self.debug_dir.mkdir(parents=True, exist_ok=True)
        meta = {
            "url": resp.url,
            "status": resp.status_code,
            "history": [
                (r.status_code, r.headers.get("Location")) for r in resp.history
            ],
            "headers": dict(resp.headers),
        }
        (self.debug_dir / f"{name}.meta.json").write_text(json.dumps(meta, indent=2))
        (self.debug_dir / f"{name}.body").write_bytes(resp.content)

    def _login(self) -> requests.Session:
        if not self.user or not self._password:
            raise BlissError(
                "BLISS_USER / BLISS_PASSWORD missing - pass user=/password= or set them in .env."
            )
        s = requests.Session()
        s.headers.update({"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"})

        # 1) Login page: ViewState, csrfToken, and the one-time Token64
        r = s.get(self.login_url, timeout=REQUEST_TIMEOUT)
        self._dump(f"{self.endpoint.name}_login_0_page", r)
        r.raise_for_status()
        m = re.search(r"var\s+Token64\s*=\s*'([^']+)'", r.text)
        if not m:
            s.close()
            raise BlissError("Token64 not found on login page - page layout changed?")
        token64 = m.group(1)
        p = _Inputs()
        p.feed(r.text)
        form = p.fields

        # 2) Fill the form the same way the page's onLogin() does
        enc = _rc4(base64.b64decode(token64), self._password.encode("latin-1"))
        form.update(
            {
                "clientDate": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "txtUsername": self.user,
                "txtPassword": "",
                "btnLogin": "Login",
                "hdnEncryptedPassword": base64.b64encode(enc).decode(),
                "hdnToken64": token64,
                "hdnNewPassword": "",
                "hdnConfirmNew": "",
            }
        )
        form.pop("btnSubmitNewPass", None)  # don't trigger the change-password button

        # 3) Submit
        r = s.post(
            self.login_url,
            data=form,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True,
            headers={"Origin": self.host, "Referer": self.login_url},
        )
        self._dump(f"{self.endpoint.name}_login_1_post", r)
        if "login.aspx" in r.url.lower():
            s.close()
            msg = re.search(r'id="lblOutput"[^>]*>([^<]*)', r.text)
            reason = (
                msg.group(1).strip()
                if msg and msg.group(1).strip()
                else "unknown reason"
            )
            raise BlissError(f"Login rejected: {reason}")

        log.info("Logged in to BLISS %s as %s", self.endpoint.name, self.user)
        return s

    def _submit(self, qs: str, tag: str) -> None:
        url = urljoin(self.base, "excel.aspx") + "?" + qs
        r = self._session.get(
            url,
            headers={"Referer": urljoin(self.base, "redirect.aspx") + "?" + qs},
            allow_redirects=True,
            timeout=REQUEST_TIMEOUT,
        )
        self._dump(f"{tag}_1_submit", r)
        if _looks_like_login(r):
            raise _SessionExpired("Got a login page - session expired or auth failed.")
        if "wait=true" not in r.url:
            raise BlissError(
                f"Expected redirect to excel.aspx?wait=true, landed on {r.url}"
            )

    def _wait(self, tag: str) -> None:
        url = urljoin(self.base, "excel.aspx?getstatus=1")
        hdrs = {
            "Accept": "application/xml, text/xml, */*; q=0.01",
            "X-Requested-With": "XMLHttpRequest",
            "Referer": urljoin(self.base, "excel.aspx?wait=true"),
        }
        t0, n, last = time.time(), 0, None
        while time.time() - t0 < self.timeout_seconds:
            r = self._session.get(url, headers=hdrs, timeout=REQUEST_TIMEOUT)
            n += 1
            self._dump(f"{tag}_2_status_{n:03d}", r)
            r.raise_for_status()
            if r.text != last:
                log.debug("[%s status %d] %s", tag, n, r.text.strip()[:200])
                last = r.text
            if _is_done(r.text):
                return
            time.sleep(self.poll_seconds)
        raise BlissError(f"Job not finished after {self.timeout_seconds}s")

    def _fetch(
        self, tag: str, url: str | None = None, depth: int = 0
    ) -> tuple[bytes, str]:
        url = url or urljoin(self.base, "excel.aspx?getresults=1")
        r = self._session.get(
            url,
            headers={"Referer": urljoin(self.base, "excel.aspx?wait=true")},
            allow_redirects=True,
            timeout=RESULTS_TIMEOUT,
        )
        self._dump(f"{tag}_3_results_{depth}", r)
        r.raise_for_status()
        data, ctype = r.content, r.headers.get("Content-Type", "").lower()
        cdisp = r.headers.get("Content-Disposition", "")

        if data[:4] == XLSX_MAGIC:
            return data, ".xlsx"
        if data[:4] == XLS_MAGIC:
            return data, ".xls"
        if "attachment" in cdisp.lower() or "excel" in ctype or "spreadsheet" in ctype:
            names = re.findall(r'filename="?([^";]+)', cdisp)
            ext = Path(names[0]).suffix if names else ".xls"
            return data, ext or ".xls"
        if b"<table" in data[:200000].lower():
            return data, ".html.xls"

        # Small HTML page pointing at the real file: follow it.
        if depth >= 3:
            raise BlissError(
                "Too many hops looking for the file - rerun with debug=True."
            )
        m = (
            re.search(rb"location(?:\.href)?\s*=\s*[\"']([^\"']+)", data)
            or re.search(rb"window\.open\(\s*[\"']([^\"']+)", data)
            or re.search(
                rb"(?:src|href)\s*=\s*[\"']([^\"']*(?:\.xlsx?|download|file|getresults)[^\"']*)",
                data,
                re.IGNORECASE,
            )
        )
        if not m:
            raise BlissError(f"Couldn't find file in results page:\n{data[:600]!r}")
        nxt = urljoin(r.url, m.group(1).decode())
        log.info("Following results link -> %s", nxt)
        return self._fetch(tag, nxt, depth + 1)


# ----------------------------------------------------------------------------- one-shot
def pull(
    start: DateLike | None = None,
    end: DateLike | None = None,
    *,
    endpoint: EndpointLike = None,
    debug: bool = False,
    **query_kwargs: Any,
) -> pd.DataFrame:
    """Login, run one query, close. Same arguments as BlissClient.query()."""
    with BlissClient(endpoint=endpoint, debug=debug) as bliss:
        return bliss.query(start, end, **query_kwargs)


# ----------------------------------------------------------------------------- CLI
def _main() -> None:
    ap = argparse.ArgumentParser(description="Pull BLISS lab data.")
    ap.add_argument(
        "--endpoint",
        default=None,
        help=f"name {sorted(ENDPOINTS)} or URL (default: $BLISS_HOST or {DEFAULT_ENDPOINT})",
    )
    ap.add_argument("--start", help="YYYY-MM-DD (default: end - days)")
    ap.add_argument("--end", help="YYYY-MM-DD, inclusive (default: today)")
    ap.add_argument("--days", type=int, default=DEFAULT_DAYS_BACK)
    ap.add_argument("--unit", help=f"BLISS Unit id (default {BASE_PARAMS['Unit']})")
    ap.add_argument("--code", help=f"BLISS Code id (default {BASE_PARAMS['Code']})")
    ap.add_argument(
        "--save",
        default=str(Path(__file__).resolve().parent / "output"),
        help="folder for the raw file",
    )
    ap.add_argument("--no-save", action="store_true")
    ap.add_argument(
        "--debug", action="store_true", help="dump every response to _debug/"
    )
    a = ap.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if a.debug else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )
    df = pull(
        a.start,
        a.end,
        days_back=a.days,
        unit=a.unit,
        code=a.code,
        endpoint=a.endpoint,
        debug=a.debug,
        save_to=None if a.no_save else a.save,
    )
    print(df.shape)
    print(df.head())


if __name__ == "__main__":
    _main()
