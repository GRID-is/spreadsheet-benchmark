#!/usr/bin/env python3
"""Drive round-trip for the google-sheets engine: convert fixtures to Google
Sheets and export the recalculated values back as xlsx.

Uploading an xlsx with target mimeType google-apps.spreadsheet makes Google
import and recalculate it; exporting that Sheet back to xlsx carries the
computed values along as cached values, which replay/google-sheets.ts scores
against the Excel oracle. That is the whole trick - no Sheets API involved,
only Drive.

Scope is drive.file, which reaches ONLY files this script itself creates. It
cannot read, modify, or even see anything else in the account's Drive, so the
uploaded corpus and its converted copies are the entire blast radius. The
script creates its own folder for the same reason: a folder made by some other
app is not ours to write into under this scope.

Drive layout: one folder, engine-compatibility-corpus, with a subfolder per
suite mirroring suites/<suite>/excel/ (so ironcalc/FINANCIAL/MIRR.xlsx lives at
engine-compatibility-corpus/ironcalc/FINANCIAL/MIRR). That folder is the single
home of the corpus on Drive: `run` uploads whatever is missing there and never
uploads a fixture Drive already holds, so re-exporting costs no new copies.

Zero dependencies (Python 3 stdlib only). The OAuth client, the tokens and
the Drive folder ID of each account live in CONFIG_DIR. The client is a Desktop-app
OAuth client of your own Google Cloud project, with the Drive API enabled.

Usage:
  python3 replay/gsheets_drive.py auth [account]      # one-time, opens browser
  python3 replay/gsheets_drive.py run <suite>...      # round-trip those suites
  python3 replay/gsheets_drive.py run --all           # every suite but real-workbook ones
  python3 replay/gsheets_drive.py run --all --since <YYYY-MM-DD>  # see below
  python3 replay/gsheets_drive.py status              # current exports / on Drive, per suite
  python3 replay/gsheets_drive.py ready               # exit 1 unless client and token are stored
  python3 replay/gsheets_drive.py migrate             # move flat "suite/rel" names into subfolders
  python3 replay/gsheets_drive.py cleanup             # trash everything created

`run` is resumable and idempotent: a fixture whose export is current is
skipped, so an interrupted run costs only the files it had not reached. A
fixture that fails does not stop the run, but makes it exit with PARTIAL_EXIT
at the end; the replay records a load error for each fixture without an export.
A run in which every fixture failed, so that none has a current export, exits
with 1 instead. An
export is current when its capture identity (<export>.json: the SHA-256 of the
fixture uploaded and of the export received, and the UTC date of the export)
matches the fixture and export on disk. A changed fixture, or an export without
an identity, is exported again; a changed fixture is uploaded again too (its
Sheet on Drive is replaced), an unchanged one is exported from the Sheet already
there.

With --since <date>, an export made before that UTC date is not current either,
and its fixture is uploaded again, so that Google converts it afresh: a dated
replay (replay/google-sheets.ts) scores only exports made on or after its own
date.
"""

import base64
import concurrent.futures
import hashlib
import html
import http.server
import io
import json
import mimetypes
import os
import posixpath
import re
import secrets
import socket
import stat
import sys
import threading
import time
import urllib.error
import urllib.parse
import urllib.request
import webbrowser
import xml.etree.ElementTree as ET
import zipfile

CONFIG_DIR = os.path.expanduser("~/.config/engine-compatibility-drive")
CLIENT_FILE = os.path.join(CONFIG_DIR, "client.json")
SCOPE = "https://www.googleapis.com/auth/drive.file"
AUTH_URL = "https://accounts.google.com/o/oauth2/v2/auth"
TOKEN_URL = "https://oauth2.googleapis.com/token"
UPLOAD_API = "https://www.googleapis.com/upload/drive/v3/files"
DRIVE_API = "https://www.googleapis.com/drive/v3/files"

SHEET_MIME = "application/vnd.google-apps.spreadsheet"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
FOLDER_MIME = "application/vnd.google-apps.folder"

REPO = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SUITES_DIR = os.path.join(REPO, "suites")
EXPORT_DIR = os.environ.get("GSHEETS_EXPORT_DIR") or os.path.join(
    REPO, ".scratch", "gsheets-exports")
FOLDER_NAME = "engine-compatibility-corpus"
WORKERS = int(os.environ.get("GSHEETS_WORKERS", "6"))
# GRID's bot account, whose Drive holds the corpus, so that no person's account is needed.
ACCOUNT = os.environ["GSHEETS_ACCOUNT"]

# The exit status of a run in which some fixtures failed.
PARTIAL_EXIT = 3

_token_lock = threading.Lock()
_print_lock = threading.Lock()


# ---------------------------------------------------------------- credentials

def token_path(account):
    return os.path.join(CONFIG_DIR, "token-%s.json" % account)


def state_path(account):
    return os.path.join(CONFIG_DIR, "state-%s.json" % account)


def load_client():
    """Read the Desktop-app OAuth client, which must be in CLIENT_FILE."""
    if not os.path.exists(CLIENT_FILE):
        sys.exit(
            "No OAuth client at %s.\nSave there the JSON of a Desktop-app OAuth client of your"
            " Google Cloud project\nwith the Drive API enabled ."
            % CLIENT_FILE)
    c = json.load(open(CLIENT_FILE)).get("installed")
    if not c:
        sys.exit("%s is not a Desktop-app OAuth client: it has no 'installed' key."
                 % CLIENT_FILE)
    return c["client_id"], c["client_secret"]


def save_private(path, obj):
    os.makedirs(CONFIG_DIR, exist_ok=True)
    os.chmod(CONFIG_DIR, stat.S_IRWXU)
    with open(path, "w") as f:
        json.dump(obj, f, indent=2)
    os.chmod(path, stat.S_IRUSR | stat.S_IWUSR)


def post_form(url, fields):
    body = urllib.parse.urlencode(fields).encode()
    req = urllib.request.Request(url, data=body, method="POST")
    req.add_header("Content-Type", "application/x-www-form-urlencoded")
    with urllib.request.urlopen(req, timeout=60) as res:
        return json.load(res)


def auth_command(account):
    return "python3 %s auth %s" % (os.path.relpath(os.path.abspath(__file__), REPO), account)


def sign_in_instructions(account):
    return "Someone who can sign in as %s runs, once: %s" % (account, auth_command(account))


def require_token(account):
    if not os.path.exists(token_path(account)):
        sys.exit("No Drive token for %s in %s.\n%s"
                 % (account, CONFIG_DIR, sign_in_instructions(account)))


def cmd_ready(account):
    """Exit 1 unless the client and a token for the account are stored. Calls no Google API, so
    a token that Google has revoked passes."""
    load_client()
    require_token(account)
    print("%s has a Drive token in %s" % (account, CONFIG_DIR))


def get_access_token(account):
    with _token_lock:
        path = token_path(account)
        require_token(account)
        tok = json.load(open(path))
        if tok.get("expires_at", 0) > time.time() + 120:
            return tok["access_token"]
        client_id, client_secret = load_client()
        try:
            fresh = post_form(TOKEN_URL, {
                "client_id": client_id,
                "client_secret": client_secret,
                "refresh_token": tok["refresh_token"],
                "grant_type": "refresh_token",
            })
        except urllib.error.HTTPError as e:
            try:
                error = json.loads(e.read())
            except ValueError:
                error = None
            if not isinstance(error, dict) or error.get("error") != "invalid_grant":
                raise
            # Google refuses a refresh token that was revoked, or expired unused.
            sys.exit("Google refused the Drive token for %s in %s: %s\n%s"
                     % (account, CONFIG_DIR, error.get("error_description", "invalid_grant"),
                        sign_in_instructions(account)))
        tok["access_token"] = fresh["access_token"]
        tok["expires_at"] = time.time() + int(fresh.get("expires_in", 3600))
        save_private(path, tok)
        return tok["access_token"]


def cmd_auth(account):
    """Loopback OAuth with PKCE. Prints a URL; you approve it in the browser."""
    client_id, client_secret = load_client()
    verifier = secrets.token_urlsafe(64)
    challenge = base64.urlsafe_b64encode(
        hashlib.sha256(verifier.encode()).digest()).decode().rstrip("=")
    state = secrets.token_urlsafe(16)
    box = {}

    class Handler(http.server.BaseHTTPRequestHandler):
        def do_GET(self):
            q = urllib.parse.parse_qs(urllib.parse.urlparse(self.path).query)
            box.update({k: v[0] for k, v in q.items()})
            self.send_response(200)
            self.send_header("Content-Type", "text/html; charset=utf-8")
            self.end_headers()
            ok = box.get("state") == state and "code" in box
            self.wfile.write(("<h2>%s</h2><p>You can close this tab.</p>" % (
                "Drive access granted." if ok else "Authorization failed."
            )).encode())

        def log_message(self, *a):
            pass

    server = HTTPServerV4(("127.0.0.1", 0), Handler)
    redirect = "http://127.0.0.1:%d" % server.server_address[1]
    url = AUTH_URL + "?" + urllib.parse.urlencode({
        "client_id": client_id,
        "redirect_uri": redirect,
        "response_type": "code",
        "scope": SCOPE,
        "access_type": "offline",
        "prompt": "consent",
        "state": state,
        "code_challenge": challenge,
        "code_challenge_method": "S256",
        "login_hint": account,
    })
    print("Approve Drive access for %s in the browser window that opens.\n"
          "If it does not open, paste this URL:\n\n%s\n" % (account, url))
    webbrowser.open(url)
    server.handle_request()
    if box.get("state") != state or "code" not in box:
        sys.exit("Authorization failed: %s" % box.get("error", "no code returned"))
    tok = post_form(TOKEN_URL, {
        "client_id": client_id,
        "client_secret": client_secret,
        "code": box["code"],
        "code_verifier": verifier,
        "grant_type": "authorization_code",
        "redirect_uri": redirect,
    })
    if "refresh_token" not in tok:
        sys.exit("Google returned no refresh token; retry (consent must be re-granted).")
    save_private(token_path(account), {
        "refresh_token": tok["refresh_token"],
        "access_token": tok["access_token"],
        "expires_at": time.time() + int(tok.get("expires_in", 3600)),
    })
    print("Authenticated %s. Token stored in %s (drive.file scope only)."
          % (account, token_path(account)))


class HTTPServerV4(http.server.HTTPServer):
    # IPv4 only: Google will not redirect to an IPv6 loopback address.
    address_family = socket.AF_INET


# ----------------------------------------------------------------- drive calls

RATE_LIMIT_REASONS = {"rateLimitExceeded", "userRateLimitExceeded"}


def is_transient(code, body):
    """Whether a Google error response is worth a retry: a 429 or 5xx, or a 403 that Drive uses
    for a rate limit, which its error reason tells apart from a refusal."""
    if code in (429, 500, 502, 503, 504):
        return True
    if code != 403:
        return False
    try:
        errors = json.loads(body)["error"]["errors"]
        return any(e.get("reason") in RATE_LIMIT_REASONS for e in errors)
    except (ValueError, KeyError, TypeError, AttributeError):
        return False


class HttpError(RuntimeError):
    def __init__(self, code, detail):
        super().__init__("HTTP %s: %s" % (code, detail))
        self.code = code


# An export refusal that Google takes longer than this to answer is final, not transient: Drive
# refuses the export of a sheet that it cannot convert after about 240 s, and again on every retry.
SLOW_REFUSAL_SECONDS = 60


def api(account, req, expect_json=True, retries=5, slow_refusal_is_final=False):
    """Send a prepared Request with auth, retrying transient Google failures.

    Drive answers a small fraction of export requests with a 500 that succeeds
    on a retry, so backoff here is load-bearing, not defensive padding. With
    slow_refusal_is_final, for an export, one of EXPORT_REFUSALS that took over
    SLOW_REFUSAL_SECONDS is raised at once.
    """
    for attempt in range(retries):
        req.add_header("Authorization", "Bearer " + get_access_token(account))
        started = time.monotonic()
        try:
            with urllib.request.urlopen(req, timeout=300) as res:
                body = res.read()
            return json.loads(body) if expect_json else body
        except urllib.error.HTTPError as e:
            body = e.read()
            slow_refusal = (slow_refusal_is_final and e.code in EXPORT_REFUSALS
                            and time.monotonic() - started > SLOW_REFUSAL_SECONDS)
            if not is_transient(e.code, body) or slow_refusal or attempt == retries - 1:
                detail = body[:300].decode("utf8", "replace")
                raise HttpError(e.code, detail) from None
            time.sleep(min(2 ** attempt, 30) + secrets.randbelow(1000) / 1000)
        except (urllib.error.URLError, TimeoutError):
            if attempt == retries - 1:
                raise
            time.sleep(min(2 ** attempt, 30))
    raise RuntimeError("unreachable")


def multipart(metadata, payload, payload_mime):
    boundary = "===" + secrets.token_hex(16) + "==="
    body = b"".join([
        ("--%s\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n"
         % boundary).encode(),
        json.dumps(metadata).encode(),
        ("\r\n--%s\r\nContent-Type: %s\r\n\r\n" % (boundary, payload_mime)).encode(),
        payload,
        ("\r\n--%s--\r\n" % boundary).encode(),
    ])
    return body, "multipart/related; boundary=%s" % boundary


def ensure_folder(account):
    """Our own folder, created under drive.file so its children are ours too."""
    st_path = state_path(account)
    state = json.load(open(st_path)) if os.path.exists(st_path) else {}
    state.setdefault("files", {})
    if state.get("folder_id"):
        return state["folder_id"], state
    req = urllib.request.Request(
        DRIVE_API + "?fields=id",
        data=json.dumps({"name": FOLDER_NAME, "mimeType": FOLDER_MIME}).encode(),
        method="POST")
    req.add_header("Content-Type", "application/json")
    folder = api(account, req)
    state["folder_id"] = folder["id"]
    state.setdefault("files", {})
    save_private(st_path, state)
    return folder["id"], state


def sha256(path):
    with open(path, "rb") as f:
        return hashlib.sha256(f.read()).hexdigest()


def today():
    return time.strftime("%Y-%m-%d", time.gmtime())


def export_is_current(suite, rel, since=None):
    """Does the export's identity match the fixture and export on disk, and, with since, was the
    export made on or after that UTC date?"""
    src = os.path.join(SUITES_DIR, suite, "excel", rel)
    dst = os.path.join(EXPORT_DIR, suite, rel)
    try:
        with open(dst + ".json") as f:
            identity = json.load(f)
        return (identity.get("sourceSha256") == sha256(src)
                and identity.get("exportSha256") == sha256(dst)
                and (since is None or identity.get("exportedOn", "") >= since))
    except (OSError, ValueError):
        return False


def list_children(account, folder_id):
    """Every live child of a folder (only what this app created is visible)."""
    out, token = [], None
    while True:
        params = {"q": "'%s' in parents and trashed = false" % folder_id,
                  "fields": "nextPageToken,files(id,name,mimeType)",
                  "pageSize": 1000}
        if token:
            params["pageToken"] = token
        page = api(account, urllib.request.Request(
            DRIVE_API + "?" + urllib.parse.urlencode(params)))
        out.extend(page.get("files", []))
        token = page.get("nextPageToken")
        if not token:
            return out


def child_folder(account, parent_id, name, cache):
    """Find or create a subfolder, memoised in cache by (parent, name)."""
    key = (parent_id, name)
    if key not in cache:
        found = [c for c in list_children(account, parent_id)
                 if c["name"] == name and c["mimeType"] == FOLDER_MIME]
        if found:
            cache[key] = found[0]["id"]
        else:
            req = urllib.request.Request(
                DRIVE_API + "?fields=id",
                data=json.dumps({"name": name, "mimeType": FOLDER_MIME,
                                 "parents": [parent_id]}).encode(),
                method="POST")
            req.add_header("Content-Type", "application/json")
            cache[key] = api(account, req)["id"]
    return cache[key]


def folder_for(account, root_id, path_dir, cache):
    """The nested folder for a fixture directory such as ironcalc/FINANCIAL."""
    folder = root_id
    for part in [p for p in path_dir.split("/") if p]:
        folder = child_folder(account, folder, part, cache)
    return folder


def drive_index(account, root_id, suites=None):
    """{"suite/rel.xlsx": Sheet id} for everything under the corpus folder, or under the folders
    of those suites. Each folder costs a request, so a suite not named is not walked."""
    index = {}

    def walk(folder_id, prefix):
        for child in list_children(account, folder_id):
            if child["mimeType"] == FOLDER_MIME:
                if prefix == "" and suites is not None and child["name"] not in suites:
                    continue
                walk(child["id"], prefix + child["name"] + "/")
            elif child["mimeType"] == SHEET_MIME:
                name = child["name"]
                if not name.endswith(".xlsx"):
                    name += ".xlsx"
                index.setdefault(prefix + name, child["id"])
    walk(root_id, "")
    return index


def source_is_current(suite, rel):
    """Does the export's identity name the fixture as it is on disk now?"""
    src = os.path.join(SUITES_DIR, suite, "excel", rel)
    dst = os.path.join(EXPORT_DIR, suite, rel)
    try:
        with open(dst + ".json") as f:
            return json.load(f).get("sourceSha256") == sha256(src)
    except (OSError, ValueError):
        return False


NS_MAIN = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"
NS_REL = "{http://schemas.openxmlformats.org/officeDocument/2006/relationships}"
# A cell's formula element, <f .../> or <f ...>text</f>; the lookahead keeps out other tags.
FORMULA = re.compile(rb"<f(?=[\s/>])[^>]*?(?:/>|>.*?</f>)", re.S)


def sheet_parts(payload):
    """[(sheet name, zip part)] of an xlsx, in workbook order."""
    with zipfile.ZipFile(io.BytesIO(payload)) as z:
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        rels = ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))
    targets = {r.get("Id"): r.get("Target") for r in rels}
    out = []
    for sheet in workbook.iter(NS_MAIN + "sheet"):
        target = targets[sheet.get(NS_REL + "id")]
        part = target[1:] if target.startswith("/") else posixpath.normpath("xl/" + target)
        out.append((sheet.get("name"), part))
    return out


def strip_formulas(payload, parts):
    """The xlsx with the formulas of those sheet parts removed, so their cells hold Excel's cached
    values as constants. The sheets stay, so references into them still resolve."""
    out = io.BytesIO()
    with zipfile.ZipFile(io.BytesIO(payload)) as src, \
            zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as dst:
        for info in src.infolist():
            data = src.read(info)
            if info.filename in parts:
                data = FORMULA.sub(b"", data)
            dst.writestr(info, data)
    return out.getvalue()


# A quoted sheet name or a string literal, whichever starts first, so that the quotes of one never
# open the other: `'Say "hi"'!A1`, `"it's"`.
QUOTED_NAME_OR_STRING = re.compile(r"""('(?:[^']|'')*')|"(?:[^"]|"")*\"""")


def blank_strings(formula):
    """The formula with each string literal emptied and its quoted sheet names kept."""
    return QUOTED_NAME_OR_STRING.sub(lambda m: m.group(1) or '""', formula)


QUOTED_PREFIX = re.compile(r"'((?:[^']|'')+)'!")
UNQUOTED_PREFIX = re.compile(r"([^\s!'\"(),;=+\-*/^&<>{}%]+)!")


def referenced_sheets(formula, sheets):
    """The sheets of the workbook, in lower case, that a formula's references name: `Sheet!A1`,
    `'My sheet'!A1`, and every sheet of a 3D reference `First:Last!A1`. A reference into another
    workbook (`[1]Sheet!A1`) names none."""
    order = [n.lower() for n in sheets]
    text = blank_strings(formula)
    prefixes = [m.group(1).replace("''", "'") for m in QUOTED_PREFIX.finditer(text)]
    prefixes += [m.group(1) for m in UNQUOTED_PREFIX.finditer(QUOTED_PREFIX.sub("", text))]
    out = set()
    for prefix in prefixes:
        if prefix.startswith("["):
            continue
        parts = [p.lower() for p in prefix.split(":")]
        named = [p for p in parts if p in order]
        if len(named) == 2:
            first, last = sorted(order.index(p) for p in named)
            out.update(order[first:last + 1])
        else:
            out.update(named)
    return out


def table_names(z, parts):
    """[(table name, sheet name)] of the tables on those (sheet name, zip part) sheets."""
    out = []
    for sheet, part in parts:
        rels = posixpath.join(posixpath.dirname(part), "_rels", posixpath.basename(part) + ".rels")
        if rels not in z.namelist():
            continue
        for r in ET.fromstring(z.read(rels)):
            target = r.get("Target") or ""
            if not (r.get("Type") or "").endswith("/table") or r.get("TargetMode") == "External":
                continue
            path = target[1:] if target.startswith("/") else posixpath.normpath(
                posixpath.join(posixpath.dirname(part), target))
            out.append((ET.fromstring(z.read(path)).get("displayName"), sheet))
    return out


def dependent_sheets(payload, dropped):
    """The sheets, other than the dropped ones, with a formula that may read a dropped sheet,
    directly or through another such sheet, defined names or a table on it; any formula with
    INDIRECT may. In an
    export without the dropped sheets' formulas, those read Excel's values, not Google's. Judged
    per sheet from the formula text, so it errs towards more sheets."""
    if not dropped:
        return []
    parts = sheet_parts(payload)
    names = [n for n, _ in parts]
    with zipfile.ZipFile(io.BytesIO(payload)) as z:
        workbook = ET.fromstring(z.read("xl/workbook.xml"))
        formulas = {n: [html.unescape(m.group(1).decode())
                        for m in re.finditer(rb"<f(?=[\s>])[^>]*(?<!/)>(.*?)</f>", z.read(p), re.S)]
                    for n, p in parts}
        tables = table_names(z, parts)
    defined = [(d.get("name"), d.text or "") for d in workbook.iter(NS_MAIN + "definedName")]
    tainted = {n.lower() for n in dropped}
    bad_names = []

    def may_read_tainted(text):
        return (referenced_sheets(text, names) & tainted or "INDIRECT(" in text.upper()
                or any(b.search(blank_strings(text)) for b in bad_names))

    while True:
        # A name may read a tainted sheet through another name, so the names grow to a fixpoint.
        while True:
            grown_names = [re.compile(r"(?<![\w.\\])%s(?![\w.])" % re.escape(name), re.I)
                           for name in [n for n, text in defined if may_read_tainted(text)]
                           + [n for n, sheet in tables if sheet.lower() in tainted]]
            if len(grown_names) == len(bad_names):
                break
            bad_names = grown_names
        grown = {n.lower() for n in names if n.lower() not in tainted
                 and any(may_read_tainted(f) for f in formulas[n])}
        if not grown:
            break
        tainted |= grown
    return [n for n in names if n.lower() in tainted and n not in dropped]


# Google answers the export of a Sheet that it cannot convert back to xlsx with a 500 or with a 503
# "Transient failure", which persist through retries.
EXPORT_REFUSALS = (500, 503)


class ExportRefused(RuntimeError):
    """Google answered an export with one of EXPORT_REFUSALS after its retries."""


def export_once(account, parent, name, payload):
    """Upload the xlsx as a temporary Sheet, export it, and delete the Sheet. Raises ExportRefused
    when Google refuses the export; a failed upload or delete raises as any request does."""
    body, content_type = multipart({"name": name, "mimeType": SHEET_MIME, "parents": [parent]},
                                   payload, XLSX_MIME)
    req = urllib.request.Request(UPLOAD_API + "?uploadType=multipart&fields=id",
                                 data=body, method="POST")
    req.add_header("Content-Type", content_type)
    file_id = api(account, req)["id"]
    try:
        return api(account, urllib.request.Request(
            "%s/%s/export?%s" % (DRIVE_API, file_id,
                                 urllib.parse.urlencode({"mimeType": XLSX_MIME}))),
            expect_json=False, retries=3, slow_refusal_is_final=True)
    except HttpError as e:
        if e.code not in EXPORT_REFUSALS:
            raise
        raise ExportRefused(str(e)) from None
    finally:
        api(account, urllib.request.Request("%s/%s" % (DRIVE_API, file_id), method="DELETE"),
            expect_json=False)


def export_dropping_sheets(account, parent, name, payload):
    """Export the fixture with the formulas of every sheet that Google can export, for a fixture
    whose export Google keeps refusing (EXPORT_REFUSALS). Adds the sheets with formulas one at a
    time, each trial a temporary Sheet of the fixture without the formulas of the sheets neither
    kept nor on trial, and keeps a sheet when its trial exports. Returns the last export that
    succeeded and the names of the sheets whose formulas it lacks."""
    with zipfile.ZipFile(io.BytesIO(payload)) as z:
        sheets = [(n, p) for n, p in sheet_parts(payload) if FORMULA.search(z.read(p))]
    kept, dropped, data = set(), [], None
    for sheet, part in sheets:
        trial = strip_formulas(payload, {p for n, p in sheets if n not in kept and n != sheet})
        try:
            data = export_once(account, parent, "%s (export trial)" % name, trial)
        except ExportRefused:
            dropped.append(sheet)
            continue
        kept.add(sheet)
    if not kept:
        raise RuntimeError("no sheet with formulas exports: %s" % ", ".join(dropped))
    return data, dropped


def roundtrip_one(account, root_id, suite, rel, index, cache, fresh=False):
    """Make the export current: upload the fixture as a Sheet unless Drive
    already holds this version of it and not fresh, then export that Sheet back,
    or where Google refuses, only the sheets it can export
    (export_dropping_sheets), and write the capture identity beside the export. Returns (Sheet id,
    uploaded)."""
    src = os.path.join(SUITES_DIR, suite, "excel", rel)
    dst = os.path.join(EXPORT_DIR, suite, rel)
    path = suite + "/" + rel
    with open(src, "rb") as f:
        payload = f.read()
    file_id = index.get(path)
    uploaded = False
    if fresh or file_id is None or not source_is_current(suite, rel):
        parent = folder_for(account, root_id, os.path.dirname(path), cache)
        metadata = {"name": os.path.basename(rel)[:-len(".xlsx")],
                    "mimeType": SHEET_MIME}
        if file_id is None:
            metadata["parents"] = [parent]
            url = UPLOAD_API + "?uploadType=multipart&fields=id"
            method = "POST"
        else:  # a changed or fresh fixture replaces the Sheet in place
            url = "%s/%s?uploadType=multipart&fields=id" % (UPLOAD_API, file_id)
            method = "PATCH"
        body, content_type = multipart(metadata, payload, XLSX_MIME)
        req = urllib.request.Request(url, data=body, method=method)
        req.add_header("Content-Type", content_type)
        # A failed response can still leave Drive holding the new fixture.
        # Retire the old identity before making that remote change.
        try:
            os.remove(dst + ".json")
        except FileNotFoundError:
            pass
        file_id = api(account, req)["id"]
        index[path] = file_id
        uploaded = True
    export = urllib.request.Request(
        "%s/%s/export?%s" % (DRIVE_API, file_id,
                             urllib.parse.urlencode({"mimeType": XLSX_MIME})))
    dropped = []
    try:
        data = api(account, export, expect_json=False, slow_refusal_is_final=True)
    except HttpError as e:
        if e.code not in EXPORT_REFUSALS:
            raise
        # The Sheet on Drive stays the whole fixture; only the export lacks some sheets' formulas.
        parent = folder_for(account, root_id, os.path.dirname(path), cache)
        data, dropped = export_dropping_sheets(
            account, parent, os.path.basename(rel)[:-len(".xlsx")], payload)
        if dropped:
            with _print_lock:
                print("  DROPPED SHEETS %s: Google failed to export the formulas of %s"
                      % (path, "; ".join(dropped)))
    os.makedirs(os.path.dirname(dst), exist_ok=True)
    identity = {"sourceSha256": hashlib.sha256(payload).hexdigest(),
                "exportSha256": hashlib.sha256(data).hexdigest(),
                "exportedOn": today()}
    if dropped:
        identity["droppedSheets"] = dropped
        dependent = dependent_sheets(payload, dropped)
        if dependent:
            identity["dependentSheets"] = dependent
    identity = json.dumps(identity)
    for path_, content in ((dst, data), (dst + ".json", identity.encode())):
        with open(path_ + ".part", "wb") as f:
            f.write(content)
        os.replace(path_ + ".part", path_)
    return file_id, uploaded


def fixtures(suite):
    root = os.path.join(SUITES_DIR, suite, "excel")
    out = []
    for dirpath, _, names in os.walk(root):
        for n in sorted(names):
            if n.endswith(".xlsx"):
                out.append(os.path.relpath(os.path.join(dirpath, n), root))
    return sorted(out)


def is_private_suite(suite):
    return suite.startswith("private-")


def is_real_workbook_suite(suite):
    # Runs only when named, like the other harnesses do (REAL_WORKBOOKS in src/suites.ts).
    try:
        with open(os.path.join(SUITES_DIR, suite, "manifest.json")) as f:
            return json.load(f).get("kind") == "real-workbooks"
    except (OSError, ValueError):
        return False


def available_suites():
    """Every suite that may be uploaded, as `run` accepts them by name."""
    return sorted(s for s in os.listdir(SUITES_DIR)
                  if os.path.isdir(os.path.join(SUITES_DIR, s, "excel"))
                  and not is_private_suite(s))


def all_suites():
    """The suites of --all: the available ones but real-workbook suites."""
    return [s for s in available_suites() if not is_real_workbook_suite(s)]


def cmd_run(account, suites, since=None):
    folder_id, state = ensure_folder(account)
    index = drive_index(account, folder_id, suites)
    cache = {}
    failed = 0
    fixture_count = 0
    for suite in suites:
        started = time.time()
        todo = [r for r in fixtures(suite)
                if suite + "/" + r not in index or not export_is_current(suite, r, since)]
        total = len(fixtures(suite))
        fixture_count += total
        if not todo:
            print("%-22s %d/%d on Drive, exports current" % (suite, total, total))
            continue
        print("%-22s %d to do (of %d)" % (suite, len(todo), total))
        # Create the folder chain up front, single-threaded, so the pool never
        # races to create the same subfolder twice.
        for r in todo:
            folder_for(account, folder_id, os.path.dirname(suite + "/" + r), cache)
        done = [0]
        uploads = [0]
        failures = []

        def work(rel):
            try:
                fid, uploaded = roundtrip_one(account, folder_id, suite, rel, index, cache,
                                              fresh=since is not None)
                state["files"][suite + "/" + rel] = fid
                uploads[0] += uploaded
            except Exception as e:  # keep going; one file is not the run
                failures.append((rel, str(e)[:160]))
            done[0] += 1
            if done[0] % 25 == 0 or done[0] == len(todo):
                with _print_lock:
                    rate = done[0] / max(time.time() - started, 1e-9)
                    print("  %s: %d/%d (%.1f files/s)"
                          % (suite, done[0], len(todo), rate))

        with concurrent.futures.ThreadPoolExecutor(WORKERS) as pool:
            list(pool.map(work, todo))
        save_private(state_path(account), state)
        for rel, err in failures:
            print("  FAILED %s/%s: %s" % (suite, rel, err))
        print("%-22s done: %d exported (%d uploaded), %d failed"
              % (suite, len(todo) - len(failures), uploads[0], len(failures)))
        failed += len(failures)
    save_private(state_path(account), state)
    if failed and failed == fixture_count:
        # A fixture skipped as current has a usable export, so only a run that left none fails.
        sys.exit("All %d fixtures failed, so none has a current export. Run again to retry them."
                 % failed)
    if failed:
        print("\n%d fixtures failed. Run again to retry them; current exports are skipped."
              % failed, file=sys.stderr)
        sys.exit(PARTIAL_EXIT)
    print("\nExports in %s, which replay/google-sheets.ts scores." % EXPORT_DIR)


def cmd_migrate(account):
    """Move Sheets named "suite/dir/file" flat in the corpus folder into the
    suite subfolders the layout uses, renamed to the bare file name. A second
    run finds nothing to move."""
    folder_id, _ = ensure_folder(account)
    flat = [c for c in list_children(account, folder_id)
            if c["mimeType"] == SHEET_MIME and "/" in c["name"]]
    if not flat:
        print("Nothing to migrate: no flat names in %s." % FOLDER_NAME)
        return
    print("%d flat names to move into subfolders" % len(flat))
    cache = {}
    for c in flat:  # folders first, single-threaded
        folder_for(account, folder_id, os.path.dirname(c["name"]), cache)
    done = [0]
    failures = []

    def work(c):
        try:
            parent = folder_for(account, folder_id, os.path.dirname(c["name"]), cache)
            name = os.path.basename(c["name"])
            if name.endswith(".xlsx"):
                name = name[:-len(".xlsx")]
            req = urllib.request.Request(
                "%s/%s?%s" % (DRIVE_API, c["id"], urllib.parse.urlencode(
                    {"addParents": parent, "removeParents": folder_id, "fields": "id"})),
                data=json.dumps({"name": name}).encode(), method="PATCH")
            req.add_header("Content-Type", "application/json")
            api(account, req)
        except Exception as e:
            failures.append((c["name"], str(e)[:160]))
        done[0] += 1
        if done[0] % 100 == 0 or done[0] == len(flat):
            with _print_lock:
                print("  moved %d/%d" % (done[0], len(flat)))

    with concurrent.futures.ThreadPoolExecutor(WORKERS) as pool:
        list(pool.map(work, flat))
    for name, err in failures:
        print("  FAILED %s: %s" % (name, err))
    print("Migrated %d, %d failed." % (len(flat) - len(failures), len(failures)))


def cmd_status(account):
    index = None
    if os.path.exists(token_path(account)):
        st_path = state_path(account)
        state = json.load(open(st_path)) if os.path.exists(st_path) else {}
        if state.get("folder_id"):
            index = drive_index(account, state["folder_id"], all_suites())
    print("%-22s %9s %9s" % ("suite", "current", "on Drive"))
    for suite in all_suites():
        fx = fixtures(suite)
        have = len([r for r in fx if export_is_current(suite, r)])
        on_drive = "?" if index is None else len([r for r in fx if suite + "/" + r in index])
        flag = "" if have == len(fx) and on_drive in ("?", len(fx)) else "   <-- incomplete"
        print("%-22s %4d/%4d %4s/%4d%s" % (suite, have, len(fx), on_drive, len(fx), flag))


def cmd_cleanup(account):
    """Trash the folder we created; drive.file reaches nothing else."""
    st_path = state_path(account)
    if not os.path.exists(st_path):
        sys.exit("Nothing recorded to clean up.")
    state = json.load(open(st_path))
    folder_id = state.get("folder_id")
    if not folder_id:
        sys.exit("No folder recorded.")
    req = urllib.request.Request(
        "%s/%s" % (DRIVE_API, folder_id),
        data=json.dumps({"trashed": True}).encode(), method="PATCH")
    req.add_header("Content-Type", "application/json")
    api(account, req)
    print("Trashed folder %s (%d files)." % (FOLDER_NAME, len(state.get("files", {}))))
    state["files"] = {}
    state.pop("folder_id", None)
    save_private(st_path, state)


def main():
    args = sys.argv[1:]
    if not args:
        sys.exit(__doc__)
    cmd, rest = args[0], args[1:]
    account = ACCOUNT
    if cmd == "auth":
        cmd_auth(rest[0] if rest else account)
    elif cmd == "run":
        since = None
        if "--since" in rest:
            i = rest.index("--since")
            since = rest[i + 1] if i + 1 < len(rest) else ""
            if not re.fullmatch(r"\d{4}-\d{2}-\d{2}", since):
                sys.exit("--since takes a UTC date, YYYY-MM-DD")
            rest = rest[:i] + rest[i + 2:]
        suites = all_suites() if rest == ["--all"] else rest
        if not suites:
            sys.exit("Name at least one suite, or --all")
        private = [s for s in suites if is_private_suite(s)]
        if private:
            sys.exit("Private suite(s), never uploaded: %s" % ", ".join(private))
        available = available_suites()
        unknown = [s for s in suites if s not in available]
        if unknown:
            sys.exit("Unknown suite(s): %s" % ", ".join(unknown))
        cmd_run(account, suites, since)
    elif cmd == "status":
        cmd_status(account)
    elif cmd == "ready":
        cmd_ready(account)
    elif cmd == "migrate":
        cmd_migrate(account)
    elif cmd == "cleanup":
        cmd_cleanup(account)
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
