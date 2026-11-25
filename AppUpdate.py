"""
Years at FLAME — Yearbook Photoshoot Registration Server
=========================================================

Runs locally on one laptop and saves every registration + photocode into
one shared JSONL log and local Excel workbook (registrations.xlsx).

SETUP (once)
    pip install flask openpyxl waitress

RUN (on the laptop that will collect the data)
    python AppUpdate.py

It prints:
    Student Link  : your constant link (PUBLIC_URL / Tailscale / ngrok / .local)
    LAN Fallback  : http://<this laptop's IP>:5000
    Admin Console : http://localhost:5000/admin   (user: admin, password shown)

HOW IT WORKS
    Students go through five steps: details, batch, studies, Student ID (typed
    twice), then a review screen. UG3 and UG4 students pick a major and a minor
    from MAJOR_MINORS. MSC, PG2 and PGPEI students type a specialisation instead.

    Students never see the photocode. On the admin page, press "Add photocode"
    next to a student, type the code twice and save. The code is checked for
    format, must not belong to anyone else, and is time-stamped. Once a student
    has a confirmed photocode, they can no longer change their entry online.

    You can download the Excel file from the admin page at any time. A copy is
    also kept next to this script as registrations.xlsx.

Files created next to this script:
    registrations.jsonl   the permanent log (never edit by hand)
    registrations.xlsx    an Excel copy, rebuilt after every change

Optional settings (.env next to this script, or environment variables):
    PUBLIC_URL           constant public HTTPS URL (e.g. Tailscale Funnel / ngrok)
    NGROK_DOMAIN         optional static ngrok domain if using ngrok
    PORT                 default 5000
    ADMIN_PASSWORD       default: a random one printed at start-up
    PHOTOCODE_PATTERN    a regular expression every photocode must match, for
                         example  [A-Z]{2}\\d{4}  (default: 3 to 30 characters,
                         no spaces)
"""
import hmac
import io
import json
import os
import re
import secrets
import socket
import subprocess
import sys
import threading
import time
import urllib.request
from datetime import datetime
from functools import wraps

from flask import Flask, Response, jsonify, request, send_file
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

HERE = os.path.dirname(os.path.abspath(__file__))


def _load_dotenv(path):
    """Load simple KEY=VALUE pairs from .env without overriding existing env vars."""
    if not os.path.isfile(path):
        return
    try:
        with open(path, encoding="utf-8") as f:
            for raw in f:
                line = raw.strip()
                if not line or line.startswith("#") or "=" not in line:
                    continue
                k, v = line.split("=", 1)
                k = k.strip()
                v = v.strip()
                if len(v) >= 2 and v[0] == v[-1] and v[0] in ("'", '"'):
                    v = v[1:-1]
                if k and k not in os.environ and v != "":
                    os.environ[k] = v
    except OSError:
        pass


_load_dotenv(os.path.join(HERE, ".env"))

DATA_FILE = os.path.join(HERE, "registrations.jsonl")
XLSX_FILE = os.path.join(HERE, "registrations.xlsx")
PORT = int(os.environ.get("PORT", "5000"))
ADMIN_PASSWORD = os.environ.get("ADMIN_PASSWORD") or secrets.token_urlsafe(6)
TURSO_DATABASE_URL = (os.environ.get("TURSO_DATABASE_URL") or "").strip()
TURSO_AUTH_TOKEN = (os.environ.get("TURSO_AUTH_TOKEN") or "").strip()
USE_TURSO = bool(TURSO_DATABASE_URL and TURSO_AUTH_TOKEN)

PHOTOCODE_PATTERN = os.environ.get("PHOTOCODE_PATTERN") or r"\S{3,30}"
try:
    PHOTO_RE = re.compile(PHOTOCODE_PATTERN)
except re.error as exc:
    sys.exit("PHOTOCODE_PATTERN is not a valid pattern: %s" % exc)

# Cleanly masked transparent-edge Years at FLAME logo PNG (96x96)
LOGO_DATA_URI = (
    "data:image/png;base64,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"
)

# --------------------------------------------------------------------- batches
BATCHES = ["UG3", "UG4", "MSC", "PG2", "PGPEI"]
UG_BATCHES = {"UG3", "UG4"}  # these batches choose a major and a minor

# ------------------------------------------------------- major / minor table
# Copied from the programme table. Each major lists the minors it can be
# combined with. An empty list means the programme offers no minor (N/A).
HRM = "Human Resource Management (HRM)"
LCS = "Literary & Cultural Studies (LCS)"
ADV = "Advertising & Branding"
DMAC = "Digital Marketing & Communications (DMAC)"
EVS = "Environmental Studies (EVS)"
IS = "International Studies (IS)"

BUSINESS = ["Finance", "Marketing", "Operations", "Business Analytics", "Entrepreneurship", HRM, LCS]
QUANT = ["Applied Mathematics", "Public Policy", "Psychology"]
MEDIA = [ADV, DMAC, "Film & Television Management", "Journalism", EVS]
SOCIAL = [IS, "Sociology", "Economics", "Philosophy"]
COMPSCI = ["Computer Science"]
ARTS = ["Dance", "Music", "Theatre & Performance Studies", "Design"]


def _minors(*groups, drop=()):
    return [m for g in groups for m in g if m not in drop]


_FINANCE_LIKE = _minors(QUANT, MEDIA, SOCIAL, COMPSCI, ARTS)
_PEOPLE_LIKE = _minors(QUANT, MEDIA, SOCIAL, ARTS)
_QUANT_LIKE = _minors(BUSINESS, MEDIA, SOCIAL, COMPSCI, ARTS)
_ADVERTISING_LIKE = _minors(BUSINESS, QUANT, SOCIAL, ARTS, drop=("Marketing",))
_MEDIA_LIKE = _minors(BUSINESS, QUANT, SOCIAL, ARTS)
_SOCIAL_LIKE = _minors(BUSINESS, QUANT, MEDIA)

MAJOR_MINORS = {
    "Finance": _FINANCE_LIKE,
    "Marketing": _minors(QUANT, MEDIA, SOCIAL, COMPSCI, ARTS, drop=(ADV, DMAC)),
    "Operations": _FINANCE_LIKE,
    "Business Analytics": _FINANCE_LIKE,
    "Entrepreneurship": _FINANCE_LIKE,
    HRM: _PEOPLE_LIKE,
    "General Management": _PEOPLE_LIKE,
    LCS: _PEOPLE_LIKE,
    "Public Policy": _QUANT_LIKE,
    "Applied Mathematics": _QUANT_LIKE,
    "Psychology": _QUANT_LIKE,
    ADV: _ADVERTISING_LIKE,
    DMAC: _ADVERTISING_LIKE,
    "Film & Television Management": _MEDIA_LIKE,
    "Communication Studies": _MEDIA_LIKE,
    "Journalism": _MEDIA_LIKE,
    EVS: _MEDIA_LIKE,
    IS: _SOCIAL_LIKE,
    "Sociology": _SOCIAL_LIKE,
    "Economics": _SOCIAL_LIKE,
    "Computer Science": _minors(BUSINESS, QUANT, drop=(HRM, LCS)),
    "Computer Science & Design": [],
    "Design Management": [],
    "Data Science & Economics": [],
}
CANON_MAJOR = {m.casefold(): m for m in MAJOR_MINORS}

LOCK = threading.Lock()
COLUMNS = ["No.", "Submitted", "Name", "Mobile", "Email", "Batch", "Major", "Minor",
           "Specialisation", "Student ID", "Photocode", "Photocode added"]

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 32 * 1024


@app.after_request
def add_headers(resp):
    resp.headers["X-Content-Type-Options"] = "nosniff"
    resp.headers["Referrer-Policy"] = "no-referrer"
    resp.headers["Cache-Control"] = "no-store"
    resp.headers["Content-Security-Policy"] = (
        "default-src 'self'; script-src 'unsafe-inline'; style-src 'unsafe-inline'; "
        "connect-src 'self'; img-src 'self' data:; base-uri 'none'; form-action 'self'")
    return resp


# -------------------------------------------------- constant link resolution
def lan_address():
    try:
        s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
        s.connect(("10.255.255.255", 1))  # no packet is sent
        ip = s.getsockname()[0]
        s.close()
        return ip
    except OSError:
        return "127.0.0.1"


def _mdns_hostname():
    if sys.platform == "darwin":
        try:
            r = subprocess.run(["scutil", "--get", "LocalHostName"], capture_output=True, text=True, timeout=0.5)
            h = r.stdout.strip()
            if h:
                return re.sub(r"[^A-Za-z0-9\-]", "", h) + ".local"
        except Exception:
            pass
    host = socket.gethostname().split(".")[0]
    host = re.sub(r"[^A-Za-z0-9\-]", "", host) or "localhost"
    return host + ".local"


def _detect_ngrok():
    try:
        req = urllib.request.Request("http://127.0.0.1:4040/api/tunnels", headers={"Accept": "application/json"})
        with urllib.request.urlopen(req, timeout=0.4) as resp:
            data = json.loads(resp.read().decode("utf-8", "ignore"))
        for t in data.get("tunnels", []):
            u = str(t.get("public_url", ""))
            if u.startswith("https://"):
                return u.rstrip("/")
        for t in data.get("tunnels", []):
            u = str(t.get("public_url", ""))
            if u.startswith("http://"):
                return u.rstrip("/")
    except Exception:
        pass
    return ""


def _detect_tailscale():
    ts_env = (os.environ.get("TS_FUNNEL_URL") or "").strip().rstrip("/")
    if ts_env:
        return ts_env
    try:
        r = subprocess.run(["tailscale", "funnel", "status", "--json"], capture_output=True, text=True, timeout=0.8)
        if r.returncode == 0 and r.stdout.strip():
            data = json.loads(r.stdout)
            web = data.get("Web") or {}
            allow = data.get("AllowFunnel") or {}
            for host_port, enabled in allow.items():
                if enabled and host_port in web:
                    host = host_port.split(":")[0].rstrip(".")
                    if host:
                        return "https://" + host
    except Exception:
        pass
    return ""


def resolve_student_urls():
    """Determine the best student link (constant public tunnel -> .local mDNS -> LAN IP)."""
    lan_ip = lan_address()
    lan_url = "http://%s:%d" % (lan_ip, PORT)
    mdns_host = _mdns_hostname()
    mdns_url = "http://%s" % mdns_host if PORT == 80 else "http://%s:%d" % (mdns_host, PORT)

    env_public = (os.environ.get("PUBLIC_URL") or "").strip().rstrip("/")
    ngrok_domain = (os.environ.get("NGROK_DOMAIN") or "").strip().rstrip("/")
    if not env_public and ngrok_domain:
        env_public = ngrok_domain if ngrok_domain.startswith("http") else ("https://" + ngrok_domain)

    if env_public:
        return {
            "primaryUrl": env_public,
            "mode": "custom",
            "label": "Constant public URL (.env)",
            "isConstant": True,
            "mdnsUrl": mdns_url,
            "lanUrl": lan_url,
            "publicUrl": env_public,
        }

    vercel_host = (os.environ.get("VERCEL_PROJECT_PRODUCTION_URL") or os.environ.get("VERCEL_URL") or "").strip().rstrip("/")
    if vercel_host:
        v_url = vercel_host if vercel_host.startswith("http") else ("https://" + vercel_host)
        return {
            "primaryUrl": v_url,
            "mode": "vercel",
            "label": "Vercel production URL (constant HTTPS)",
            "isConstant": True,
            "mdnsUrl": mdns_url,
            "lanUrl": lan_url,
            "publicUrl": v_url,
        }

    ts_url = _detect_tailscale()
    if ts_url:
        return {
            "primaryUrl": ts_url,
            "mode": "tailscale",
            "label": "Tailscale Funnel (constant HTTPS)",
            "isConstant": True,
            "mdnsUrl": mdns_url,
            "lanUrl": lan_url,
            "publicUrl": ts_url,
        }

    ngrok_url = _detect_ngrok()
    if ngrok_url:
        return {
            "primaryUrl": ngrok_url,
            "mode": "ngrok",
            "label": "ngrok public tunnel",
            "isConstant": True,
            "mdnsUrl": mdns_url,
            "lanUrl": lan_url,
            "publicUrl": ngrok_url,
        }

    if mdns_host != "localhost.local":
        return {
            "primaryUrl": mdns_url,
            "mode": "mdns",
            "label": "Constant local hostname (.local)",
            "isConstant": True,
            "mdnsUrl": mdns_url,
            "lanUrl": lan_url,
            "publicUrl": "",
        }

    return {
        "primaryUrl": lan_url,
        "mode": "lan",
        "label": "LAN IP (changes across Wi-Fi)",
        "isConstant": False,
        "mdnsUrl": mdns_url,
        "lanUrl": lan_url,
        "publicUrl": "",
    }


# ------------------------------------------------ pure-Python QR Code SVG (V1-10)
_QR_SPECS = [
    (1, 26, 7, 1, 19, 0, 0, []),
    (2, 44, 10, 1, 34, 0, 0, [6, 18]),
    (3, 70, 15, 1, 55, 0, 0, [6, 22]),
    (4, 100, 20, 1, 80, 0, 0, [6, 26]),
    (5, 134, 26, 1, 108, 0, 0, [6, 30]),
    (6, 172, 18, 2, 68, 0, 0, [6, 34]),
    (7, 196, 20, 2, 78, 0, 0, [6, 22, 38]),
    (8, 242, 24, 2, 97, 0, 0, [6, 24, 42]),
    (9, 292, 30, 2, 116, 0, 0, [6, 26, 46]),
    (10, 346, 18, 2, 68, 2, 69, [6, 28, 50]),
]

_GF_EXP = [0] * 512
_GF_LOG = [0] * 256
_x = 1
for _i in range(255):
    _GF_EXP[_i] = _x
    _GF_LOG[_x] = _i
    _x <<= 1
    if _x & 0x100:
        _x ^= 0x11D
for _i in range(255, 512):
    _GF_EXP[_i] = _GF_EXP[_i - 255]


def _gf_mul(a, b):
    if a == 0 or b == 0:
        return 0
    return _GF_EXP[_GF_LOG[a] + _GF_LOG[b]]


def _rs_gen_poly(degree):
    poly = [1]
    for i in range(degree):
        nxt = [0] * (len(poly) + 1)
        a = _GF_EXP[i]
        for j, c in enumerate(poly):
            nxt[j] ^= c
            nxt[j + 1] ^= _gf_mul(c, a)
        poly = nxt
    return poly


def _rs_encode(data, ec_len):
    gen = _rs_gen_poly(ec_len)
    res = list(data) + [0] * ec_len
    for i in range(len(data)):
        coef = res[i]
        if coef != 0:
            for j in range(1, len(gen)):
                res[i + j] ^= _gf_mul(gen[j], coef)
    return res[len(data):]


def _bch_format(mask_id):
    data = (0b01 << 3) | mask_id
    v = data << 10
    for i in range(4, -1, -1):
        if v & (1 << (i + 10)):
            v ^= 0x537 << i
    return ((data << 10) | v) ^ 0x5412


def _bch_version(ver):
    v = ver << 12
    for i in range(5, -1, -1):
        if v & (1 << (i + 12)):
            v ^= 0x1F25 << i
    return (ver << 12) | v


def _qr_penalty(mat):
    n = len(mat)
    pen = 0
    for r in range(n):
        run = 1
        for c in range(1, n):
            if mat[r][c] == mat[r][c - 1]:
                run += 1
            else:
                if run >= 5:
                    pen += 3 + (run - 5)
                run = 1
        if run >= 5:
            pen += 3 + (run - 5)
    for c in range(n):
        run = 1
        for r in range(1, n):
            if mat[r][c] == mat[r - 1][c]:
                run += 1
            else:
                if run >= 5:
                    pen += 3 + (run - 5)
                run = 1
        if run >= 5:
            pen += 3 + (run - 5)
    for r in range(n - 1):
        for c in range(n - 1):
            v = mat[r][c]
            if v == mat[r][c + 1] == mat[r + 1][c] == mat[r + 1][c + 1]:
                pen += 3
    p1 = [1, 0, 1, 1, 1, 0, 1, 0, 0, 0, 0]
    p2 = [0, 0, 0, 0, 1, 0, 1, 1, 1, 0, 1]
    for r in range(n):
        row = mat[r]
        for c in range(n - 10):
            seg = row[c:c + 11]
            if seg == p1 or seg == p2:
                pen += 40
    for c in range(n):
        for r in range(n - 10):
            seg = [mat[r + k][c] for k in range(11)]
            if seg == p1 or seg == p2:
                pen += 40
    dark = sum(sum(row) for row in mat)
    pct = (dark * 100) // (n * n)
    pen += min(abs(pct - 50) // 5, abs((pct + 4) // 5 * 5 - 50) // 5) * 10
    return pen


def make_qr_matrix(text):
    raw = text.encode("utf-8")
    spec = None
    for s in _QR_SPECS:
        ver, _, _, g1b, g1d, g2b, g2d, _ = s
        total_data = g1b * g1d + g2b * g2d
        count_bits = 8 if ver <= 9 else 16
        needed_bits = 4 + count_bits + len(raw) * 8
        if needed_bits <= total_data * 8:
            spec = s
            break
    if spec is None:
        raise ValueError("URL is too long for QR Version 10 (max 271 bytes).")

    ver, total_cw, ec_len, g1b, g1d, g2b, g2d, align = spec
    total_data = g1b * g1d + g2b * g2d
    count_bits = 8 if ver <= 9 else 16

    bits = []

    def push(val, width):
        for shift in range(width - 1, -1, -1):
            bits.append((val >> shift) & 1)

    push(0b0100, 4)
    push(len(raw), count_bits)
    for b in raw:
        push(b, 8)
    cap_bits = total_data * 8
    term = min(4, cap_bits - len(bits))
    push(0, term)
    while len(bits) % 8 != 0:
        bits.append(0)
    data_cw = []
    for i in range(0, len(bits), 8):
        val = 0
        for b in bits[i:i + 8]:
            val = (val << 1) | b
        data_cw.append(val)
    pad_bytes = (0xEC, 0x11)
    pi = 0
    while len(data_cw) < total_data:
        data_cw.append(pad_bytes[pi & 1])
        pi += 1

    blocks = []
    idx = 0
    for _ in range(g1b):
        blk = data_cw[idx:idx + g1d]
        idx += g1d
        blocks.append((blk, _rs_encode(blk, ec_len)))
    for _ in range(g2b):
        blk = data_cw[idx:idx + g2d]
        idx += g2d
        blocks.append((blk, _rs_encode(blk, ec_len)))

    stream = []
    max_d = max(len(b[0]) for b in blocks)
    for i in range(max_d):
        for blk, _ in blocks:
            if i < len(blk):
                stream.append(blk[i])
    for i in range(ec_len):
        for _, ec in blocks:
            stream.append(ec[i])

    n = 21 + 4 * (ver - 1)
    mat = [[0] * n for _ in range(n)]
    func = [[False] * n for _ in range(n)]

    def set_fn(r, c, v):
        mat[r][c] = 1 if v else 0
        func[r][c] = True

    def place_finder(r0, c0):
        for dr in range(-1, 8):
            for dc in range(-1, 8):
                r, c = r0 + dr, c0 + dc
                if 0 <= r < n and 0 <= c < n:
                    inside = 0 <= dr <= 6 and 0 <= dc <= 6
                    dark = inside and (
                        dr in (0, 6) or dc in (0, 6) or (2 <= dr <= 4 and 2 <= dc <= 4)
                    )
                    set_fn(r, c, dark)

    place_finder(0, 0)
    place_finder(0, n - 7)
    place_finder(n - 7, 0)

    for i in range(8, n - 8):
        set_fn(6, i, i % 2 == 0)
        set_fn(i, 6, i % 2 == 0)

    for r0 in align:
        for c0 in align:
            if func[r0][c0]:
                continue
            for dr in range(-2, 3):
                for dc in range(-2, 3):
                    dark = max(abs(dr), abs(dc)) != 1
                    set_fn(r0 + dr, c0 + dc, dark)

    set_fn(4 * ver + 9, 8, 1)

    for i in range(9):
        if i != 6:
            func[8][i] = True
            func[i][8] = True
    for i in range(8):
        func[8][n - 1 - i] = True
    for i in range(7):
        func[n - 1 - i][8] = True

    if ver >= 7:
        vbits = _bch_version(ver)
        for i in range(18):
            bit = (vbits >> i) & 1
            a = n - 11 + (i % 3)
            b = i // 3
            set_fn(a, b, bit)
            set_fn(b, a, bit)

    bit_stream = []
    for cw in stream:
        for s in range(7, -1, -1):
            bit_stream.append((cw >> s) & 1)

    bi = 0
    up = True
    c = n - 1
    while c > 0:
        if c == 6:
            c -= 1
        rows_iter = range(n - 1, -1, -1) if up else range(n)
        for r in rows_iter:
            for dc in (0, -1):
                cc = c + dc
                if not func[r][cc]:
                    mat[r][cc] = bit_stream[bi] if bi < len(bit_stream) else 0
                    bi += 1
        up = not up
        c -= 2

    mask_fns = [
        lambda r, c: (r + c) % 2 == 0,
        lambda r, c: r % 2 == 0,
        lambda r, c: c % 3 == 0,
        lambda r, c: (r + c) % 3 == 0,
        lambda r, c: ((r // 2) + (c // 3)) % 2 == 0,
        lambda r, c: ((r * c) % 2 + (r * c) % 3) == 0,
        lambda r, c: (((r * c) % 2 + (r * c) % 3) % 2) == 0,
        lambda r, c: (((r + c) % 2 + (r * c) % 3) % 2) == 0,
    ]

    best_mat = None
    best_pen = None
    for mid, mfn in enumerate(mask_fns):
        cand = [row[:] for row in mat]
        for r in range(n):
            for c in range(n):
                if not func[r][c] and mfn(r, c):
                    cand[r][c] ^= 1
        fmt = _bch_format(mid)
        for i in range(15):
            bit = (fmt >> i) & 1
            if i < 6:
                cand[8][i] = bit
            elif i < 8:
                cand[8][i + 1] = bit
            else:
                cand[14 - i][8] = bit
            if i < 7:
                cand[n - 1 - i][8] = bit
            elif i < 8:
                cand[8][n - 8] = bit
            else:
                cand[8][n - 15 + i] = bit
        p = _qr_penalty(cand)
        if best_pen is None or p < best_pen:
            best_pen = p
            best_mat = cand

    return best_mat


def render_qr_svg(text, fg="#011F78", bg="#FFFFFF"):
    mat = make_qr_matrix(text)
    n = len(mat)
    q = 4
    size = n + q * 2
    parts = []
    for r in range(n):
        for c in range(n):
            if mat[r][c]:
                parts.append("M%d,%dh1v1h-1z" % (c + q, r + q))
    return (
        '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {s} {s}" '
        'shape-rendering="crispEdges" role="img" aria-label="QR code">'
        '<rect width="{s}" height="{s}" fill="{bg}"/>'
        '<path d="{d}" fill="{fg}"/></svg>'
    ).format(s=size, bg=bg, fg=fg, d="".join(parts))


# ----------------------------------------------------------------- validation
def clean(s, n=200):
    s = str(s if s is not None else "")
    s = s.encode("utf-8", "ignore").decode("utf-8")  # drops lone surrogates
    s = re.sub(r"[\x00-\x1f\x7f-\x9f]", " ", s)      # control characters break Excel files
    return re.sub(r"\s+", " ", s).strip()[:n]


def key(student_id):
    return clean(student_id, 40).upper()


def prog_ok(v):
    return bool(re.fullmatch(r"[\w &,.'()/+\-]{2,80}", v)) and v[0] not in "=+-@"


def validate(d):
    errors = {}
    name = clean(d.get("name"), 100)
    if len(name) < 2 or not re.search(r"[^\W\d_]", name) or name[0] in "=+-@":
        errors["name"] = "Enter your full name."

    phone = clean(d.get("phone"), 30)
    if not re.fullmatch(r"\+?\d{7,15}", re.sub(r"[\s\-().]", "", phone)):
        errors["phone"] = "Enter a mobile number with 7 to 15 digits."

    email = clean(d.get("email"), 120)
    if not re.fullmatch(r"[^\s@]+@[^\s@]+\.[^\s@]{2,}", email) or email[0] in "=+-@":
        errors["email"] = "Enter a valid email address."

    batch = clean(d.get("batch"), 10).upper()
    major = minor = spec = "N/A"
    if batch not in BATCHES:
        errors["batch"] = "Choose your batch."
    elif batch in UG_BATCHES:
        m = CANON_MAJOR.get(clean(d.get("major"), 80).casefold())
        if not m:
            errors["major"] = "Choose your major from the list."
        else:
            major = m
            allowed = MAJOR_MINORS[m]
            raw = clean(d.get("minor"), 80)
            if not allowed:
                if raw and raw.upper() != "N/A":
                    errors["minor"] = "This major does not offer a minor."
                minor = "N/A"
            elif not raw:
                minor = "None"
            else:
                found = {x.casefold(): x for x in allowed}.get(raw.casefold())
                if found:
                    minor = found
                else:
                    errors["minor"] = "That minor is not offered with this major."
    else:
        s = clean(d.get("specialisation"), 80)
        if not prog_ok(s):
            errors["specialisation"] = "Enter your specialisation."
        else:
            spec = s

    sid = clean(d.get("studentId"), 40)
    if not re.fullmatch(r"[A-Za-z0-9][A-Za-z0-9\-_/.]{2,29}", sid):
        errors["studentId"] = "Enter your Student ID."
    elif sid != clean(d.get("studentId2"), 40):
        errors["studentId2"] = "The Student ID does not match."

    if errors:
        return None, errors
    return {
        "t": "add",
        "submittedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "name": name, "phone": phone, "email": email,
        "batch": batch, "major": major, "minor": minor, "specialisation": spec,
        "studentId": sid,
    }, None


# -------------------------------------------------------------------- storage
class TursoError(RuntimeError):
    pass


_DB_READY = False


def _turso_http_url(url):
    if url.startswith("libsql://"):
        return "https://" + url[len("libsql://"):].rstrip("/")
    return url.rstrip("/")


def _turso_arg(v):
    if v is None:
        return {"type": "null"}
    if isinstance(v, bool):
        return {"type": "integer", "value": "1" if v else "0"}
    if isinstance(v, int):
        return {"type": "integer", "value": str(v)}
    if isinstance(v, float):
        return {"type": "float", "value": v}
    return {"type": "text", "value": str(v)}


def _turso_val(cell):
    if not isinstance(cell, dict):
        return cell
    t = cell.get("type")
    if t == "null":
        return None
    if t == "integer":
        return int(cell.get("value", 0))
    if t == "float":
        return float(cell.get("value", 0.0))
    return cell.get("value", "")


def _turso_parse_rows(result):
    cols = [c.get("name", "") for c in result.get("cols", [])]
    out = []
    for raw_row in result.get("rows", []):
        out.append({cols[i]: _turso_val(raw_row[i]) for i in range(len(cols))})
    return out


def _turso_pipeline(statements):
    endpoint = _turso_http_url(TURSO_DATABASE_URL) + "/v2/pipeline"
    reqs = []
    for item in statements:
        if isinstance(item, str):
            sql, args = item, []
        else:
            sql, args = item
        reqs.append({
            "type": "execute",
            "stmt": {"sql": sql, "args": [_turso_arg(a) for a in args]},
        })
    reqs.append({"type": "close"})
    body = json.dumps({"requests": reqs}).encode("utf-8")
    req = urllib.request.Request(
        endpoint,
        data=body,
        headers={
            "Authorization": "Bearer " + TURSO_AUTH_TOKEN,
            "Content-Type": "application/json",
        },
        method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            payload = json.loads(resp.read().decode("utf-8"))
    except Exception as exc:
        raise TursoError("Turso request failed: %s" % exc) from exc
    results = []
    for res in payload.get("results", []):
        if res.get("type") == "error":
            err = res.get("error") or {}
            raise TursoError(err.get("message") or "Turso SQL error")
        response = res.get("response") or {}
        if response.get("type") == "execute":
            results.append(response.get("result") or {})
    return results


def _ensure_turso_schema():
    global _DB_READY
    if _DB_READY:
        return
    _turso_pipeline([
        """CREATE TABLE IF NOT EXISTS registrations (
            student_key TEXT PRIMARY KEY,
            student_id TEXT NOT NULL,
            submitted_at TEXT NOT NULL,
            name TEXT NOT NULL,
            phone TEXT NOT NULL,
            email TEXT NOT NULL,
            batch TEXT NOT NULL,
            major TEXT NOT NULL,
            minor TEXT NOT NULL,
            specialisation TEXT NOT NULL,
            times INTEGER NOT NULL DEFAULT 1,
            photocode TEXT NOT NULL DEFAULT '',
            photocode_at TEXT NOT NULL DEFAULT ''
        )""",
        """CREATE UNIQUE INDEX IF NOT EXISTS idx_photocode_nocase
           ON registrations(LOWER(photocode)) WHERE photocode != ''""",
        """CREATE TABLE IF NOT EXISTS event_log (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            kind TEXT NOT NULL,
            student_id TEXT NOT NULL,
            payload TEXT NOT NULL,
            created_at TEXT NOT NULL
        )""",
    ])
    _DB_READY = True


def append_record(rec):
    with open(DATA_FILE, "a", encoding="utf-8") as f:
        f.write(json.dumps(rec, ensure_ascii=False) + "\n")
        f.flush()
        os.fsync(f.fileno())


def _load_rows_local():
    """Current state per Student ID, oldest first. Replays the local JSONL log in order."""
    rows = {}
    if not os.path.exists(DATA_FILE):
        return []
    with open(DATA_FILE, encoding="utf-8") as f:
        for line in f:
            line = line.strip()
            if not line:
                continue
            try:
                rec = json.loads(line)
            except ValueError:
                continue
            if not isinstance(rec, dict):
                continue
            k = key(rec.get("studentId", ""))
            kind = rec.get("t")
            if kind == "del":
                rows.pop(k, None)
            elif kind == "add":
                prev = rows.get(k)
                rows[k] = {
                    "submittedAt": str(rec.get("submittedAt", "")),
                    "name": str(rec.get("name", "")), "phone": str(rec.get("phone", "")),
                    "email": str(rec.get("email", "")), "batch": str(rec.get("batch", "")),
                    "major": str(rec.get("major", "")), "minor": str(rec.get("minor", "")),
                    "specialisation": str(rec.get("specialisation", "N/A")),
                    "studentId": str(rec.get("studentId", "")),
                    "times": (prev["times"] + 1) if prev else 1,
                    "photocode": prev["photocode"] if prev else "",
                    "photocodeAt": prev["photocodeAt"] if prev else "",
                }
            elif kind == "photo" and k in rows:
                rows[k]["photocode"] = str(rec.get("photocode", ""))
                rows[k]["photocodeAt"] = str(rec.get("at", ""))
    return sorted(rows.values(), key=lambda r: r["submittedAt"])


def load_rows():
    """Current state per Student ID, oldest first (Turso DB or local JSONL)."""
    if not USE_TURSO:
        return _load_rows_local()
    _ensure_turso_schema()
    res = _turso_pipeline([(
        "SELECT submitted_at, name, phone, email, batch, major, minor, "
        "specialisation, student_id, times, photocode, photocode_at "
        "FROM registrations ORDER BY submitted_at ASC",
        [],
    )])
    parsed = _turso_parse_rows(res[0])
    return [{
        "submittedAt": str(r.get("submitted_at", "")),
        "name": str(r.get("name", "")),
        "phone": str(r.get("phone", "")),
        "email": str(r.get("email", "")),
        "batch": str(r.get("batch", "")),
        "major": str(r.get("major", "")),
        "minor": str(r.get("minor", "")),
        "specialisation": str(r.get("specialisation", "N/A")),
        "studentId": str(r.get("student_id", "")),
        "times": int(r.get("times") or 1),
        "photocode": str(r.get("photocode") or ""),
        "photocodeAt": str(r.get("photocode_at") or ""),
    } for r in parsed]


def save_submission(rec):
    """Save or update a registration. Returns (status, replaced_bool)."""
    sid_key = key(rec["studentId"])
    if not USE_TURSO:
        rows = _load_rows_local()
        existing = next((r for r in rows if key(r["studentId"]) == sid_key), None)
        if existing and existing["photocode"]:
            return "locked", False
        append_record(rec)
        write_xlsx_copy(_load_rows_local())
        return "ok", existing is not None

    _ensure_turso_schema()
    res = _turso_pipeline([(
        "SELECT photocode FROM registrations WHERE student_key = ?",
        [sid_key],
    )])
    found = _turso_parse_rows(res[0])
    if found and found[0].get("photocode"):
        return "locked", False
    replaced = bool(found)
    _turso_pipeline([
        (
            """INSERT INTO registrations (
                student_key, student_id, submitted_at, name, phone, email,
                batch, major, minor, specialisation, times, photocode, photocode_at
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1, '', '')
            ON CONFLICT(student_key) DO UPDATE SET
                student_id = excluded.student_id,
                submitted_at = excluded.submitted_at,
                name = excluded.name,
                phone = excluded.phone,
                email = excluded.email,
                batch = excluded.batch,
                major = excluded.major,
                minor = excluded.minor,
                specialisation = excluded.specialisation,
                times = registrations.times + 1""",
            [
                sid_key, rec["studentId"], rec["submittedAt"], rec["name"],
                rec["phone"], rec["email"], rec["batch"], rec["major"],
                rec["minor"], rec["specialisation"],
            ],
        ),
        (
            "INSERT INTO event_log (kind, student_id, payload, created_at) VALUES (?, ?, ?, ?)",
            ["add", rec["studentId"], json.dumps(rec, ensure_ascii=False), rec["submittedAt"]],
        ),
    ])
    return "ok", replaced


def save_photocode(sid_key, code, at):
    """Assign a photocode to a student. Returns (status, info_dict)."""
    if not USE_TURSO:
        rows = _load_rows_local()
        row = next((r for r in rows if key(r["studentId"]) == sid_key), None)
        if row is None:
            return "missing", None
        for other in rows:
            if other is not row and other["photocode"] and other["photocode"].casefold() == code.casefold():
                return "dup", other
        if row["photocode"] != code:
            append_record({"t": "photo", "studentId": row["studentId"], "photocode": code, "at": at})
            write_xlsx_copy(_load_rows_local())
        return "ok", {"name": row["name"], "studentId": row["studentId"], "photocode": code}

    _ensure_turso_schema()
    res = _turso_pipeline([
        ("SELECT student_id, name, photocode FROM registrations WHERE student_key = ?", [sid_key]),
        ("SELECT student_key, student_id, name FROM registrations WHERE LOWER(photocode) = LOWER(?) AND student_key != ?", [code, sid_key]),
    ])
    target_rows = _turso_parse_rows(res[0])
    if not target_rows:
        return "missing", None
    dup_rows = _turso_parse_rows(res[1])
    if dup_rows:
        return "dup", {"name": dup_rows[0]["name"], "studentId": dup_rows[0]["student_id"]}
    row = target_rows[0]
    if str(row.get("photocode") or "") != code:
        try:
            _turso_pipeline([
                ("UPDATE registrations SET photocode = ?, photocode_at = ? WHERE student_key = ?", [code, at, sid_key]),
                ("INSERT INTO event_log (kind, student_id, payload, created_at) VALUES (?, ?, ?, ?)",
                 ["photo", row["student_id"], json.dumps({"photocode": code, "at": at}), at]),
            ])
        except TursoError as exc:
            if "UNIQUE" in str(exc).upper():
                return "dup", {"name": "another student", "studentId": ""}
            raise
    return "ok", {"name": row["name"], "studentId": row["student_id"], "photocode": code}


def delete_registration(raw_sid, at):
    sid_clean = clean(raw_sid, 40)
    sid_key = key(sid_clean)
    if not USE_TURSO:
        append_record({"t": "del", "studentId": sid_clean, "at": at})
        write_xlsx_copy(_load_rows_local())
        return
    _ensure_turso_schema()
    _turso_pipeline([
        ("DELETE FROM registrations WHERE student_key = ?", [sid_key]),
        ("INSERT INTO event_log (kind, student_id, payload, created_at) VALUES (?, ?, ?, ?)",
         ["del", sid_clean, json.dumps({"at": at}), at]),
    ])


def build_workbook(rows):
    wb = Workbook()
    ws = wb.active
    ws.title = "Registrations"
    ws.append(COLUMNS)
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="011F78")  # Logo Royal Blue
        c.alignment = Alignment(vertical="center")
    for i, r in enumerate(rows, 1):
        values = [i, r["submittedAt"], r["name"], r["phone"], r["email"], r["batch"] or "-",
                  r["major"], r["minor"], r["specialisation"], r["studentId"],
                  r["photocode"] or "Pending", r["photocodeAt"] or None]
        for col, v in enumerate(values, 1):
            cell = ws.cell(row=i + 1, column=col, value=v)
            if isinstance(v, str):
                cell.data_type = "s"  # always text, never a formula
                cell.number_format = "@"
    for col, w in enumerate([6, 20, 26, 16, 30, 9, 34, 34, 26, 16, 16, 20], 1):
        ws.column_dimensions[get_column_letter(col)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = "A1:%s%d" % (get_column_letter(len(COLUMNS)), max(len(rows) + 1, 2))
    return wb


def write_xlsx_copy(rows):
    """Best effort. If the file is open in Excel it is locked; the log is safe regardless."""
    tmp = XLSX_FILE + ".tmp"
    try:
        build_workbook(rows).save(tmp)
        os.replace(tmp, XLSX_FILE)
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------- routes
def render_form():
    def js(v):
        return json.dumps(v).replace("<", "\\u003c")
    return (FORM_HTML.replace("/*BASE_CSS*/", BASE_CSS)
            .replace("__LOGO__", LOGO_DATA_URI)
            .replace("__TABLE__", js(MAJOR_MINORS))
            .replace("__BATCHES__", js(BATCHES))
            .replace("__UG__", js(sorted(UG_BATCHES))))


@app.get("/")
def index():
    return Response(render_form(), mimetype="text/html")


@app.post("/api/submit")
def submit():
    d = request.get_json(silent=True)
    if not isinstance(d, dict):
        return jsonify(ok=False, errors={"form": "Invalid request."}), 400
    rec, errors = validate(d)
    if errors:
        return jsonify(ok=False, errors=errors), 422
    try:
        with LOCK:
            status, replaced = save_submission(rec)
            if status == "locked":
                return jsonify(ok=False, errors={"studentId": "This Student ID already has a confirmed photocode, "
                                                              "so it cannot be changed online. Ask the organiser."}), 409
    except (OSError, TursoError):
        return jsonify(ok=False, errors={"form": "The server could not save your entry. Tell the organiser."}), 500
    return jsonify(ok=True, replaced=replaced, savedAt=rec["submittedAt"],
                   name=rec["name"], studentId=rec["studentId"], batch=rec["batch"],
                   major=rec["major"], minor=rec["minor"], specialisation=rec["specialisation"])


# ---- admin: sign-in with a small lock-out after repeated wrong passwords
FAILS = {}
FAIL_LIMIT, FAIL_WINDOW = 8, 60


def _now():
    return time.time()


def _client_ip():
    ip = request.remote_addr or "?"
    if os.environ.get("VERCEL") or ip in ("127.0.0.1", "::1"):
        xff = request.headers.get("X-Forwarded-For", "").strip()
        if xff:
            return xff.split(",")[0].strip() or ip
    return ip


def locked_out(ip):
    rec = FAILS.get(ip)
    return bool(rec) and rec[0] >= FAIL_LIMIT and _now() - rec[1] < FAIL_WINDOW


def note_failure(ip):
    now = _now()
    if len(FAILS) > 1000:
        for k in [k for k, v in FAILS.items() if now - v[1] >= FAIL_WINDOW]:
            FAILS.pop(k, None)
    rec = FAILS.get(ip)
    if not rec or now - rec[1] >= FAIL_WINDOW:
        FAILS[ip] = [1, now]
    else:
        rec[0] += 1


def admin_required(fn):
    @wraps(fn)
    def wrapper(*a, **kw):
        ip = _client_ip()
        if locked_out(ip):
            return Response("Too many wrong passwords. Wait a minute and try again.", 429,
                            {"Retry-After": str(FAIL_WINDOW)})
        auth = request.authorization
        if auth and hmac.compare_digest((auth.password or "").encode(), ADMIN_PASSWORD.encode()):
            FAILS.pop(ip, None)
            return fn(*a, **kw)
        if auth:
            note_failure(ip)
        return Response("Admin login required.", 401, {"WWW-Authenticate": 'Basic realm="Years at FLAME Admin"'})
    return wrapper


@app.get("/admin")
@admin_required
def admin_page():
    return Response(ADMIN_HTML.replace("/*BASE_CSS*/", BASE_CSS).replace("__LOGO__", LOGO_DATA_URI),
                    mimetype="text/html")


@app.get("/admin/info")
@admin_required
def admin_info():
    info = resolve_student_urls()
    info["port"] = PORT
    info["dataDir"] = HERE
    info["storage"] = "turso" if USE_TURSO else "local"
    return jsonify(info)


@app.get("/admin/qr.svg")
@admin_required
def admin_qr():
    target = (request.args.get("url") or "").strip()
    if not target:
        target = resolve_student_urls()["primaryUrl"]
    if len(target) > 260 or not (target.startswith("http://") or target.startswith("https://")):
        return Response("Invalid URL for QR code.", 400, mimetype="text/plain")
    try:
        svg = render_qr_svg(target)
    except ValueError as exc:
        return Response(str(exc), 400, mimetype="text/plain")
    return Response(svg, mimetype="image/svg+xml")


@app.get("/admin/data")
@admin_required
def admin_data():
    with LOCK:
        rows = load_rows()
    return jsonify(rows=list(reversed(rows)))


@app.get("/admin/export.xlsx")
@admin_required
def admin_export():
    with LOCK:
        rows = load_rows()
    buf = io.BytesIO()
    build_workbook(rows).save(buf)
    buf.seek(0)
    name = "registrations-%s.xlsx" % datetime.now().strftime("%Y-%m-%d")
    return send_file(buf, as_attachment=True, download_name=name,
                     mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.post("/admin/photocode")
@admin_required
def admin_photocode():
    d = request.get_json(silent=True)
    if not isinstance(d, dict):
        return jsonify(ok=False, error="Invalid request."), 400
    sid = key(d.get("studentId"))
    code = clean(d.get("photocode"), 40)
    code2 = clean(d.get("photocode2"), 40)
    if not code or not PHOTO_RE.fullmatch(code):
        return jsonify(ok=False, error="That photocode is not in the expected format."), 422
    if code != code2:
        return jsonify(ok=False, error="The two entries do not match."), 422
    at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with LOCK:
            status, info = save_photocode(sid, code, at)
    except (OSError, TursoError):
        return jsonify(ok=False, error="Could not save the photocode."), 500
    if status == "missing":
        return jsonify(ok=False, error="That student is no longer in the list. Refresh the page."), 404
    if status == "dup":
        who = ("%s (%s)" % (info["name"], info["studentId"])) if info and info.get("studentId") else "another student"
        return jsonify(ok=False, error="This photocode is already assigned to %s." % who), 409
    return jsonify(ok=True, name=info["name"], studentId=info["studentId"], photocode=code)


@app.post("/admin/delete")
@admin_required
def admin_delete():
    d = request.get_json(silent=True)
    if not isinstance(d, dict) or not d.get("studentId"):
        return jsonify(ok=False), 400
    at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    try:
        with LOCK:
            delete_registration(d["studentId"], at)
    except (OSError, TursoError):
        return jsonify(ok=False), 500
    return jsonify(ok=True)


# ------------------------------------------------------------------ page parts
# Palette sampled directly from the Years at FLAME logo:
#   #F6F7F7 ground · #FFFFFF surface · #011F78 primary royal blue
#   #D9AC58 secondary ochre gold · #43687E slate blue
BASE_CSS = r"""
:root {
  --ground: #F6F7F7;
  --surface: #FFFFFF;
  --subtle: #F1F3F6;
  --ink: #111827;
  --muted: #4B5565;
  --line: #D8DDE6;
  --line-strong: #B8C2D4;
  --brand: #011F78;
  --brand-hover: #0A2B8F;
  --brand-press: #00175C;
  --brand-ink: #FFFFFF;
  --brand-tint: #EDF1FA;
  --gold: #D9AC58;
  --gold-border: #E2BE74;
  --gold-tint: #FDF7EB;
  --gold-ink: #78520E;
  --slate: #43687E;
  --good: #116B45;
  --good-tint: #E8F5EE;
  --bad: #B42318;
  --bad-tint: #FEF0EE;
  --r-sm: 4px;
  --r-md: 6px;
  --r-lg: 8px;
  --shadow: 0 1px 2px rgba(1, 31, 120, 0.05);
  --font-sans: -apple-system, BlinkMacSystemFont, "Helvetica Neue", "Segoe UI", Arial, sans-serif;
  --font-mono: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
  color-scheme: light;
}
*, *::before, *::after { box-sizing: border-box; }
[hidden] { display: none !important; }
html, body { margin: 0; padding: 0; }
body {
  background: var(--ground);
  color: var(--ink);
  font-family: var(--font-sans);
  font-size: 16px;
  line-height: 1.5;
  -webkit-font-smoothing: antialiased;
}

/* Full-width institutional header bar */
.topbar {
  background: var(--surface);
  border-top: 3px solid var(--brand);
  border-bottom: 1px solid var(--line);
}
.topbar-inner {
  max-width: 680px;
  margin: 0 auto;
  height: 64px;
  padding: 0 20px;
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
}
.topbar-inner.wide { max-width: 1220px; }
.brand {
  display: inline-flex;
  align-items: center;
  gap: 12px;
  min-width: 0;
}
.brand img {
  width: 40px;
  height: 40px;
  display: block;
  flex: none;
}
.brand-name {
  font-size: 15px;
  font-weight: 600;
  color: var(--brand);
  letter-spacing: -0.01em;
  white-space: nowrap;
}
.brand-sep {
  color: var(--line-strong);
  font-weight: 400;
}
.brand-sub {
  font-size: 14px;
  font-weight: 500;
  color: var(--muted);
  white-space: nowrap;
}
.step-meta {
  font-size: 13px;
  font-weight: 500;
  color: var(--muted);
  white-space: nowrap;
}

/* Containers */
.main {
  max-width: 680px;
  margin: 0 auto;
  padding: 32px 20px 64px;
  display: grid;
  gap: 24px;
}
.main.wide {
  max-width: 1220px;
  padding: 28px 20px 64px;
  gap: 20px;
}

.intro { display: grid; gap: 6px; }
h1 {
  margin: 0;
  font-size: 28px;
  font-weight: 600;
  line-height: 1.18;
  letter-spacing: -0.02em;
  color: var(--ink);
}
h2 {
  margin: 0;
  font-size: 20px;
  font-weight: 600;
  line-height: 1.25;
  letter-spacing: -0.015em;
  color: var(--ink);
  outline: none;
}
.lede {
  margin: 0;
  font-size: 15px;
  color: var(--muted);
  max-width: 60ch;
}

/* 5-step underline rail */
.steps {
  list-style: none;
  margin: 0;
  padding: 0;
  display: grid;
  grid-template-columns: repeat(5, 1fr);
  gap: 10px;
}
.steps li {
  font-size: 13px;
  font-weight: 500;
  color: var(--muted);
  padding-bottom: 8px;
  border-bottom: 2px solid var(--line);
  white-space: nowrap;
  overflow: hidden;
  text-overflow: ellipsis;
  transition: color 0.15s, border-color 0.15s;
}
.steps li[aria-current="step"] {
  color: var(--brand);
  font-weight: 600;
  border-bottom-color: var(--brand);
}
.steps li.done {
  color: var(--ink);
  border-bottom-color: var(--good);
  cursor: pointer;
}
.steps li.done:hover { color: var(--brand); }

/* Primary surface sheet */
.card {
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r-lg);
  padding: 28px;
  display: grid;
  gap: 20px;
  min-width: 0;
  box-shadow: var(--shadow);
}
.card.accent-top { border-top: 3px solid var(--brand); }

.panel { display: grid; gap: 16px; }
.field { display: grid; gap: 6px; min-width: 0; position: relative; }
label, legend {
  font-size: 13px;
  font-weight: 500;
  color: var(--ink);
}
.opt { color: var(--muted); font-weight: 400; }

input[type=text], input[type=tel], input[type=email], input[type=search] {
  font: inherit;
  font-size: 16px;
  color: var(--ink);
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  padding: 0 12px;
  height: 46px;
  width: 100%;
  min-width: 0;
  transition: border-color 0.15s, box-shadow 0.15s;
}
input:hover:not(:disabled) { border-color: var(--line-strong); }
input::placeholder { color: #7B8698; }
input:disabled { background: var(--subtle); color: var(--muted); cursor: not-allowed; }
input:focus-visible {
  outline: none;
  border-color: var(--brand);
  box-shadow: 0 0 0 3px var(--brand-tint);
}
input[aria-invalid="true"] {
  border-color: var(--bad);
  box-shadow: 0 0 0 3px var(--bad-tint);
}

.mono { font-family: var(--font-mono); letter-spacing: 0.01em; }
.err { color: var(--bad); font-size: 13px; margin: 0; font-weight: 500; }
.hint { font-size: 13px; color: var(--muted); margin: 0; }
.hint.good { color: var(--good); font-weight: 500; }
.hint.bad { color: var(--bad); font-weight: 500; }

fieldset.verify-box {
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  padding: 16px;
  margin: 0;
  min-width: 0;
  display: grid;
  gap: 14px;
  background: var(--ground);
}
fieldset.plain { border: 0; padding: 0; margin: 0; min-width: 0; display: grid; gap: 12px; }
legend { padding: 0 6px; }
fieldset.plain legend { padding: 0; }

/* Batch selector grid */
.rgrid { display: grid; grid-template-columns: repeat(auto-fit, minmax(112px, 1fr)); gap: 10px; }
.radio { position: relative; display: block; cursor: pointer; }
.radio input { position: absolute; inset: 0; width: 100%; height: 100%; margin: 0; opacity: 0; cursor: pointer; }
.radio .tile {
  display: grid;
  gap: 2px;
  padding: 14px 12px;
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  background: var(--surface);
  text-align: left;
  transition: border-color 0.15s, background-color 0.15s, transform 0.12s;
}
.radio:hover .tile { border-color: var(--line-strong); transform: translateY(-1px); }
.radio .tile strong { font-size: 15px; font-weight: 600; color: var(--ink); }
.radio .tile small { font-size: 13px; color: var(--muted); font-weight: 400; }
.radio input:checked + .tile {
  background: var(--brand-tint);
  border-color: var(--brand);
  box-shadow: 0 0 0 1px var(--brand) inset;
}
.radio input:checked + .tile strong { color: var(--brand); }
.radio input:focus-visible + .tile { outline: 2px solid var(--brand); outline-offset: 2px; }

/* Searchable combobox */
.combo { position: relative; }
.combo input { padding-right: 38px; }
.combo-clear {
  position: absolute;
  right: 8px;
  top: 50%;
  transform: translateY(-50%);
  width: 26px;
  height: 26px;
  border-radius: var(--r-sm);
  border: 0;
  background: var(--subtle);
  color: var(--muted);
  font-size: 16px;
  line-height: 1;
  cursor: pointer;
  display: grid;
  place-items: center;
}
.combo-clear:hover { background: var(--line); color: var(--ink); }
.combo-menu {
  position: absolute;
  left: 0;
  right: 0;
  top: calc(100% + 4px);
  z-index: 40;
  max-height: 232px;
  overflow-y: auto;
  background: var(--surface);
  border: 1px solid var(--line-strong);
  border-radius: var(--r-md);
  padding: 4px;
  box-shadow: 0 10px 28px rgba(17, 24, 39, 0.12);
  display: grid;
  gap: 2px;
}
.combo-opt {
  text-align: left;
  width: 100%;
  border: 0;
  background: transparent;
  color: var(--ink);
  font: inherit;
  font-size: 14px;
  padding: 8px 10px;
  border-radius: var(--r-sm);
  cursor: pointer;
}
.combo-opt:hover, .combo-opt[aria-selected="true"] {
  background: var(--brand-tint);
  color: var(--brand);
  font-weight: 500;
}
.combo-empty { padding: 10px; font-size: 13px; color: var(--muted); margin: 0; }

/* Buttons & links */
.actions { display: flex; gap: 10px; justify-content: flex-end; flex-wrap: wrap; }
.btn {
  font: inherit;
  font-size: 15px;
  font-weight: 600;
  height: 44px;
  padding: 0 20px;
  border-radius: var(--r-md);
  border: 1px solid transparent;
  cursor: pointer;
  text-decoration: none;
  white-space: nowrap;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  gap: 8px;
  transition: background-color 0.15s, border-color 0.15s, transform 0.1s;
}
.btn:active:not(:disabled) { transform: translateY(1px); }
.btn.primary { background: var(--brand); color: var(--brand-ink); }
.btn.primary:hover:not(:disabled) { background: var(--brand-hover); }
.btn.primary:active:not(:disabled) { background: var(--brand-press); }
.btn.secondary { background: var(--surface); border-color: var(--line-strong); color: var(--ink); }
.btn.secondary:hover:not(:disabled) { background: var(--subtle); border-color: #9AA6BC; }
.btn.small { height: 38px; padding: 0 14px; font-size: 14px; }
.btn:disabled { opacity: 0.55; cursor: not-allowed; }
.btn:focus-visible, .link:focus-visible { outline: 2px solid var(--brand); outline-offset: 2px; }

.link {
  background: none;
  border: 0;
  color: var(--brand);
  font: inherit;
  font-size: 14px;
  font-weight: 500;
  cursor: pointer;
  padding: 2px 0;
  text-decoration: underline;
  text-underline-offset: 3px;
  white-space: nowrap;
}
.link:hover { color: var(--brand-hover); }
.link.bad { color: var(--bad); }

.status { margin: 0; font-size: 14px; color: var(--muted); min-height: 1.3em; }
.status.bad { color: var(--bad); font-weight: 500; }

/* Review list using horizontal rules (no box-in-box clutter) */
.proof-row {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  flex-wrap: wrap;
  padding: 14px 16px;
  border-radius: var(--r-md);
  background: var(--brand-tint);
  border-left: 3px solid var(--brand);
}
.proof-name { font-size: 17px; font-weight: 600; color: var(--ink); margin: 0; }
.proof-meta { font-size: 14px; color: var(--muted); margin: 2px 0 0; }
.proof-id {
  font-family: var(--font-mono);
  font-size: 14px;
  font-weight: 600;
  color: var(--brand);
}

.rlist { display: grid; }
.rgroup {
  border-top: 1px solid var(--line);
  padding: 14px 0;
  display: grid;
  gap: 8px;
}
.rgroup:first-child { border-top: 0; padding-top: 4px; }
.rhead { display: flex; justify-content: space-between; align-items: center; gap: 12px; }
.rhead h3 { margin: 0; font-size: 13px; font-weight: 600; color: var(--muted); }
dl { margin: 0; display: grid; grid-template-columns: minmax(110px, max-content) 1fr; gap: 6px 16px; }
dt { color: var(--muted); font-size: 14px; }
dd { margin: 0; font-weight: 500; font-size: 15px; overflow-wrap: anywhere; min-width: 0; color: var(--ink); }

/* Done confirmation */
.done-head { display: flex; gap: 12px; align-items: center; }
.tick {
  width: 36px;
  height: 36px;
  border-radius: 50%;
  background: var(--good-tint);
  display: grid;
  place-items: center;
  flex: none;
}
.tick svg { width: 18px; height: 18px; stroke: var(--good); fill: none; stroke-width: 2.5; stroke-linecap: round; stroke-linejoin: round; }

/* Admin unified metrics strip */
.metrics {
  background: var(--surface);
  border: 1px solid var(--line);
  border-radius: var(--r-lg);
  display: grid;
  grid-template-columns: repeat(4, minmax(0, 1fr));
  box-shadow: var(--shadow);
  overflow: hidden;
}
.metric {
  padding: 16px 20px;
  border-right: 1px solid var(--line);
  display: grid;
  gap: 4px;
  align-content: start;
}
.metric:last-child { border-right: 0; }
.metric-label { font-size: 13px; font-weight: 500; color: var(--muted); }
.metric-val { font-size: 26px; font-weight: 600; line-height: 1.15; letter-spacing: -0.02em; font-variant-numeric: tabular-nums; color: var(--ink); }
.metric-sub { font-size: 13px; color: var(--muted); margin: 0; }
.bar-track { height: 4px; border-radius: 2px; background: var(--line); overflow: hidden; margin-top: 6px; }
.bar-fill { height: 100%; background: var(--brand); width: 0%; transition: width 0.25s ease; }

/* Admin toolbar */
.toolbar {
  display: flex;
  flex-wrap: wrap;
  gap: 10px;
  align-items: center;
  justify-content: space-between;
}
.searchbox { flex: 1 1 260px; max-width: 380px; }
.searchbox input { height: 38px; font-size: 14px; }
.seg-group { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; }
.seg {
  height: 36px;
  padding: 0 12px;
  border-radius: var(--r-md);
  border: 1px solid var(--line);
  background: var(--surface);
  color: var(--muted);
  font: inherit;
  font-size: 13px;
  font-weight: 500;
  cursor: pointer;
  white-space: nowrap;
  transition: border-color 0.15s, color 0.15s, background-color 0.15s;
}
.seg:hover { border-color: var(--line-strong); color: var(--ink); }
.seg.active { background: var(--brand-tint); border-color: var(--brand); color: var(--brand); font-weight: 600; }

/* Admin table */
.tablewrap {
  overflow-x: auto;
  border: 1px solid var(--line);
  border-radius: var(--r-lg);
  background: var(--surface);
  box-shadow: var(--shadow);
}
table { border-collapse: collapse; width: 100%; min-width: 1200px; font-size: 14px; font-variant-numeric: tabular-nums; }
th, td { text-align: left; padding: 11px 14px; border-bottom: 1px solid var(--line); white-space: nowrap; }
th {
  font-size: 13px;
  font-weight: 600;
  color: var(--muted);
  background: var(--ground);
  position: sticky;
  top: 0;
  z-index: 2;
}
tbody tr:hover td { background: #F8FAFC; }
tbody tr:last-child td { border-bottom: 0; }
td.mono { font-family: var(--font-mono); font-size: 13px; }
td.empty { padding: 36px 16px; text-align: center; color: var(--muted); white-space: normal; }

.tag {
  display: inline-block;
  font-size: 12px;
  font-weight: 600;
  padding: 2px 8px;
  border-radius: var(--r-sm);
  background: var(--brand-tint);
  color: var(--brand);
}
.tag.good { background: var(--good-tint); color: var(--good); }
.tag.warn { background: var(--gold-tint); color: var(--gold-ink); border: 1px solid var(--gold-border); }
.pcell { display: flex; gap: 10px; align-items: center; }

dialog {
  border: 1px solid var(--line-strong);
  border-radius: var(--r-lg);
  background: var(--surface);
  color: var(--ink);
  padding: 0;
  width: min(500px, calc(100vw - 28px));
  font-family: var(--font-sans);
  box-shadow: 0 20px 50px rgba(17, 24, 39, 0.18);
}
dialog::backdrop { background: rgba(17, 24, 39, 0.45); }
.dlg { padding: 24px; display: grid; gap: 16px; }
.readback {
  margin: 0;
  font-size: 14px;
  padding: 10px 12px;
  border-radius: var(--r-md);
  background: var(--brand-tint);
  color: var(--ink);
  min-height: 2.5em;
  overflow-wrap: anywhere;
}

.qr-standee {
  background: var(--surface);
  border: 1px solid var(--line);
  border-top: 4px solid var(--brand);
  border-radius: var(--r-lg);
  padding: 22px;
  display: grid;
  gap: 14px;
  justify-items: center;
  text-align: center;
}
.qr-frame {
  width: 210px;
  height: 210px;
  padding: 8px;
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  background: var(--surface);
}
.qr-frame img { width: 100%; height: 100%; display: block; }
.qr-url {
  font-family: var(--font-mono);
  font-size: 13px;
  font-weight: 600;
  color: var(--brand);
  background: var(--brand-tint);
  padding: 7px 12px;
  border-radius: var(--r-md);
  word-break: break-all;
}
.alt-links {
  display: grid;
  gap: 4px;
  font-size: 13px;
  color: var(--muted);
  background: var(--ground);
  border: 1px solid var(--line);
  border-radius: var(--r-md);
  padding: 10px 12px;
}

@media (max-width: 900px) {
  .metrics { grid-template-columns: repeat(2, minmax(0, 1fr)); }
  .metric:nth-child(2) { border-right: 0; }
  .metric:nth-child(1), .metric:nth-child(2) { border-bottom: 1px solid var(--line); }
}
@media (max-width: 560px) {
  .topbar-inner { padding: 0 16px; }
  .main { padding: 20px 16px 56px; }
  .card { padding: 20px 16px; }
  h1 { font-size: 24px; }
  .steps { gap: 6px; }
  .steps li { font-size: 12px; }
  .actions .btn { flex: 1 1 auto; }
  .metrics { grid-template-columns: 1fr; }
  .metric { border-right: 0; border-bottom: 1px solid var(--line); }
  .metric:last-child { border-bottom: 0; }
}
@media (prefers-reduced-motion: reduce) { * { transition: none !important; animation: none !important; } }

@media print {
  body { background: #fff !important; }
  .topbar, .main, dialog:not([open]), .no-print { display: none !important; }
  dialog[open] { position: static; border: 0; box-shadow: none; width: 100%; max-width: 400px; margin: 24px auto; }
  .qr-standee { border: 1px solid var(--brand); border-top: 4px solid var(--brand); }
}
"""

FORM_HTML = r"""<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Years at FLAME — Photoshoot registration</title>
<style>/*BASE_CSS*/</style>
</head><body>
<header class="topbar">
  <div class="topbar-inner">
    <div class="brand">
      <img src="__LOGO__" alt="Years at FLAME">
      <span class="brand-name">Years at FLAME</span>
    </div>
    <span class="step-meta" id="stepcount">Step 1 of 5</span>
  </div>
</header>

<main class="main">
  <div class="intro">
    <h1>Photoshoot registration</h1>
    <p class="lede">Enter your details for the yearbook photoshoot. You will review everything before saving.</p>
  </div>

  <ol class="steps" id="steps" aria-label="Progress">
    <li>1. Details</li><li>2. Batch</li><li>3. Studies</li><li>4. Student ID</li><li>5. Review</li>
  </ol>

  <form class="card" id="reg" novalidate>
    <section class="panel" data-step="1" aria-labelledby="h-1">
      <h2 id="h-1" tabindex="-1">Your details</h2>
      <div class="field"><label for="f-name">Full name <span class="opt">(for the yearbook)</span></label>
        <input id="f-name" type="text" autocomplete="name" placeholder="Aarav Mehta" aria-describedby="e-name"><p class="err" id="e-name" hidden></p></div>
      <div class="field"><label for="f-phone">Mobile number</label>
        <input id="f-phone" type="tel" inputmode="tel" autocomplete="tel" placeholder="+91 98765 43210" aria-describedby="e-phone"><p class="err" id="e-phone" hidden></p></div>
      <div class="field"><label for="f-email">Email address</label>
        <input id="f-email" type="email" inputmode="email" autocomplete="email" autocapitalize="off" spellcheck="false" placeholder="name@flame.edu.in" aria-describedby="e-email"><p class="err" id="e-email" hidden></p></div>
    </section>

    <section class="panel" data-step="2" aria-labelledby="h-2" hidden>
      <h2 id="h-2" tabindex="-1">Your batch</h2>
      <fieldset class="plain" id="f-batch" aria-describedby="e-batch">
        <legend>Select your cohort</legend>
        <div class="rgrid" id="batchGrid"></div>
        <p class="err" id="e-batch" hidden></p>
      </fieldset>
    </section>

    <section class="panel" data-step="3" aria-labelledby="h-3" hidden>
      <h2 id="h-3" tabindex="-1">What you study</h2>
      <div class="panel" id="ugFields" hidden>
        <div class="field"><label for="f-major">Major</label>
          <div class="combo" id="combo-major">
            <input id="f-major" type="text" role="combobox" aria-autocomplete="list" aria-expanded="false" aria-controls="menu-major" autocomplete="off" placeholder="Search or choose a major" aria-describedby="h-major e-major">
            <button type="button" class="combo-clear" id="clear-major" aria-label="Clear major" hidden>&times;</button>
            <div class="combo-menu" id="menu-major" role="listbox" hidden></div>
          </div>
          <p class="hint" id="h-major">Pick your major from the list.</p><p class="err" id="e-major" hidden></p></div>
        <div class="field"><label for="f-minor">Minor <span class="opt">(optional)</span></label>
          <div class="combo" id="combo-minor">
            <input id="f-minor" type="text" role="combobox" aria-autocomplete="list" aria-expanded="false" aria-controls="menu-minor" autocomplete="off" placeholder="Search or choose a minor" aria-describedby="minorHint e-minor" disabled>
            <button type="button" class="combo-clear" id="clear-minor" aria-label="Clear minor" hidden>&times;</button>
            <div class="combo-menu" id="menu-minor" role="listbox" hidden></div>
          </div>
          <p class="hint" id="minorHint">Choose your major first.</p><p class="err" id="e-minor" hidden></p></div>
      </div>
      <div class="panel" id="pgFields" hidden>
        <div class="field"><label for="f-specialisation">Specialisation</label>
          <input id="f-specialisation" type="text" autocomplete="off" placeholder="Enter your specialisation" aria-describedby="e-specialisation">
          <p class="err" id="e-specialisation" hidden></p></div>
      </div>
    </section>

    <section class="panel" data-step="4" aria-labelledby="h-4" hidden>
      <h2 id="h-4" tabindex="-1">Confirm your Student ID</h2>
      <p class="hint">Type it twice. Pasting is switched off in the second box so the repeat is a real check.</p>
      <fieldset class="verify-box"><legend>Student ID</legend>
        <div class="field"><label for="f-studentId">Student ID</label>
          <input id="f-studentId" class="mono" type="text" autocomplete="off" autocapitalize="off" spellcheck="false" aria-describedby="e-studentId"><p class="err" id="e-studentId" hidden></p></div>
        <div class="field"><label for="f-studentId2">Re-enter Student ID</label>
          <input id="f-studentId2" class="mono" type="text" autocomplete="off" autocapitalize="off" spellcheck="false" aria-describedby="e-studentId2 m-studentId2">
          <p class="hint" id="m-studentId2" aria-live="polite"></p><p class="err" id="e-studentId2" hidden></p></div>
      </fieldset>
    </section>

    <section class="panel" data-step="5" aria-labelledby="h-5" hidden>
      <h2 id="h-5" tabindex="-1">Check and submit</h2>
      <p class="hint">Read every line. Select Edit to change anything before you submit.</p>
      <div class="proof-row" id="proofBanner"></div>
      <div class="rlist" id="review"></div>
    </section>

    <div class="actions">
      <button type="button" class="btn secondary" id="back" hidden>Back</button>
      <button type="submit" class="btn primary" id="next">Continue</button>
    </div>
    <p class="status" id="status" role="status"></p>
  </form>

  <section class="card accent-top" id="done" hidden aria-labelledby="h-done">
    <div class="done-head">
      <span class="tick" aria-hidden="true"><svg viewBox="0 0 24 24"><path d="M5 12.5l4.5 4.5L19 7.5"/></svg></span>
      <h2 id="h-done" tabindex="-1">Registration saved</h2>
    </div>
    <dl id="doneSummary"></dl>
    <p id="doneMsg" class="hint" style="font-size:14px"></p>
    <div class="actions"><button type="button" class="btn primary" id="again">Register another student</button></div>
  </section>
</main>

<script>
(function () {
  "use strict";
  var TABLE = __TABLE__;
  var BATCHES = __BATCHES__;
  var UG = __UG__;
  var MAJORS = Object.keys(TABLE);
  var BATCH_SUB = { UG3: "3rd year UG", UG4: "4th year UG", MSC: "MSc", PG2: "PG year 2", PGPEI: "PGPEI" };
  var TOTAL = 5;
  var byId = function (id) { return document.getElementById(id); };
  function el(tag, cls) { var e = document.createElement(tag); if (cls) e.className = cls; return e; }
  function collapse(s) { return String(s || "").replace(/\s+/g, " ").trim(); }
  function pick(v, list) {
    var k = collapse(v).toLowerCase();
    if (!k) return "";
    for (var i = 0; i < list.length; i++) if (list[i].toLowerCase() === k) return list[i];
    return "";
  }
  function val(k) { var n = byId("f-" + k); return n && typeof n.value === "string" ? n.value.trim() : ""; }
  function getBatch() { var r = document.querySelector('input[name="batch"]:checked'); return r ? r.value : ""; }
  function isUG() { return UG.indexOf(getBatch()) !== -1; }
  function curMajor() { return pick(val("major"), MAJORS); }
  function allowed() { var m = curMajor(); return m ? TABLE[m] : []; }

  var step = 1, busy = false, lastMajor = "";
  var stepOf = { name: 1, phone: 1, email: 1, batch: 2, major: 3, minor: 3, specialisation: 3, studentId: 4, studentId2: 4 };
  var stepFields = { 1: ["name", "phone", "email"], 2: ["batch"], 3: ["major", "minor", "specialisation"], 4: ["studentId", "studentId2"] };

  /* ---- batch ---- */
  BATCHES.forEach(function (b) {
    var lab = el("label", "radio"), inp = el("input"), tile = el("span", "tile"), st = el("strong"), sm = el("small");
    inp.type = "radio"; inp.name = "batch"; inp.value = b;
    st.textContent = b; sm.textContent = BATCH_SUB[b] || "Cohort";
    tile.append(st, sm);
    inp.addEventListener("change", function () { setErr("batch", ""); applyBatch(); });
    lab.append(inp, tile); byId("batchGrid").append(lab);
  });
  function applyBatch() {
    var b = getBatch(), ug = isUG();
    byId("ugFields").hidden = !(b && ug);
    byId("pgFields").hidden = !(b && !ug);
    byId("h-3").textContent = ug ? "Your major and minor" : "Your specialisation";
  }

  /* ---- searchable combobox for Major & Minor ---- */
  function setupCombo(inputId, menuId, clearId, getItems, onPick) {
    var inp = byId(inputId), menu = byId(menuId), clr = byId(clearId), activeIdx = -1, open = false;
    function syncClear() { clr.hidden = inp.disabled || !inp.value; }
    function close() { open = false; activeIdx = -1; menu.hidden = true; inp.setAttribute("aria-expanded", "false"); }
    function renderMenu() {
      if (inp.disabled) { close(); return; }
      var q = collapse(inp.value).toLowerCase(), all = getItems();
      var matches = q ? all.filter(function (x) { return x.toLowerCase().indexOf(q) !== -1; }) : all.slice();
      menu.textContent = "";
      if (!matches.length) {
        var empty = el("p", "combo-empty"); empty.textContent = "No matching option"; menu.append(empty);
      } else {
        matches.forEach(function (item, idx) {
          var opt = el("button", "combo-opt");
          opt.type = "button"; opt.setAttribute("role", "option");
          opt.setAttribute("aria-selected", idx === activeIdx ? "true" : "false");
          opt.textContent = item;
          opt.addEventListener("mousedown", function (e) {
            e.preventDefault();
            inp.value = item;
            syncClear();
            close();
            onPick(item);
          });
          menu.append(opt);
        });
      }
      menu.hidden = false; open = true; inp.setAttribute("aria-expanded", "true");
    }
    inp.addEventListener("focus", function () { activeIdx = -1; renderMenu(); });
    inp.addEventListener("click", function () { if (!open) renderMenu(); });
    inp.addEventListener("input", function () { activeIdx = -1; syncClear(); renderMenu(); });
    inp.addEventListener("blur", function () { setTimeout(close, 120); });
    inp.addEventListener("keydown", function (e) {
      var opts = menu.querySelectorAll(".combo-opt");
      if (e.key === "ArrowDown") { e.preventDefault(); if (!open) renderMenu(); else if (opts.length) { activeIdx = (activeIdx + 1) % opts.length; renderMenu(); } }
      else if (e.key === "ArrowUp") { e.preventDefault(); if (open && opts.length) { activeIdx = (activeIdx - 1 + opts.length) % opts.length; renderMenu(); } }
      else if (e.key === "Enter" && open && activeIdx >= 0 && opts[activeIdx]) {
        e.preventDefault();
        var chosen = opts[activeIdx].textContent;
        inp.value = chosen; syncClear(); close(); onPick(chosen);
      } else if (e.key === "Escape" && open) { e.stopPropagation(); close(); }
    });
    clr.addEventListener("click", function () { inp.value = ""; syncClear(); inp.focus(); onPick(""); renderMenu(); });
    return { syncClear: syncClear, close: close };
  }

  /* ---- major and minor, driven by the table ---- */
  var minorCombo = null;
  function syncMinor() {
    var m = curMajor(), minor = byId("f-minor"), hint = byId("minorHint"), list = m ? TABLE[m] : [];
    if (!m) {
      minor.disabled = true; hint.textContent = "Choose your major first.";
      if (minorCombo) { minorCombo.syncClear(); minorCombo.close(); }
      return;
    }
    if (!list.length) {
      minor.value = "N/A"; minor.disabled = true;
      hint.textContent = m + " does not offer a minor, so this is set to N/A.";
      setErr("minor", ""); lastMajor = m;
      if (minorCombo) { minorCombo.syncClear(); minorCombo.close(); }
      return;
    }
    if (minor.disabled && minor.value === "N/A") minor.value = "";
    minor.disabled = false;
    var note = "";
    if (m !== lastMajor && collapse(minor.value) && !pick(minor.value, list)) {
      minor.value = "";
      note = "Your minor was cleared because it is not offered with " + m + ". ";
    }
    hint.textContent = note + "Optional. " + list.length + " minors are offered with " + m + ". Leave blank if you have none.";
    lastMajor = m;
    if (minorCombo) minorCombo.syncClear();
  }

  var majorInput = byId("f-major"), minorInput = byId("f-minor");
  var majorCombo = setupCombo("f-major", "menu-major", "clear-major", function () { return MAJORS; }, function () {
    setErr("major", ""); syncMinor();
  });
  minorCombo = setupCombo("f-minor", "menu-minor", "clear-minor", function () { return allowed(); }, function () {
    setErr("minor", "");
  });
  majorInput.addEventListener("input", function () { setErr("major", ""); syncMinor(); });
  majorInput.addEventListener("change", function () { var m = curMajor(); if (m) majorInput.value = m; syncMinor(); majorCombo.syncClear(); });
  minorInput.addEventListener("input", function () { setErr("minor", ""); });
  minorInput.addEventListener("change", function () { var p = pick(minorInput.value, allowed()); if (p) minorInput.value = p; minorCombo.syncClear(); });
  byId("f-specialisation").addEventListener("input", function () { setErr("specialisation", ""); });

  /* ---- validation ---- */
  var PROG_RE = /^[\p{L}\p{N}_ &,.'()\/+\-]{2,80}$/u;
  function progBad(v) { return !PROG_RE.test(v) || /^[=+\-@]/.test(v); }
  var rules = {
    name: function (v) { return v.length >= 2 && /\p{L}/u.test(v) ? "" : "Enter your full name."; },
    phone: function (v) { return /^\+?\d{7,15}$/.test(v.replace(/[\s\-().]/g, "")) ? "" : "Enter a mobile number with 7 to 15 digits. A leading + is fine."; },
    email: function (v) { return /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(v) ? "" : "Enter an email address like name@flame.edu.in."; },
    batch: function () { return getBatch() ? "" : "Choose your batch."; },
    major: function (v) { if (!isUG()) return ""; return pick(v, MAJORS) ? "" : "Choose your major from the list."; },
    minor: function (v) {
      if (!isUG()) return "";
      var m = curMajor(), list = allowed();
      if (!m || !list.length || !collapse(v)) return "";
      return pick(v, list) ? "" : "Pick a minor from the list. That one is not offered with " + m + ".";
    },
    specialisation: function (v) { if (isUG()) return ""; return progBad(collapse(v)) ? "Enter your specialisation." : ""; },
    studentId: function (v) { return /^[A-Za-z0-9][A-Za-z0-9\-_/.]{2,29}$/.test(v) ? "" : "Enter your Student ID: 3 to 30 letters, numbers or - _ / ."; },
    studentId2: function (v) { return !v ? "Type your Student ID again to confirm it." : (v === val("studentId") ? "" : "This does not match the Student ID above."); }
  };
  function setErr(k, msg) {
    var input = byId("f-" + k), err = byId("e-" + k);
    if (!err) return;
    if (msg) { err.textContent = msg; err.hidden = false; if (input) input.setAttribute("aria-invalid", "true"); }
    else { err.textContent = ""; err.hidden = true; if (input) input.removeAttribute("aria-invalid"); }
  }
  function focusField(k) {
    var n = byId("f-" + k); if (!n) return;
    if (k === "batch") n = n.querySelector("input");
    if (n) n.focus();
  }
  function validateStep(n) {
    var first = null;
    stepFields[n].forEach(function (k) {
      var msg = rules[k](val(k));
      setErr(k, msg);
      if (msg && !first) first = k;
    });
    return first;
  }
  ["name", "phone", "email"].forEach(function (k) { byId("f-" + k).addEventListener("input", function () { setErr(k, ""); }); });

  /* ---- Student ID: typed twice, no pasting into the second box ---- */
  function matchHint(a, b, hintId) {
    var m = byId(hintId);
    if (!b) { m.textContent = ""; m.className = "hint"; }
    else if (a === b) { m.textContent = "Matches"; m.className = "hint good"; }
    else if (a.indexOf(b) === 0) { m.textContent = ""; m.className = "hint"; }
    else { m.textContent = "Does not match yet"; m.className = "hint bad"; }
  }
  byId("f-studentId2").addEventListener("paste", function (e) { e.preventDefault(); var m = byId("m-studentId2"); m.textContent = "Pasting is off here. Type it again."; m.className = "hint"; });
  byId("f-studentId2").addEventListener("drop", function (e) { e.preventDefault(); });
  byId("f-studentId2").addEventListener("input", function () { setErr("studentId2", ""); matchHint(val("studentId"), val("studentId2"), "m-studentId2"); });
  byId("f-studentId").addEventListener("input", function () { setErr("studentId", ""); matchHint(val("studentId"), val("studentId2"), "m-studentId2"); });

  /* ---- steps ---- */
  function say(msg, kind) { var s = byId("status"); s.textContent = msg || ""; s.className = "status" + (kind ? " " + kind : ""); }
  function goStep(n, focus) {
    step = n;
    Array.prototype.forEach.call(document.querySelectorAll(".panel[data-step]"), function (p) { p.hidden = Number(p.getAttribute("data-step")) !== n; });
    Array.prototype.forEach.call(document.querySelectorAll("#steps li"), function (li, i) {
      li.classList.toggle("done", i + 1 < n);
      if (i + 1 === n) li.setAttribute("aria-current", "step"); else li.removeAttribute("aria-current");
    });
    byId("stepcount").textContent = "Step " + n + " of " + TOTAL;
    byId("back").hidden = n === 1;
    byId("next").textContent = n === TOTAL ? "Confirm and submit" : "Continue";
    if (n === 3) applyBatch();
    if (n === TOTAL) buildReview();
    say("");
    if (focus !== false) byId("h-" + n).focus();
  }
  Array.prototype.forEach.call(document.querySelectorAll("#steps li"), function (li, i) {
    li.addEventListener("click", function () { if (i + 1 < step && !byId("reg").hidden) goStep(i + 1); });
  });

  function buildReview() {
    var ug = isUG(), list = allowed(), studies, summaryProg;
    if (ug) {
      var minorShown = !list.length ? "N/A" : (pick(val("minor"), list) || "None");
      studies = [["Major", curMajor()], ["Minor", minorShown]];
      summaryProg = curMajor() + (minorShown !== "N/A" && minorShown !== "None" ? " / " + minorShown : "");
    } else {
      var sp = collapse(val("specialisation"));
      studies = [["Specialisation", sp]];
      summaryProg = sp;
    }
    var banner = byId("proofBanner"); banner.textContent = "";
    var left = el("div"), pn = el("p", "proof-name"), pm = el("p", "proof-meta"), pid = el("div", "proof-id");
    pn.textContent = collapse(val("name"));
    pm.textContent = getBatch() + " — " + summaryProg;
    pid.textContent = val("studentId");
    left.append(pn, pm); banner.append(left, pid);

    var groups = [
      { title: "Contact", step: 1, rows: [["Name", collapse(val("name"))], ["Mobile", val("phone")], ["Email", val("email")]] },
      { title: "Batch", step: 2, rows: [["Batch", getBatch()]] },
      { title: "Studies", step: 3, rows: studies },
      { title: "Identification", step: 4, rows: [["Student ID", val("studentId"), true], ["Photocode", "Added by the organiser"]] }
    ];
    var host = byId("review"); host.textContent = "";
    groups.forEach(function (g) {
      var sec = el("section", "rgroup"), head = el("div", "rhead"), h = el("h3"), b = el("button", "link");
      h.textContent = g.title; b.type = "button"; b.textContent = "Edit";
      b.setAttribute("aria-label", "Edit " + g.title.toLowerCase());
      b.addEventListener("click", function () { goStep(g.step); });
      head.append(h, b);
      var dl = el("dl");
      g.rows.forEach(function (r) {
        var dt = el("dt"), dd = el("dd", r[2] ? "mono" : "");
        dt.textContent = r[0]; dd.textContent = r[1]; dl.append(dt, dd);
      });
      sec.append(head, dl); host.append(sec);
    });
  }

  /* ---- submit ---- */
  function payload() {
    var ug = isUG(), list = ug ? allowed() : [];
    return {
      name: collapse(val("name")), phone: val("phone"), email: val("email"), batch: getBatch(),
      major: ug ? curMajor() : "", minor: ug && list.length ? pick(val("minor"), list) : "",
      specialisation: ug ? "" : collapse(val("specialisation")),
      studentId: val("studentId"), studentId2: val("studentId2")
    };
  }
  function submit() {
    if (busy) return;
    busy = true; byId("next").disabled = true; say("Saving...");
    fetch("/api/submit", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify(payload()) })
      .then(function (r) { return r.json().catch(function () { return { ok: false, errors: { form: "Unexpected reply from the server." } }; }); })
      .then(function (res) {
        busy = false; byId("next").disabled = false;
        if (res.ok) { showDone(res); return; }
        var errs = res.errors || {}, firstKey = Object.keys(errs)[0];
        if (firstKey && stepOf[firstKey]) {
          goStep(stepOf[firstKey], false);
          Object.keys(errs).forEach(function (k) { setErr(k, errs[k]); });
          say("Please fix the highlighted field.", "bad");
          focusField(firstKey);
        } else say(errs.form || "Could not save. Try again.", "bad");
      })
      .catch(function () {
        busy = false; byId("next").disabled = false;
        say("Could not reach the server. Check your connection and press Confirm and submit again. Nothing is lost.", "bad");
      });
  }
  function showDone(res) {
    byId("reg").hidden = true; byId("steps").hidden = true;
    byId("stepcount").textContent = "Saved";
    var prog = (res.major && res.major !== "N/A")
      ? (res.major + (res.minor && res.minor !== "N/A" && res.minor !== "None" ? " (" + res.minor + ")" : ""))
      : (res.specialisation && res.specialisation !== "N/A" ? res.specialisation : "-");
    var sum = byId("doneSummary"); sum.textContent = "";
    [["Name", res.name], ["Batch", res.batch || "-"], ["Studies", prog], ["Student ID", res.studentId], ["Saved", res.savedAt]].forEach(function (r) {
      var dt = el("dt"), dd = el("dd", r[0] === "Student ID" ? "mono" : "");
      dt.textContent = r[0]; dd.textContent = r[1]; sum.append(dt, dd);
    });
    byId("doneMsg").textContent = (res.replaced ? "An earlier entry with this Student ID was updated. " : "") + "Show this screen at the photoshoot desk. The organiser will add your photocode separately.";
    byId("done").hidden = false; byId("h-done").focus();
  }
  byId("again").addEventListener("click", function () {
    byId("reg").reset(); lastMajor = "";
    Object.keys(rules).forEach(function (k) { setErr(k, ""); });
    var m = byId("m-studentId2"); m.textContent = ""; m.className = "hint";
    applyBatch(); syncMinor(); majorCombo.syncClear();
    byId("done").hidden = true; byId("reg").hidden = false; byId("steps").hidden = false;
    goStep(1);
  });
  byId("back").addEventListener("click", function () { if (step > 1) goStep(step - 1); });
  byId("reg").addEventListener("submit", function (e) {
    e.preventDefault();
    if (step < TOTAL) {
      var bad = validateStep(step);
      if (bad) { say("Fix the highlighted field to continue.", "bad"); focusField(bad); return; }
      goStep(step + 1);
    } else {
      for (var s = 1; s <= 4; s++) {
        var b2 = validateStep(s);
        if (b2) { goStep(s, false); say("Fix the highlighted field to continue.", "bad"); focusField(b2); return; }
      }
      submit();
    }
  });

  applyBatch(); syncMinor(); goStep(1, false);
})();
</script>
</body></html>
"""

ADMIN_HTML = r"""<!doctype html>
<html lang="en"><head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<title>Years at FLAME — Registrations</title>
<style>/*BASE_CSS*/</style>
</head><body>
<header class="topbar">
  <div class="topbar-inner wide">
    <div class="brand">
      <img src="__LOGO__" alt="Years at FLAME">
      <span class="brand-name">Years at FLAME</span>
      <span class="brand-sep" aria-hidden="true">/</span>
      <span class="brand-sub">Admin</span>
    </div>
    <div style="display:flex;gap:10px;align-items:center">
      <button type="button" class="btn secondary small" id="openQrBtn">Student link &amp; QR</button>
      <a class="btn primary small" href="/admin/export.xlsx" download>Download Excel</a>
    </div>
  </div>
</header>

<main class="main wide">
  <div style="display:flex;flex-wrap:wrap;gap:8px 16px;align-items:baseline;justify-content:space-between">
    <h1>Registrations</h1>
    <p class="lede" id="count" style="font-size:14px">Loading...</p>
  </div>

  <section class="metrics" aria-label="Summary">
    <div class="metric">
      <span class="metric-label">Registered</span>
      <span class="metric-val" id="kpiTotal">0</span>
      <p class="metric-sub" id="storageSub">Saved locally</p>
    </div>
    <div class="metric">
      <span class="metric-label">Awaiting photocode</span>
      <span class="metric-val" id="kpiPending" style="color:var(--gold-ink)">0</span>
      <p class="metric-sub"><button type="button" class="link" id="kpiPendingBtn" style="font-size:13px">Filter queue</button></p>
    </div>
    <div class="metric">
      <span class="metric-label">Photocodes confirmed</span>
      <span class="metric-val" id="kpiDone" style="color:var(--brand)">0</span>
      <div class="bar-track"><div class="bar-fill" id="kpiBar"></div></div>
    </div>
    <div class="metric">
      <span class="metric-label">Student link</span>
      <p class="metric-sub" id="kpiMode" style="color:var(--ink);font-weight:500">Checking...</p>
      <div style="display:flex;gap:14px;margin-top:4px">
        <button type="button" class="link" id="copyUrlBtn" style="font-size:13px">Copy link</button>
        <button type="button" class="link" id="quickQrBtn" style="font-size:13px">Show QR</button>
      </div>
    </div>
  </section>

  <section class="toolbar" aria-label="Filters">
    <div class="searchbox">
      <input id="search" type="search" placeholder="Search name, Student ID, photocode, mobile, major" aria-label="Search registrations">
    </div>
    <div style="display:flex;flex-wrap:wrap;gap:10px;align-items:center">
      <div class="seg-group" id="batchChips" role="group" aria-label="Filter by batch">
        <button type="button" class="seg active" data-batch="">All batches</button>
        <button type="button" class="seg" data-batch="UG3">UG3</button>
        <button type="button" class="seg" data-batch="UG4">UG4</button>
        <button type="button" class="seg" data-batch="MSC">MSC</button>
        <button type="button" class="seg" data-batch="PG2">PG2</button>
        <button type="button" class="seg" data-batch="PGPEI">PGPEI</button>
      </div>
      <div class="seg-group" id="statusChips" role="group" aria-label="Filter by status">
        <button type="button" class="seg active" data-st="all">All</button>
        <button type="button" class="seg" data-st="pending">Pending</button>
        <button type="button" class="seg" data-st="confirmed">Confirmed</button>
      </div>
    </div>
  </section>

  <p class="status" id="status" role="status"></p>
  <div class="tablewrap"><table>
    <thead><tr><th>No.</th><th>Submitted</th><th>Name</th><th>Mobile</th><th>Email</th><th>Batch</th><th>Major</th><th>Minor</th><th>Specialisation</th><th>Student ID</th><th>Photocode</th><th></th></tr></thead>
    <tbody id="body"></tbody>
  </table></div>
</main>

<dialog id="photoDlg" aria-labelledby="dlgTitle">
  <form class="dlg" id="photoForm" novalidate>
    <h2 id="dlgTitle">Add photocode</h2>
    <dl id="dlgWho"></dl>
    <div class="field"><label for="p1">Photocode</label>
      <input id="p1" class="mono" type="text" autocomplete="off" autocapitalize="off" spellcheck="false"></div>
    <div class="field"><label for="p2">Re-enter photocode</label>
      <input id="p2" class="mono" type="text" autocomplete="off" autocapitalize="off" spellcheck="false" aria-describedby="pm">
      <p class="hint" id="pm" aria-live="polite"></p></div>
    <p class="readback" id="readback" aria-live="polite"></p>
    <p class="err" id="dlgErr" role="alert" hidden></p>
    <div class="actions">
      <button type="button" class="btn secondary" id="dlgCancel">Cancel</button>
      <button type="submit" class="btn primary" id="dlgSave">Save photocode</button>
    </div>
  </form>
</dialog>

<dialog id="qrDlg" aria-labelledby="qrTitle">
  <div class="dlg">
    <div class="rhead no-print">
      <h2 id="qrTitle">Student link &amp; QR standee</h2>
      <button type="button" class="link" id="qrClose">Close</button>
    </div>
    <div class="qr-standee" id="qrPrintCard">
      <div style="display:flex;align-items:center;gap:10px">
        <img src="__LOGO__" alt="Years at FLAME" style="width:42px;height:42px">
        <div style="text-align:left">
          <div style="font-size:13px;font-weight:600;color:var(--brand)">Years at FLAME</div>
          <div style="font-size:17px;font-weight:600;color:var(--ink)">Photoshoot registration</div>
        </div>
      </div>
      <div class="qr-frame"><img id="qrImg" alt="Scan to open student registration"></div>
      <div class="qr-url" id="qrPrimaryUrl">Loading...</div>
      <p style="margin:0;font-size:13px;color:var(--muted)">Scan with your phone camera to register before your portrait.</p>
    </div>
    <div class="alt-links no-print" id="altLinks"></div>
    <div class="actions no-print">
      <button type="button" class="btn secondary small" id="qrCopyBtn">Copy link</button>
      <button type="button" class="btn primary small" id="qrPrintBtn">Print QR standee</button>
    </div>
  </div>
</dialog>

<script>
(function () {
  "use strict";
  var byId = function (id) { return document.getElementById(id); };
  var body = byId("body"), dlg = byId("photoDlg"), qrDlg = byId("qrDlg"), p1 = byId("p1"), p2 = byId("p2");
  var rows = [], pendingRemove = null, rmTimer = null, target = null;
  var filterBatch = "", filterStatus = "all", query = "", linkInfo = null;
  function el(t, c) { var e = document.createElement(t); if (c) e.className = c; return e; }
  function say(m, k) { var s = byId("status"); s.textContent = m || ""; s.className = "status" + (k ? " " + k : ""); }

  function matchesFilter(r) {
    if (filterBatch && (r.batch || "") !== filterBatch) return false;
    if (filterStatus === "pending" && r.photocode) return false;
    if (filterStatus === "confirmed" && !r.photocode) return false;
    if (query) {
      var hay = [r.name, r.studentId, r.photocode, r.phone, r.email, r.batch, r.major, r.minor, r.specialisation].join(" ").toLowerCase();
      if (hay.indexOf(query) === -1) return false;
    }
    return true;
  }

  function render() {
    body.textContent = "";
    var total = rows.length;
    var waiting = rows.filter(function (r) { return !r.photocode; }).length;
    var done = total - waiting;
    var pct = total ? Math.round((done / total) * 100) : 0;

    byId("kpiTotal").textContent = String(total);
    byId("kpiPending").textContent = String(waiting);
    byId("kpiDone").textContent = String(done) + " (" + pct + "%)";
    byId("kpiBar").style.width = pct + "%";

    byId("count").textContent = total + (total === 1 ? " registration" : " registrations") +
      " · " + waiting + " awaiting a photocode · refreshes every 10s";

    var list = rows.filter(matchesFilter);
    if (!list.length) {
      var tr = el("tr"), td = el("td", "empty"); td.colSpan = 12;
      td.textContent = total ? "No registrations match the current search or filter." : "No registrations yet. They appear here as students submit.";
      tr.append(td); body.append(tr); return;
    }
    list.forEach(function (r) {
      var tr = el("tr");
      var cells = [[String(rows.length - rows.indexOf(r)), ""], [r.submittedAt, ""], [r.name, ""], [r.phone, ""], [r.email, ""],
                   [r.batch || "-", ""], [r.major, ""], [r.minor, ""], [r.specialisation, ""], [r.studentId, "mono"]];
      cells.forEach(function (c, ci) {
        var td = el("td", c[1]);
        if (ci === 5) {
          var bp = el("span", "tag"); bp.textContent = c[0]; td.append(bp);
        } else {
          td.textContent = c[0];
        }
        if (ci === 9 && r.times > 1) { var p = el("span", "tag"); p.style.marginLeft = "8px"; p.textContent = "Updated x" + r.times; td.append(p); }
        tr.append(td);
      });
      var pc = el("td"), wrap = el("div", "pcell"), btn = el("button", "link");
      btn.type = "button";
      if (r.photocode) {
        var code = el("span", "mono"); code.textContent = r.photocode;
        var pill = el("span", "tag good"); pill.textContent = "Confirmed";
        pill.title = "Entered twice on " + r.photocodeAt;
        btn.textContent = "Change"; btn.setAttribute("aria-label", "Change photocode for " + r.name);
        wrap.append(code, pill, btn);
      } else {
        var pend = el("span", "tag warn"); pend.textContent = "Pending";
        btn.textContent = "Add photocode"; btn.setAttribute("aria-label", "Add photocode for " + r.name);
        wrap.append(pend, btn);
      }
      btn.addEventListener("click", function () { openDlg(r); });
      pc.append(wrap); tr.append(pc);
      var act = el("td"), rm = el("button", "link bad"), k = r.studentId;
      rm.type = "button"; rm.textContent = pendingRemove === k ? "Confirm remove" : "Remove";
      rm.setAttribute("aria-label", "Remove " + r.name);
      rm.addEventListener("click", function () { remove(k); });
      act.append(rm); tr.append(act); body.append(tr);
    });
  }

  function load() {
    fetch("/admin/data").then(function (r) { if (!r.ok) throw new Error(); return r.json(); })
      .then(function (d) { rows = d.rows || []; render(); })
      .catch(function () { say("Could not refresh. Is the server still running?", "bad"); });
  }
  function loadInfo() {
    fetch("/admin/info").then(function (r) { if (!r.ok) throw new Error(); return r.json(); })
      .then(function (d) {
        linkInfo = d;
        byId("kpiMode").textContent = d.label || d.primaryUrl;
        byId("kpiMode").title = d.primaryUrl || "";
        if (byId("storageSub")) {
          byId("storageSub").textContent = d.storage === "turso" ? "Synced to Turso DB" : "Saved to registrations.xlsx";
        }
      })
      .catch(function () { byId("kpiMode").textContent = "Unavailable"; });
  }
  function copyText(txt, msg) {
    if (!txt) return;
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(txt).then(function () { say(msg || "Copied link to clipboard."); });
    } else {
      window.prompt("Copy this link:", txt);
    }
  }
  function openQrModal() {
    var url = (linkInfo && linkInfo.primaryUrl) ? linkInfo.primaryUrl : window.location.origin;
    byId("qrPrimaryUrl").textContent = url;
    byId("qrImg").src = "/admin/qr.svg?url=" + encodeURIComponent(url) + "&_=" + Date.now();
    var alt = byId("altLinks"); alt.textContent = "";
    if (linkInfo) {
      var l1 = el("div"); l1.textContent = "Active mode: " + linkInfo.label + (linkInfo.isConstant ? " (constant)" : " (changes per Wi-Fi)");
      var l2 = el("div"); l2.textContent = "Same Wi-Fi (.local): " + linkInfo.mdnsUrl;
      var l3 = el("div"); l3.textContent = "LAN IP fallback: " + linkInfo.lanUrl;
      alt.append(l1, l2, l3);
    }
    qrDlg.showModal();
  }

  byId("openQrBtn").addEventListener("click", openQrModal);
  byId("quickQrBtn").addEventListener("click", openQrModal);
  byId("qrClose").addEventListener("click", function () { qrDlg.close(); });
  byId("qrCopyBtn").addEventListener("click", function () { copyText(byId("qrPrimaryUrl").textContent, "Student link copied."); });
  byId("copyUrlBtn").addEventListener("click", function () { copyText(linkInfo ? linkInfo.primaryUrl : window.location.origin, "Student link copied."); });
  byId("qrPrintBtn").addEventListener("click", function () { window.print(); });

  function remove(id) {
    if (pendingRemove !== id) {
      pendingRemove = id; clearTimeout(rmTimer);
      rmTimer = setTimeout(function () { pendingRemove = null; render(); }, 4000);
      render(); return;
    }
    pendingRemove = null; clearTimeout(rmTimer);
    fetch("/admin/delete", { method: "POST", headers: { "Content-Type": "application/json" }, body: JSON.stringify({ studentId: id }) })
      .then(function (r) { if (!r.ok) throw new Error(); say("Entry removed."); load(); })
      .catch(function () { say("Could not remove that entry.", "bad"); });
  }

  byId("search").addEventListener("input", function (e) { query = e.target.value.trim().toLowerCase(); render(); });
  function bindChipGroup(containerId, attr, onSelect) {
    var box = byId(containerId);
    box.addEventListener("click", function (e) {
      var btn = e.target.closest("button.seg"); if (!btn) return;
      Array.prototype.forEach.call(box.querySelectorAll(".seg"), function (c) { c.classList.toggle("active", c === btn); });
      onSelect(btn.getAttribute(attr) || "");
      render();
    });
  }
  bindChipGroup("batchChips", "data-batch", function (v) { filterBatch = v; });
  bindChipGroup("statusChips", "data-st", function (v) { filterStatus = v || "all"; });
  byId("kpiPendingBtn").addEventListener("click", function () {
    filterStatus = (filterStatus === "pending") ? "all" : "pending";
    Array.prototype.forEach.call(byId("statusChips").querySelectorAll(".seg"), function (c) {
      c.classList.toggle("active", c.getAttribute("data-st") === filterStatus);
    });
    render();
  });

  /* ---- photocode dialog ---- */
  function showErr(msg) { var e = byId("dlgErr"); e.textContent = msg || ""; e.hidden = !msg; }
  function updateReadback() {
    var a = p1.value.trim(), pm = byId("pm"), b = p2.value.trim();
    byId("readback").textContent = a && target ? "Save " + a + " against " + target.name + " (" + target.studentId + ")." : "";
    if (!b) { pm.textContent = ""; pm.className = "hint"; }
    else if (a === b) { pm.textContent = "Matches"; pm.className = "hint good"; }
    else if (a.indexOf(b) === 0) { pm.textContent = ""; pm.className = "hint"; }
    else { pm.textContent = "Does not match yet"; pm.className = "hint bad"; }
  }
  function programme(r) {
    if (r.major && r.major !== "N/A") return r.major + " (minor: " + r.minor + ")";
    return r.specialisation && r.specialisation !== "N/A" ? r.specialisation : "-";
  }
  function openDlg(r) {
    target = r;
    byId("dlgTitle").textContent = r.photocode ? "Change photocode" : "Add photocode";
    var who = byId("dlgWho"); who.textContent = "";
    [["Name", r.name], ["Student ID", r.studentId], ["Batch", r.batch || "-"], ["Studies", programme(r)]].forEach(function (x) {
      var dt = el("dt"), dd = el("dd", x[0] === "Student ID" ? "mono" : "");
      dt.textContent = x[0]; dd.textContent = x[1]; who.append(dt, dd);
    });
    p1.value = ""; p2.value = ""; showErr(""); updateReadback();
    byId("dlgSave").disabled = false;
    dlg.showModal(); p1.focus();
  }
  p1.addEventListener("input", function () { showErr(""); updateReadback(); });
  p2.addEventListener("input", function () { showErr(""); updateReadback(); });
  p2.addEventListener("paste", function (e) { e.preventDefault(); var pm = byId("pm"); pm.textContent = "Pasting is off here. Type it again."; pm.className = "hint"; });
  p2.addEventListener("drop", function (e) { e.preventDefault(); });
  byId("dlgCancel").addEventListener("click", function () { dlg.close(); });
  byId("photoForm").addEventListener("submit", function (e) {
    e.preventDefault();
    var a = p1.value.trim(), b = p2.value.trim();
    if (!a) { showErr("Type the photocode."); p1.focus(); return; }
    if (a !== b) { showErr("The two entries do not match."); p2.focus(); return; }
    var save = byId("dlgSave"); save.disabled = true;
    fetch("/admin/photocode", { method: "POST", headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ studentId: target.studentId, photocode: a, photocode2: b }) })
      .then(function (r) { return r.json().catch(function () { return { ok: false, error: "Unexpected reply from the server." }; }); })
      .then(function (j) {
        save.disabled = false;
        if (j.ok) { dlg.close(); say("Photocode saved for " + j.name + "."); load(); }
        else showErr(j.error || "Could not save the photocode.");
      })
      .catch(function () { save.disabled = false; showErr("Could not reach the server. Try again."); });
  });

  loadInfo(); load(); setInterval(load, 10000);
})();
</script>
</body></html>
"""


if __name__ == "__main__":
    urls = resolve_student_urls()
    print("\n  Years at FLAME — Photoshoot Registration is running.\n")
    print("  Student Link  : %s  [%s]" % (urls["primaryUrl"], urls["label"]))
    if urls["primaryUrl"] != urls["mdnsUrl"]:
        print("  Same Wi-Fi    : %s" % urls["mdnsUrl"])
    if urls["primaryUrl"] != urls["lanUrl"]:
        print("  LAN Fallback  : %s" % urls["lanUrl"])
    print("  Admin Console : http://localhost:%d/admin" % PORT)
    print("                  user: admin   password: %s" % ADMIN_PASSWORD)
    print("  Data saved in : %s" % HERE)
    if urls["mode"] in ("mdns", "lan"):
        print("\n  Tip: For a constant public HTTPS link across 4G & campus Wi-Fi,")
        print("       set PUBLIC_URL in .env (see CONSTANT_LINK_GUIDE.md).")
    print("\n  Press Ctrl+C to stop.\n")
    try:
        from waitress import serve
    except ImportError:
        serve = None
    if serve:
        serve(app, host="0.0.0.0", port=PORT, threads=8)
    else:
        app.run(host="0.0.0.0", port=PORT, threaded=True)
