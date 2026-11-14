#!/usr/bin/env python3
"""
Embedded Job Checker - Noida / Gurgaon
---------------------------------------
Excel list ki har company ke careers portal pe jaake embedded / firmware
jobs dhoondta hai aur result Excel mein deta hai. Pichli run ke baad aayi
nayi jobs "NEW" mark hoti hain.

Usage:
  pip install -r requirements.txt
  python job_checker.py                    # ek baar run
  python job_checker.py --every 60         # har 60 min mein dobara
  python job_checker.py --render           # JS-wale pages ke liye (playwright chahiye)
  python job_checker.py --any-location     # location filter band
  python job_checker.py --rediscover       # portal cache reset karke phir se dhoondo
  python job_checker.py --selftest         # bina internet ke parser test
"""
import argparse
import datetime as dt
import hashlib
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from urllib.parse import quote, urljoin, urlparse

HERE = Path(__file__).resolve().parent
DEFAULT_XLSX = HERE / "Embedded_Companies_Noida_Gurgaon.xlsx"
PORTALS_YAML = HERE / "portals.yaml"
CACHE_FILE = HERE / ".portal_cache.json"
SEEN_FILE = HERE / ".seen_jobs.json"
OUT_DIR = HERE / "results"

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}
TIMEOUT = 25
MAX_PER_TERM = 200          # Workday / Oracle pagination limit per search term

# ---- Filters: yahan apne hisaab se badlo ---------------------------------
SEARCH_TERMS = ["embedded", "firmware"]
ROLE_RE = re.compile(
    r"embedded|firmware|\bbsp\b|autosar|rtos|device driver|linux kernel|bootloader|"
    r"microcontroller|\bmcu\b|\becu\b|j1939|\buds\b|board support|low[- ]level|\bhil\b",
    re.I)
LOC_RE = re.compile(
    r"noida|gurgaon|gurugram|manesar|delhi|\bncr\b|faridabad|ghaziabad|dharuhera|bawal",
    re.I)
VAGUE_LOC_RE = re.compile(r"\d+\s+locations|multiple|^\s*india\s*$|remote", re.I)

# ---- Portal detection -----------------------------------------------------
PATTERNS = [
    ("workday", re.compile(
        r"https?://([a-z0-9-]+)\.(wd\d+)\.myworkdayjobs\.com/(?:[a-z]{2}-[A-Z]{2}/)?([A-Za-z0-9_\-]+)")),
    ("greenhouse", re.compile(
        r"https?://(?:boards|job-boards)(?:\.eu)?\.greenhouse\.io/(?:embed/job_board(?:/js)?\?for=)?([A-Za-z0-9_\-]+)")),
    ("lever", re.compile(r"https?://jobs\.lever\.co/([A-Za-z0-9_\-]+)")),
    ("smartrecruiters", re.compile(
        r"https?://(?:jobs|careers)\.smartrecruiters\.com/([A-Za-z0-9_\-]+)")),
    ("oracle", re.compile(
        r"https?://([a-z0-9.-]+\.oraclecloud\.com)/hcmUI/CandidateExperience/[a-z]{2}(?:-[A-Z]{2})?/sites/([A-Za-z0-9_]+)")),
]
A_RE = re.compile(r"<a\b[^>]*?href=[\"']([^\"'#]+)[\"'][^>]*>(.*?)</a>", re.I | re.S)
TAG_RE = re.compile(r"<[^>]+>")
HOP_RE = re.compile(r"job|career|opening|vacanc|search|apply|join", re.I)


def session():
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(total=2, backoff_factor=1, status_forcelist=[429, 500, 502, 503, 504],
                  allowed_methods=["GET", "POST"])
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://", HTTPAdapter(max_retries=retry))
    return s


def fetch(s, url):
    """(final_url, html) ya (url, None) agar fail."""
    try:
        r = s.get(url, timeout=TIMEOUT, allow_redirects=True)
        if r.status_code >= 400:
            return url, None
        if "html" not in r.headers.get("Content-Type", "html"):
            return r.url, None
        return r.url, r.text
    except Exception:
        return url, None


def detect(url, html=""):
    blob = (url or "") + "\n" + (html or "")
    for typ, rx in PATTERNS:
        for m in rx.finditer(blob):
            if typ == "workday":
                tenant, wd, site = m.groups()
                if site.lower() in ("wday", "en-us", "job"):
                    continue
                return {"type": "workday", "tenant": tenant, "wd": wd, "site": site, "url": m.group(0)}
            if typ == "oracle":
                return {"type": "oracle", "host": m.group(1), "site": m.group(2), "url": m.group(0)}
            token = m.group(1)
            if token.lower() in ("embed", "api", "v1"):
                continue
            return {"type": typ, "id": token, "url": m.group(0)}
    if html and "eightfold" in html.lower():
        return {"type": "eightfold", "host": urlparse(url).netloc, "url": url}
    return None


def candidates(website):
    if not website:
        return []
    w = website.strip().strip("/")
    w = re.sub(r"^https?://", "", w)
    host, _, path = w.partition("/")
    base = f"https://{host}" + (f"/{path}" if path else "")
    bare = host[4:] if host.startswith("www.") else host
    return [f"{base}/careers", f"https://www.{bare}/careers", f"https://careers.{bare}",
            f"https://jobs.{bare}", f"{base}/en/careers", f"{base}/career", f"{base}/careers.html"]


def resolve(s, name, website, ov):
    """Company ka portal dhoondo. ov = portals.yaml override."""
    if ov.get("type") in ("eightfold", "generic") and ov.get("url"):
        info = {"type": ov["type"], "url": ov["url"], "host": urlparse(ov["url"]).netloc}
        if ov.get("domain"):
            info["domain"] = ov["domain"]
        return info
    cands = ([ov["url"]] if ov.get("url") else []) + candidates(website)
    first_ok, tried = None, set()
    for u in cands:
        if u in tried:
            continue
        tried.add(u)
        final, html = fetch(s, u)
        if html is None:
            continue
        hit = detect(final, html)
        if hit:
            return hit
        first_ok = first_ok or final
        # Careers page se "Search jobs" jaise links follow karo (max 3)
        hops = []
        for href, text in A_RE.findall(html)[:500]:
            full = urljoin(final, href.strip())
            hit = detect(full)
            if hit:
                return hit
            if HOP_RE.search(href + " " + TAG_RE.sub(" ", text)) and full.startswith("http"):
                if urlparse(full).netloc != urlparse(final).netloc and full not in hops:
                    hops.append(full)
        for h in hops[:3]:
            f2, h2 = fetch(s, h)
            if h2:
                hit = detect(f2, h2)
                if hit:
                    return hit
        if ov.get("url"):
            break   # override diya hai to wahi page generic maano
    if first_ok:
        return {"type": "generic", "url": first_ok}
    return None


# ---- Fetchers (har portal type ke liye) -----------------------------------
def J(company, title, loc, url, posted, src):
    return {"company": company, "title": (title or "").strip(), "location": (loc or "").strip(),
            "url": url or "", "posted": str(posted or "")[:25], "portal": src, "note": ""}


def f_workday(s, info, company, website):
    host = f"{info['tenant']}.{info['wd']}.myworkdayjobs.com"
    api = f"https://{host}/wday/cxs/{info['tenant']}/{info['site']}/jobs"
    out = []
    for term in SEARCH_TERMS:
        offset, total = 0, None
        while offset < MAX_PER_TERM:
            r = s.post(api, json={"appliedFacets": {}, "limit": 20, "offset": offset, "searchText": term},
                       headers={"Accept": "application/json", "Content-Type": "application/json"},
                       timeout=TIMEOUT)
            r.raise_for_status()
            d = r.json()
            posts = d.get("jobPostings") or []
            if total is None:
                total = d.get("total") or 0
            for p in posts:
                out.append(J(company, p.get("title"), p.get("locationsText"),
                             f"https://{host}/{info['site']}{p.get('externalPath', '')}",
                             p.get("postedOn"), "Workday"))
            offset += 20
            if not posts or offset >= total:
                break
    return out


def f_greenhouse(s, info, company, website):
    r = s.get(f"https://boards-api.greenhouse.io/v1/boards/{info['id']}/jobs", timeout=TIMEOUT)
    r.raise_for_status()
    return [J(company, j.get("title"), (j.get("location") or {}).get("name"), j.get("absolute_url"),
              j.get("updated_at"), "Greenhouse") for j in r.json().get("jobs", [])]


def f_lever(s, info, company, website):
    r = s.get(f"https://api.lever.co/v0/postings/{info['id']}?mode=json", timeout=TIMEOUT)
    r.raise_for_status()
    out = []
    for j in r.json():
        ts = j.get("createdAt")
        posted = dt.datetime.fromtimestamp(ts / 1000).date() if ts else ""
        out.append(J(company, j.get("text"), (j.get("categories") or {}).get("location"),
                     j.get("hostedUrl"), posted, "Lever"))
    return out


def f_smartrecruiters(s, info, company, website):
    out = []
    for term in SEARCH_TERMS:
        r = s.get(f"https://api.smartrecruiters.com/v1/companies/{info['id']}/postings",
                  params={"q": term, "limit": 100}, timeout=TIMEOUT)
        r.raise_for_status()
        for p in r.json().get("content", []):
            loc = p.get("location") or {}
            out.append(J(company, p.get("name"), f"{loc.get('city', '')}, {loc.get('country', '')}",
                         f"https://jobs.smartrecruiters.com/{info['id']}/{p.get('id')}",
                         p.get("releasedDate"), "SmartRecruiters"))
    return out


def f_oracle(s, info, company, website):
    out = []
    for term in SEARCH_TERMS:
        finder = (f"findReqs;siteNumber={info['site']},facetsList=LOCATIONS,"
                  f"limit=100,keyword=\"{term}\",sortBy=POSTING_DATES_DESC")
        url = (f"https://{info['host']}/hcmRestApi/resources/latest/recruitingCEJobRequisitions"
               f"?onlyData=true&expand=requisitionList&finder={quote(finder, safe=';=,')}")
        r = s.get(url, headers={"Accept": "application/json"}, timeout=TIMEOUT)
        r.raise_for_status()
        for item in r.json().get("items", []):
            for p in item.get("requisitionList", []):
                out.append(J(company, p.get("Title"), p.get("PrimaryLocation"),
                             f"https://{info['host']}/hcmUI/CandidateExperience/en/sites/{info['site']}/job/{p.get('Id')}",
                             p.get("PostedDate"), "Oracle HCM"))
    return out


def f_eightfold(s, info, company, website):
    host = info.get("host") or urlparse(info["url"]).netloc
    domain = info.get("domain") or re.sub(r"^www\.", "", (website or host).split("/")[0])
    out = []
    for term in SEARCH_TERMS:
        r = s.get(f"https://{host}/api/apply/v2/jobs",
                  params={"domain": domain, "query": term, "start": 0, "num": 100},
                  headers={"Accept": "application/json"}, timeout=TIMEOUT)
        r.raise_for_status()
        for p in r.json().get("positions", []):
            loc = p.get("location") or ", ".join(p.get("locations") or [])
            url = p.get("canonicalPositionUrl") or f"https://{host}/careers/job/{p.get('id')}"
            posted = dt.datetime.fromtimestamp(p["t_create"]).date() if p.get("t_create") else ""
            out.append(J(company, p.get("name"), loc, url, posted, "Eightfold"))
    return out


def generic_from_html(company, page_url, html):
    out = []
    for href, text in A_RE.findall(html or ""):
        t = " ".join(TAG_RE.sub(" ", text).split())
        if t and len(t) < 160 and ROLE_RE.search(t):
            j = J(company, t, "", urljoin(page_url, href), "", "Careers page")
            j["note"] = "Simple page scan - details link pe check karo"
            out.append(j)
    return out


def f_generic(s, info, company, website):
    final, html = fetch(s, info["url"])
    return generic_from_html(company, final, html)


def render_html(url):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as p:
        b = p.chromium.launch()
        pg = b.new_page(user_agent=HEADERS["User-Agent"])
        pg.goto(url, timeout=60000, wait_until="networkidle")
        html = pg.content()
        b.close()
        return html


FETCHERS = {"workday": f_workday, "greenhouse": f_greenhouse, "lever": f_lever,
            "smartrecruiters": f_smartrecruiters, "oracle": f_oracle,
            "eightfold": f_eightfold, "generic": f_generic}


# ---- Filtering ------------------------------------------------------------
def keep(j, any_location=False):
    if not ROLE_RE.search(j["title"]):
        return False
    if any_location:
        return True
    loc = j["location"]
    if LOC_RE.search(loc):
        return True
    if not loc or VAGUE_LOC_RE.search(loc):
        j["note"] = j["note"] or "Location clear nahi - check karo"
        return True
    return False


def job_key(j):
    return hashlib.sha1(f"{j['company']}|{j['url'] or j['title']}".encode()).hexdigest()[:16]


# ---- IO -------------------------------------------------------------------
def load_json(p):
    try:
        return json.loads(Path(p).read_text(encoding="utf-8"))
    except Exception:
        return {}


def save_json(p, d):
    Path(p).write_text(json.dumps(d, indent=1, ensure_ascii=False), encoding="utf-8")


def load_yaml():
    if not PORTALS_YAML.exists():
        return {}
    import yaml
    return yaml.safe_load(PORTALS_YAML.read_text(encoding="utf-8")) or {}


def load_companies(path):
    from openpyxl import load_workbook
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb["Companies"] if "Companies" in wb.sheetnames else wb.active
    rows = list(ws.iter_rows(values_only=True))
    hdr = [str(h or "").strip().lower() for h in rows[0]]
    ci, wi = hdr.index("company"), hdr.index("website")
    out, seen = [], set()
    for r in rows[1:]:
        name = (r[ci] or "").strip() if r[ci] else ""
        site = (r[wi] or "").strip() if r[wi] else ""
        key = site.lower() or name.lower()
        if name and key not in seen:
            seen.add(key)
            out.append((name, site))
    return out


def write_excel(jobs, status, path):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    hdr = ["New?", "Company", "Job title", "Location", "Posted", "Portal", "Link", "Note"]
    ws.append(hdr)
    for c in ws[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
    green = PatternFill("solid", fgColor="C6EFCE")
    for j in jobs:
        ws.append(["NEW" if j["new"] else "", j["company"], j["title"], j["location"],
                   j["posted"], j["portal"], "Open", j["note"]])
        r = ws.max_row
        for c in ws[r]:
            c.font = Font(name="Arial", size=10)
            c.alignment = Alignment(vertical="top", wrap_text=True)
        if j["url"]:
            ws.cell(r, 7).hyperlink = j["url"]
            ws.cell(r, 7).font = Font(name="Arial", size=10, color="0563C1", underline="single")
        if j["new"]:
            ws.cell(r, 1).fill = green
            ws.cell(r, 1).font = Font(name="Arial", size=10, bold=True)
    for k, w in zip("ABCDEFGH", [7, 30, 50, 30, 14, 14, 8, 34]):
        ws.column_dimensions[k].width = w
    ws.freeze_panes = "C2"
    ws.auto_filter.ref = f"A1:H{ws.max_row}"

    st = wb.create_sheet("Portal status")
    st.append(["Company", "Portal type", "Portal URL", "Status", "Matching jobs", "Detail"])
    for c in st[1]:
        c.font = Font(name="Arial", bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor="1F3864")
    red = PatternFill("solid", fgColor="FFC7CE")
    for row in status:
        st.append(row)
        r = st.max_row
        for c in st[r]:
            c.font = Font(name="Arial", size=10)
        if row[3] != "OK":
            st.cell(r, 4).fill = red
    for k, w in zip("ABCDEF", [34, 14, 60, 12, 14, 60]):
        st.column_dimensions[k].width = w
    st.auto_filter.ref = f"A1:F{st.max_row}"
    wb.save(path)


# ---- Main run -------------------------------------------------------------
def run_once(args):
    companies = load_companies(args.companies)
    cfg = load_yaml()
    overrides = cfg.get("overrides") or {}
    skip = set(cfg.get("skip") or [])
    cache = {} if args.rediscover else load_json(CACHE_FILE)
    seen = load_json(SEEN_FILE)
    now = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    print(f"\n[{now}] {len(companies)} companies check ho rahi hain...\n")

    def work(c):
        name, website = c
        ov = overrides.get(name) or {}
        info = None if ov else cache.get(name)
        s = session()
        try:
            if not info:
                info = resolve(s, name, website, ov)
            if not info:
                return name, None, [], "Portal nahi mila - portals.yaml mein careers URL add karo"
            if info["type"] == "generic" and args.render:
                return name, info, None, "defer"
            return name, info, FETCHERS[info["type"]](s, info, name, website), ""
        except Exception as e:
            return name, info, [], f"{type(e).__name__}: {str(e)[:160]}"

    results, deferred = [], []
    todo = [c for c in companies if c[0] not in skip]
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(work, c): c for c in todo}
        for i, f in enumerate(as_completed(futs), 1):
            name, info, jobs, err = f.result()
            if err == "defer":
                deferred.append((name, info))
                continue
            results.append((name, info, jobs, err))
            tag = info["type"] if info else "-"
            print(f"  [{i:>2}/{len(todo)}] {name[:38]:<38} {tag:<15} {'ERR' if err else len(jobs)}")

    for name, info in deferred:   # playwright ek-ek karke
        try:
            html = render_html(info["url"])
            results.append((name, info, generic_from_html(name, info["url"], html), ""))
        except Exception as e:
            results.append((name, info, [], f"render: {type(e).__name__}: {str(e)[:120]}"))
        print(f"  [render] {name}")

    all_jobs, status = [], []
    for name, info, jobs, err in sorted(results, key=lambda x: x[0].lower()):
        if info and not err and info["type"] != "generic":
            cache[name] = info
        uniq = {}
        for j in jobs or []:
            if keep(j, args.any_location):
                uniq[j["url"] or j["title"]] = j
        for j in uniq.values():
            k = job_key(j)
            j["new"] = k not in seen
            seen.setdefault(k, now)
            all_jobs.append(j)
        status.append([name, info["type"] if info else "-", (info or {}).get("url", ""),
                       "ERROR" if err else ("OK" if info else "NOT FOUND"), len(uniq), err])

    all_jobs.sort(key=lambda j: (not j["new"], j["company"].lower(), j["title"].lower()))
    save_json(CACHE_FILE, cache)
    save_json(SEEN_FILE, seen)
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / f"jobs_{dt.datetime.now():%Y%m%d_%H%M}.xlsx"
    write_excel(all_jobs, status, out)

    new = [j for j in all_jobs if j["new"]]
    bad = sum(1 for r in status if r[3] != "OK")
    print(f"\nTotal matching jobs: {len(all_jobs)}  |  NEW: {len(new)}  |  Portal issues: {bad}")
    for j in new[:25]:
        print(f"  NEW  {j['company'][:25]:<25} {j['title'][:55]:<55} {j['location'][:25]}")
    print(f"\nResult file: {out}\n")


def selftest():
    cases = {
        "https://nxp.wd3.myworkdayjobs.com/en-US/careers": ("workday", "careers"),
        "https://acme.wd1.myworkdayjobs.com/External_Careers/job/Noida/X_R1": ("workday", "External_Careers"),
        "https://boards.greenhouse.io/acme": ("greenhouse", "acme"),
        "https://job-boards.greenhouse.io/embed/job_board?for=acme": ("greenhouse", "acme"),
        "https://jobs.lever.co/acme-inc": ("lever", "acme-inc"),
        "https://jobs.smartrecruiters.com/AcmeCorp": ("smartrecruiters", "AcmeCorp"),
        "https://abcd.fa.ocs.oraclecloud.com/hcmUI/CandidateExperience/en/sites/CX_1/jobs": ("oracle", "CX_1"),
    }
    ok = True
    for url, (typ, ident) in cases.items():
        d = detect(url)
        got = d and d["type"] == typ and ident in (d.get("site"), d.get("id"))
        ok &= bool(got)
        print(("PASS" if got else "FAIL"), url, "->", d)
    eg = detect("https://careers.acme.com/careers", "<script src='//static.vscdn.net/eightfold.js'>")
    ok &= bool(eg and eg["type"] == "eightfold")
    print("PASS" if eg and eg["type"] == "eightfold" else "FAIL", "eightfold html ->", eg)
    jobs = [J("A", "Embedded Software Engineer - AUTOSAR", "Noida, Uttar Pradesh", "u1", "", "t"),
            J("A", "Firmware Engineer", "2 Locations", "u2", "", "t"),
            J("A", "Embedded Engineer", "Bengaluru, India", "u3", "", "t"),
            J("A", "HR Executive", "Gurugram", "u4", "", "t")]
    exp = [True, True, False, False]
    got = [keep(j) for j in jobs]
    ok &= got == exp
    print("PASS" if got == exp else "FAIL", "filter ->", got)
    html = '<a href="/jobs/123">Senior Firmware Engineer (STM32)</a><a href="/about">About us</a>'
    g = generic_from_html("A", "https://x.com/careers", html)
    ok &= len(g) == 1 and g[0]["url"] == "https://x.com/jobs/123"
    print("PASS" if len(g) == 1 else "FAIL", "generic ->", [x["title"] for x in g])
    print("\nSELFTEST", "OK" if ok else "FAILED")
    return ok


def main():
    ap = argparse.ArgumentParser(description="Embedded job checker - Noida/Gurgaon")
    ap.add_argument("--companies", default=str(DEFAULT_XLSX), help="Company list Excel ka path")
    ap.add_argument("--every", type=int, default=0, help="Har N minute mein dobara chalao")
    ap.add_argument("--workers", type=int, default=8, help="Ek saath kitni companies check ho")
    ap.add_argument("--render", action="store_true", help="Simple pages ko browser se render karo (playwright)")
    ap.add_argument("--any-location", action="store_true", help="Location filter band karo")
    ap.add_argument("--rediscover", action="store_true", help="Portal cache bhool ke phir se dhoondo")
    ap.add_argument("--selftest", action="store_true", help="Offline parser test")
    args = ap.parse_args()
    if args.selftest:
        sys.exit(0 if selftest() else 1)
    if not Path(args.companies).exists():
        sys.exit(f"Company list nahi mili: {args.companies}\nExcel file script ke folder mein rakho ya --companies do.")
    while True:
        run_once(args)
        if not args.every:
            break
        print(f"Agla check {args.every} min baad... (Ctrl+C se band karo)")
        time.sleep(args.every * 60)


if __name__ == "__main__":
    main()
