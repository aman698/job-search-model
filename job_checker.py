#!/usr/bin/env python3
"""
Embedded Job Checker v2 - Delhi NCR (Noida / Greater Noida / Gurugram / Manesar)
--------------------------------------------------------------------------------
Excel list ki har company ke OFFICIAL careers portal pe embedded / firmware jobs
dhoondta hai aur ek Excel deta hai:  Jobs | Portal Status | Summary.

Usage:
  pip install -r requirements.txt
  python job_checker.py                     # ek baar run
  python job_checker.py --every 360         # har 6 ghante mein dobara
  python job_checker.py --only "Stryker"    # sirf ek company (debug)
  python job_checker.py --any-location      # NCR filter band (poore India/world ki jobs)
  python job_checker.py --selftest          # bina internet ke parser test

Har company ka portal recipe `portals.yaml` mein hai (Workday / Eightfold / Oracle /
SmartRecruiters / Lever / TalentBrew / Avature / SuccessFactors / Jibe / Cornerstone /
JSON API / HTML page / manual). Jo company yaml mein nahi hai, uska portal script
website se khud dhoondhne ki koshish karta hai.
"""
import argparse
import datetime as dt
import hashlib
import html as H
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

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}
TIMEOUT = 40
MAX_DETAILS = 120          # ek company ke max kitne job descriptions kholne hain

# ===================== Filters (yahan apne hisaab se badlo) =====================
SEARCH_TERMS = ["embedded", "firmware"]
# Title mein ye ho to job seedha relevant
ROLE_RE = re.compile(
    r"embedded|firmware|\bbsp\b|autosar|\brtos\b|device driver|linux kernel|kernel|bootloader|"
    r"micro-?controller|\bmcu\b|\becu\b|j1939|\buds\b|board support|low[- ]level|\bhil\b|"
    r"\bcan\b|diagnostic|vehicle (?:electronics|software)|\bbms\b|motor control", re.I)
# Description mein ye "strong" keywords (kam se kam 2 alag) hon to bhi relevant
STRONG_RE = re.compile(
    r"firmware|\brtos\b|freertos|device drivers?|\bbsp\b|board support|bootloader|micro-?controllers?|"
    r"\bmcus?\b|autosar|\bcan[- ]?(?:bus|fd|protocol)?\b(?=[ ,/)])|j1939|\buds\b|iso[- ]?14229|embedded c\b|"
    r"embedded linux|linux kernel|kernel driver|stm32|arm cortex|cortex-m|\bi2c\b|\bspi\b|\buart\b|"
    r"\bhil\b|\becus?\b|lauterbach|jtag|\bswd\b|yocto|u-boot", re.I)
CORE_RE = re.compile(
    r"firmware|\brtos\b|freertos|device drivers?|\bbsp\b|board support|bootloader|micro-?controllers?|\bmcus?\b|"
    r"autosar|j1939|\buds\b|iso[- ]?14229|embedded c\b|embedded (?:software|systems?|linux)|linux kernel|stm32|\becus?\b|"
    r"\bcan\b(?: bus|-fd| fd|, ?lin| protocol)", re.I)
NONFW_TITLE_RE = re.compile(r"\brtl\b|design verification|\bdv\b|physical design|layout|analog|mixed[- ]signal|\bdft\b|"
                            r"\bsta\b|synthesis|mask design|\bpcb\b|hardware|mechanical|\bsmt\b|process engineer|"
                            r"data (?:engineer|scien)|cloud|devops|full[- ]?stack|frontend|front[- ]end|backend|java\b|\.net|"
                            r"salesforce|\bsap\b|android app|web\b", re.I)
ENG_TITLE_RE = re.compile(r"engineer|developer|programmer|validation|test|specialist|lead|architect|scientist|"
                          r"design|integrat|intern|trainee|graduate|technician|expert", re.I)
NOT_ROLE_RE = re.compile(r"sales|account manager|marketing|finance|\bhr\b|human resource|recruit|legal|"
                         r"procurement|purchase|buyer|accountant|payroll|logistics|warehouse|business development", re.I)
FALSE_POS_RE = re.compile(r"embedded memor(?:y|ies)|embedded (?:within|in) (?:the )?(?:customer|team|business)|"
                          r"embedded finance|embedded insurance|embedded analytics|embedded iot\.", re.I)

LOC_RE = re.compile(r"noida|gurgaon|gurugram|manesar|delhi|\bncr\b|faridabad|ghaziabad|dharuhera|bawal|dhankot|"
                    r"uttar pradesh|\bu\.?p\.?,? (?:in|india)|haryana", re.I)
VAGUE_LOC_RE = re.compile(r"\d+\s+locations|multiple|^\s*india\s*$|^\s*in\s*$|remote|anywhere|pan[- ]india", re.I)
NOIDA_RE = re.compile(r"noida|ghaziabad|delhi|uttar pradesh", re.I)
GURGAON_RE = re.compile(r"gurgaon|gurugram|manesar|haryana|faridabad|dharuhera|bawal|dhankot", re.I)

# Candidate profile -> skill weights (match score)
SKILLS = [  # (label, regex, weight)
    ("CAN", r"\bcan\b(?![ -]do)|can[- ]?bus|can[- ]?fd", 3), ("CAN-FD", r"can[- ]?fd", 1), ("J1939", r"j1939", 3),
    ("UDS", r"\buds\b|iso[- ]?14229", 3), ("AUTOSAR", r"autosar", 2), ("ECU", r"\becus?\b", 2),
    ("Vehicle diagnostics", r"diagnostic(?:s)? (?:protocol|stack|services)|obd", 2),
    ("Embedded C", r"embedded c\b|\bc programming|\bc/c\+\+|\bin c\b|\bc,\s*c\+\+", 2), ("C++", r"c\+\+", 0),
    ("RTOS", r"\brtos\b|real[- ]time operating", 2), ("FreeRTOS", r"freertos", 2),
    ("BSP", r"\bbsp\b|board support", 2), ("Device drivers", r"device drivers?|kernel drivers?|driver development", 2),
    ("Bootloader", r"bootloader|u-boot", 2), ("MCU", r"micro-?controllers?|\bmcus?\b", 2), ("STM32", r"stm32", 3),
    ("ARM Cortex", r"arm cortex|cortex-m|\barm\b", 1), ("I2C", r"\bi2c\b", 1), ("SPI", r"\bspi\b", 1),
    ("UART", r"\buart\b", 1), ("ADC", r"\badc\b", 1), ("PWM", r"\bpwm\b", 0), ("JTAG/SWD/GDB", r"jtag|\bswd\b|\bgdb\b|lauterbach|trace32|t32", 1),
    ("Embedded Linux", r"embedded linux|yocto|buildroot|linux kernel", 1), ("HIL", r"\bhil\b|hardware[- ]in[- ]the[- ]loop", 1),
    ("CANoe/CANalyzer", r"canoe|canalyzer|vector tools", 2), ("ISO 26262", r"iso[- ]?26262|functional safety", 1),
    ("ASPICE", r"a?spice", 0), ("MATLAB/Simulink", r"simulink|matlab", 0), ("BMS", r"\bbms\b|battery management", 1),
    ("Motor control", r"motor control|bldc|inverter", 1), ("BLE/Wi-Fi", r"\bble\b|bluetooth|wi-?fi", 0),
    ("Python", r"python", 0), ("FPGA", r"\bfpga\b", 0), ("EOL/calibration", r"end[- ]of[- ]line|\beol\b|calibration", 1),
]
SKILLS = [(n, re.compile(r, re.I), w) for n, r, w in SKILLS]
SENIOR_TITLE_RE = re.compile(r"\b(staff|principal|manager|director|head|lead|architect|vp|chief|senior staff|sr\.? staff)\b", re.I)
TEST_TITLE_RE = re.compile(r"test|validation|verification|qa\b|quality", re.I)
HW_TITLE_RE = re.compile(r"hardware|pcb|layout|analog|rtl|physical design|dft|\bsmt\b|prototyp", re.I)
DEV_TITLE_RE = re.compile(r"firmware|embedded|bsp|driver|autosar|software engineer|software developer|ecu|controls?", re.I)
DOMAIN_BONUS = [(re.compile(r"automotive|vehicle|\b2w\b|\bev\b|ecu|tier-1|oem", re.I), 3),
                (re.compile(r"semiconductor|mcu|chip|soc", re.I), 2),
                (re.compile(r"power|charger|inverter|energy|ev ", re.I), 1)]
PRIO_BONUS = {"High": 2, "Medium": 1, "Low": 0}


# ===================== HTTP =====================
def session():
    import requests
    from requests.adapters import HTTPAdapter
    from urllib3.util.retry import Retry
    s = requests.Session()
    s.headers.update(HEADERS)
    retry = Retry(total=2, backoff_factor=1, status_forcelist=[429, 502, 503, 504], allowed_methods=["GET", "POST"])
    s.mount("https://", HTTPAdapter(max_retries=retry))
    s.mount("http://", HTTPAdapter(max_retries=retry))
    return s


class Blocked(Exception):
    pass


def chk(r):
    if r.status_code in (401, 403, 407, 429, 451) or (r.status_code >= 400 and "captcha" in r.text[:3000].lower()):
        raise Blocked(f"HTTP {r.status_code} from {urlparse(r.url).netloc}")
    if r.status_code >= 400:
        raise Blocked(f"HTTP {r.status_code} from {r.url[:120]}")
    return r


def get_json(s, url, **kw):
    kw.setdefault("timeout", TIMEOUT)
    h = {"Accept": "application/json"}
    h.update(kw.pop("headers", {}) or {})
    return chk(s.get(url, headers=h, **kw)).json()


def post_json(s, url, body, **kw):
    kw.setdefault("timeout", TIMEOUT)
    h = {"Accept": "application/json", "Content-Type": "application/json"}
    h.update(kw.pop("headers", {}) or {})
    return chk(s.post(url, json=body, headers=h, **kw)).json()


def strip_html(t):
    t = re.sub(r"<(script|style)\b.*?</\1>", " ", t or "", flags=re.S | re.I)
    t = re.sub(r"<br\s*/?>|</p>|</li>|</div>|</h\d>", " . ", t, flags=re.I)
    return " ".join(H.unescape(re.sub(r"<[^>]+>", " ", t)).split())


def J(title, loc, url, posted="", desc="", portal="", extra_loc=""):
    return {"title": " ".join(H.unescape(title or "").split()), "location": " ".join((loc or "").split()),
            "url": url or "", "posted": str(posted or ""), "desc": desc or "", "portal": portal,
            "extra_loc": extra_loc or "", "flags": [], "exp_hint": ""}


def iso_from_ts(ts):
    try:
        ts = float(ts)
        if ts > 1e11:
            ts /= 1000
        return dt.datetime.fromtimestamp(ts, dt.timezone.utc).date().isoformat()
    except Exception:
        return ""


def ncr(text):
    return bool(LOC_RE.search(text or ""))


def eng_candidate(title):
    return bool(title) and not NOT_ROLE_RE.search(title) and (ROLE_RE.search(title) or ENG_TITLE_RE.search(title))


# ===================== Fetchers =====================
# Har fetcher list[job] return karta hai. Location/role filter baad mein central hota hai.

def f_workday(s, cfg, ctx):
    m = re.match(r"https?://([a-z0-9-]+)\.(wd\d+)\.myworkdayjobs\.com/(?:[a-z]{2}-[A-Z]{2}/)?([A-Za-z0-9_\-]+)", cfg["url"])
    if not m:
        raise ValueError("workday url format: https://<tenant>.<wdN>.myworkdayjobs.com/<site>")
    tenant, wd, site = m.groups()
    host = f"{tenant}.{wd}.myworkdayjobs.com"
    api = f"https://{host}/wday/cxs/{tenant}/{site}/jobs"
    d = post_json(s, api, {"appliedFacets": {}, "limit": 1, "offset": 0, "searchText": ""})
    acc = {}

    def walk(facets, parent=None):
        for f in facets or []:
            fp = f.get("facetParameter") or parent
            for v in f.get("values") or []:
                if "values" in v:
                    walk([v], fp)
                elif LOC_RE.search(v.get("descriptor") or "") and fp:
                    acc.setdefault(fp, []).append(v.get("id"))
    walk(d.get("facets"))
    postings = {}

    def search(term, applied, cap):
        off, total = 0, None
        while off < cap:
            dd = post_json(s, api, {"appliedFacets": applied, "limit": 20, "offset": off, "searchText": term})
            if total is None:
                total = dd.get("total") or 0          # Workday sends total only on the first page
            ps = dd.get("jobPostings") or []
            for p in ps:
                postings[p.get("externalPath")] = p
            off += 20
            if not ps or off >= total:
                break

    if acc and not ctx["any_location"]:
        fp = "locations" if "locations" in acc else max(acc, key=lambda k: len(acc[k]))
        search("", {fp: acc[fp]}, 400)                  # saari NCR jobs
        ctx["note"] = f"Workday NCR filter ({fp}): {len(postings)} NCR jobs listed"
    else:
        for t in SEARCH_TERMS:
            search(t, {}, 200)
        ctx["note"] = "Workday: no NCR location facet; keyword search (all locations)"
    out = []
    n = 0
    for path, p in postings.items():
        title, loc = p.get("title") or "", p.get("locationsText") or ""
        url = f"https://{host}/{site}{path}"
        j = J(title, loc, url, p.get("postedOn") or "", "", "Workday")
        need = eng_candidate(title) and (ncr(loc) or VAGUE_LOC_RE.search(loc) or ctx["any_location"])
        if need and n < MAX_DETAILS:
            n += 1
            try:
                info = get_json(s, f"https://{host}/wday/cxs/{tenant}/{site}{path}").get("jobPostingInfo", {})
                j["desc"] = strip_html(info.get("jobDescription"))
                j["extra_loc"] = " ; ".join([info.get("location") or ""] + list(info.get("additionalLocations") or []))
                if info.get("startDate"):
                    j["posted"] = f"{info['startDate']} ({p.get('postedOn')})" if p.get("postedOn") else info["startDate"]
                j["url"] = info.get("externalUrl") or url
            except Exception:
                pass
        out.append(j)
    return out


def f_eightfold(s, cfg, ctx):
    host, dom = cfg["host"], cfg["domain"]
    mode = cfg.get("mode", "pcsx")
    locs = cfg.get("locations") or ["Noida, Uttar Pradesh, India", "Gurugram, Haryana, India"]
    pos = {}
    for loc in (locs if not ctx["any_location"] else [""]):
        for term in ([""] if loc else SEARCH_TERMS):
            start = 0
            while start < 1500:
                if mode == "pcsx":
                    d = get_json(s, f"https://{host}/api/pcsx/search", params={"domain": dom, "query": term, "location": loc,
                                                                               "start": start, "num": 50}).get("data") or {}
                else:
                    d = get_json(s, f"https://{host}/api/apply/v2/jobs", params={"domain": dom, "query": term, "location": loc,
                                                                                 "start": start, "num": 50})
                ps = d.get("positions") or []
                for p in ps:
                    pos[p["id"]] = p
                start += len(ps)
                if not ps or start >= (d.get("count") or 0):
                    break
    out, n = [], 0
    for pid, p in pos.items():
        locs_ = p.get("locations") or [p.get("location") or ""]
        loc = " ; ".join(locs_)
        if mode == "pcsx":
            url = f"https://{host}{p.get('positionUrl')}" if p.get("positionUrl") else f"https://{host}/careers/job/{pid}"
            posted = iso_from_ts(p.get("postedTs") or p.get("creationTs"))
        else:
            url = p.get("canonicalPositionUrl") or f"https://{host}/careers/job/{pid}"
            posted = iso_from_ts(p.get("t_create"))
        j = J(p.get("name"), loc, url, posted, "", "Eightfold")
        if eng_candidate(j["title"]) and (ncr(loc) or ctx["any_location"]) and n < MAX_DETAILS:
            n += 1
            try:
                if mode == "pcsx":
                    dd = get_json(s, f"https://{host}/api/pcsx/position_details",
                                  params={"position_id": pid, "domain": dom, "hl": "en"}).get("data") or {}
                    j["desc"] = strip_html(dd.get("jobDescription"))
                else:
                    dd = get_json(s, f"https://{host}/api/apply/v2/jobs/{pid}", params={"domain": dom})
                    j["desc"] = strip_html(dd.get("job_description"))
            except Exception:
                pass
        out.append(j)
    ctx["note"] = f"Eightfold ({mode}): {len(out)} jobs at {', '.join(locs)}"
    return out


def f_smartrecruiters(s, cfg, ctx):
    cid = cfg["id"]
    out, off, n = [], 0, 0
    while off < 2000:
        d = get_json(s, f"https://api.smartrecruiters.com/v1/companies/{cid}/postings",
                     params={"limit": 100, "offset": off, **({} if ctx["any_location"] else {"country": cfg.get("country", "in")})})
        ps = d.get("content") or []
        for p in ps:
            l = p.get("location") or {}
            loc = l.get("fullLocation") or f"{l.get('city', '')}, {l.get('region', '')}, {l.get('country', '')}"
            j = J(p.get("name"), loc, f"https://jobs.smartrecruiters.com/{cid}/{p.get('id')}", (p.get("releasedDate") or "")[:10], "", "SmartRecruiters")
            if eng_candidate(j["title"]) and (ncr(loc) or ctx["any_location"]) and n < MAX_DETAILS:
                n += 1
                try:
                    dd = get_json(s, f"https://api.smartrecruiters.com/v1/companies/{cid}/postings/{p.get('id')}")
                    secs = (dd.get("jobAd") or {}).get("sections") or {}
                    j["desc"] = strip_html(" ".join((v or {}).get("text", "") for v in secs.values()))
                    j["url"] = dd.get("postingUrl") or j["url"]
                except Exception:
                    pass
            out.append(j)
        off += len(ps)
        if not ps or off >= (d.get("totalFound") or 0):
            break
    ctx["note"] = f"SmartRecruiters '{cid}': {len(out)} India jobs listed"
    return out


def f_oracle(s, cfg, ctx):
    host, site = cfg["host"], cfg["site"]
    reqs = {}
    terms = cfg.get("terms") or (SEARCH_TERMS + ["noida", "gurugram", "gurgaon"])
    if cfg.get("location_facet_id") and not ctx["any_location"]:
        off = 0                                          # poore India facet ki saari jobs
        while off < 2000:
            finder = (f"findReqs;siteNumber={site},facetsList=LOCATIONS,limit=25,offset={off},"
                      f"selectedLocationsFacet={cfg['location_facet_id']},sortBy=POSTING_DATES_DESC")
            url = (f"https://{host}/hcmRestApi/resources/latest/recruitingCEJobRequisitions"
                   f"?onlyData=true&expand=requisitionList.secondaryLocations&finder={quote(finder, safe=';=,')}")
            it = (get_json(s, url).get("items") or [{}])[0]
            rl = it.get("requisitionList") or []
            for p in rl:
                reqs[p.get("Id")] = p
            off += 25
            if not rl or off >= (it.get("TotalJobsCount") or 0):
                break
        terms = []
    for term in terms:
        finder = (f"findReqs;siteNumber={site},facetsList=LOCATIONS,limit=200,keyword=\"{term}\",sortBy=POSTING_DATES_DESC")
        url = (f"https://{host}/hcmRestApi/resources/latest/recruitingCEJobRequisitions"
               f"?onlyData=true&expand=requisitionList.secondaryLocations&finder={quote(finder, safe=';=,')}")
        for it in get_json(s, url).get("items", []):
            for p in it.get("requisitionList", []):
                reqs[p.get("Id")] = p
    out, n = [], 0
    for rid, p in reqs.items():
        loc = " ; ".join([p.get("PrimaryLocation") or ""] + [x.get("Name") or "" for x in (p.get("secondaryLocations") or [])])
        jurl = (cfg.get("public_job_url") or f"https://{host}/hcmUI/CandidateExperience/en/sites/{site}/job/{{id}}").replace("{id}", str(rid))
        j = J(p.get("Title"), loc, jurl, p.get("PostedDate"), "", "Oracle HCM")
        if eng_candidate(j["title"]) and (ncr(loc) or VAGUE_LOC_RE.search(p.get("PrimaryLocation") or "") or ctx["any_location"]) and n < MAX_DETAILS:
            n += 1
            try:
                u = (f"https://{host}/hcmRestApi/resources/latest/recruitingCEJobRequisitionDetails"
                     f"?expand=all&onlyData=true&finder=ById;Id=%22{rid}%22,siteNumber={site}")
                it = (get_json(s, u).get("items") or [{}])[0]
                j["desc"] = strip_html(" ".join(str(it.get(k) or "") for k in (
                    "ExternalDescriptionStr", "ExternalQualificationsStr", "ExternalResponsibilitiesStr")))
                j["extra_loc"] = json.dumps([it.get("workLocation"), it.get("otherWorkLocations")], ensure_ascii=False)[:800]
            except Exception:
                pass
        out.append(j)
    ctx["note"] = f"Oracle HCM: {len(out)} jobs " + (f"in location facet {cfg['location_facet_id']}" if not terms else f"for terms {terms}")
    return out


def f_lever(s, cfg, ctx):
    d = get_json(s, f"https://api.lever.co/v0/postings/{cfg['id']}?mode=json")
    out = []
    for j in d:
        desc = (j.get("descriptionPlain") or "") + " " + " ".join(
            (l.get("text") or "") + " " + strip_html(l.get("content")) for l in (j.get("lists") or []))
        out.append(J(j.get("text"), (j.get("categories") or {}).get("location"), j.get("hostedUrl"),
                     iso_from_ts(j.get("createdAt")), desc, "Lever"))
    ctx["note"] = f"Lever '{cfg['id']}': {len(out)} postings"
    return out


def f_greenhouse(s, cfg, ctx):
    d = get_json(s, f"https://boards-api.greenhouse.io/v1/boards/{cfg['id']}/jobs", params={"content": "true"})
    out = [J(j.get("title"), (j.get("location") or {}).get("name"), j.get("absolute_url"), (j.get("updated_at") or "")[:10],
             strip_html(H.unescape(j.get("content") or "")), "Greenhouse") for j in d.get("jobs", [])]
    ctx["note"] = f"Greenhouse: {len(out)} postings"
    return out


def f_talentbrew(s, cfg, ctx):
    base = cfg["url"].rstrip("/")
    jobs, page = {}, 1
    while page < 60:
        r = chk(s.get(f"{base}/search-jobs/results", params={
            "ActiveFacetID": 0, "CurrentPage": page, "RecordsPerPage": 100, "Distance": 50, "RadiusUnitType": 0,
            "Keywords": "", "Location": "", "ShowRadius": "False", "IsPagination": "True",
            "SearchResultsModuleName": "Search Results", "SearchFiltersModuleName": "Search Filters",
            "SortCriteria": 0, "SortDirection": 0, "SearchType": 5},
            headers={"X-Requested-With": "XMLHttpRequest"}, timeout=TIMEOUT))
        h = r.json().get("results", "")
        n0 = len(jobs)
        for it in re.findall(r"<li[^>]*>(.*?)</li>", h, re.S):
            m = re.search(r'href="(/job/[^"]+)"[^>]*>(.*?)</a>', it, re.S)
            if not m:
                continue
            loc = " ".join(strip_html(x) for x in re.findall(r'class="[^"]*location[^"]*"[^>]*>(.*?)</span>', it, re.S))
            title = strip_html(re.split(r"<span|<div|<p\b|<br", m.group(2))[0]) or strip_html(m.group(2)).split(" . ")[0]
            jobs[m.group(1)] = J(title, loc, base + m.group(1), "", "", "TalentBrew")
        if len(jobs) == n0:
            break
        page += 1
    out, n = [], 0
    for j in jobs.values():
        if eng_candidate(j["title"]) and (ncr(j["location"] + " " + j["url"]) or ctx["any_location"]) and n < MAX_DETAILS:
            n += 1
            try:
                t = chk(s.get(j["url"], timeout=TIMEOUT)).text
                m = re.search(r'<script type="application/ld\+json">(.*?)</script>', t, re.S)
                ld = json.loads(m.group(1)) if m else {}
                j["desc"] = strip_html(ld.get("description") or t)
                j["posted"] = (ld.get("datePosted") or "")[:10]
            except Exception:
                pass
        out.append(j)
    ctx["note"] = f"TalentBrew: {len(out)} jobs listed (all locations)"
    return out


def _avature_parse(t, base):
    out = []
    for a in re.findall(r"<article\b.*?</article>", t, re.S):
        m = re.search(r'href="([^"]*JobDetail[^"]*)"[^>]*>(.*?)</a>', a, re.S)
        if not m:
            continue
        info = strip_html(a)
        out.append(J(strip_html(m.group(2)), info, urljoin(base, H.unescape(m.group(1))), "", "", "Avature"))
    return out


def f_avature(s, cfg, ctx):
    base = cfg["url"].rstrip("/")
    per = int(cfg.get("per_page") or cfg.get("page_size") or 50)
    offp, perp = cfg.get("offset_param", "folderOffset"), cfg.get("per_page_param", "folderRecordsPerPage")
    terms = cfg.get("terms") or [""]
    jobs = {}
    for term in terms:
        off = 0
        while off < int(cfg.get("max", 3000)):
            u = f"{base}/SearchJobs/{quote(term)}" if term else f"{base}/SearchJobs/"
            prm = {perp: per, offp: off}
            if cfg.get("list_filter_mode"):
                prm["listFilterMode"] = 1
            t = chk(s.get(u, params=prm, timeout=TIMEOUT)).text
            rs = _avature_parse(t, base)
            new = 0
            for j in rs:
                if j["url"] not in jobs:
                    jobs[j["url"]] = j
                    new += 1
            off += len(rs) or per
            if not rs or not new:
                break
    out, n = [], 0
    for j in jobs.values():
        info = j["location"]
        rest = info[len(j["title"]):] if info.startswith(j["title"]) else info
        j["location"] = re.split(r"\s*(?:•|\bJob ID\b|\bShare\b|Learn more)", rest.strip(" ."))[0].strip(" .") or rest[:80]
        if eng_candidate(j["title"]) and (ncr(info) or VAGUE_LOC_RE.search(info) or ctx["any_location"]) \
                and (ROLE_RE.search(j["title"]) or ncr(info)) and n < MAX_DETAILS:
            n += 1
            try:
                t = chk(s.get(j["url"], timeout=TIMEOUT)).text
                j["desc"] = strip_html(t)[:20000]
                m = re.search(r"(?:Posted|Posting date|Date posted)[^0-9A-Za-z]{0,20}([A-Z][a-z]{2,8} \d{1,2}, 20\d\d|\d{1,2}[-/][A-Za-z0-9]{1,3}[-/]20\d\d|20\d\d-\d\d-\d\d)", j["desc"])
                j["posted"] = m.group(1) if m else ""
                m = re.search(r"Posted since\s*(\d{1,2}-[A-Za-z]{3}-20\d\d)", j["desc"])
                if m:
                    j["posted"] = m.group(1)
                if cfg.get("company_filter") and cfg["company_filter"].lower() not in j["desc"].lower():
                    continue
            except Exception:
                pass
        out.append(j)
    ctx["note"] = f"Avature: {len(out)} jobs listed for terms {terms}"
    return out


def f_sf_csb(s, cfg, ctx):
    host = cfg["host"]
    qs = cfg.get("terms") or [""]
    jobs = {}
    for q in qs:
        start = 0
        while start < 2000:
            t = chk(s.get(f"https://{host}/search/", params={"q": q, "locationsearch": cfg.get("locationsearch", "India"),
                                                             "startrow": start}, timeout=TIMEOUT)).text
            rows = re.findall(r'<tr class="data-row">.*?</tr>', t, re.S)
            for row in rows:
                u = re.search(r'href="(/job/[^"]+)"', row)
                if not u:
                    continue
                title = strip_html((re.search(r'class="jobTitle-link[^"]*"[^>]*>(.*?)</a>', row, re.S) or re.search(r">(.*?)</a>", row, re.S)).group(1))
                loc = strip_html((re.search(r'class="jobLocation[^"]*"[^>]*>(.*?)</span>', row, re.S) or re.search(r"()", "")).group(1))
                date = strip_html((re.search(r'class="jobDate[^"]*"[^>]*>(.*?)</span>', row, re.S) or re.search(r"()", "")).group(1))
                jobs[u.group(1)] = J(title, loc, f"https://{host}" + H.unescape(u.group(1)), date, "", "SuccessFactors")
            m = re.search(r"Results\s*<b>[^<]*</b>\s*of\s*<b>([^<]*)</b>", t)
            total = int(re.sub(r"\D", "", m.group(1)) or 0) if m else 0
            start += len(rows)
            if not rows or start >= total:
                break
    out, n = [], 0
    for j in jobs.values():
        if eng_candidate(j["title"]) and (ncr(j["location"]) or VAGUE_LOC_RE.search(j["location"]) or not j["location"] or ctx["any_location"]) and n < MAX_DETAILS:
            n += 1
            try:
                t = chk(s.get(j["url"], timeout=TIMEOUT)).text
                m = re.search(r'<span class="jobdescription">(.*)', t, re.S)
                j["desc"] = strip_html(m.group(1) if m else t)[:20000]
                d = re.search(r'itemprop="datePosted"[^>]*content="([^"]+)"', t)
                if d:
                    j["posted"] = d.group(1)[:10]
                g = re.search(r'class="jobGeoLocation"[^>]*>(.*?)<', t, re.S)
                if g:
                    j["extra_loc"] = strip_html(g.group(1))
            except Exception:
                pass
        out.append(j)
    ctx["note"] = f"SuccessFactors CSB: {len(out)} jobs (locationsearch={cfg.get('locationsearch', 'India')})"
    return out


def f_jibe(s, cfg, ctx):
    host = cfg["host"]
    out, seen = [], set()
    for kw in (cfg.get("terms") or [""]):
        page, got = 1, 0
        while page < 40:
            p = {"page": page, "limit": 100}
            if not ctx["any_location"]:
                p["location"] = cfg.get("location", "India")
            if kw:
                p["keywords"] = kw
            d = get_json(s, f"https://{host}/api/jobs", params=p)
            js = d.get("jobs") or []
            for x in js:
                x = x.get("data") or {}
                rid = x.get("req_id") or x.get("slug")
                if rid in seen:
                    continue
                seen.add(rid)
                locs = " ; ".join([x.get("full_location") or x.get("location_name") or ""] +
                                  [str(l) for l in (x.get("additional_locations") or x.get("tags9") or [])])
                url = (x.get("meta_data") or {}).get("canonical_url") or f"https://{host}/jobs/{rid}"
                out.append(J(x.get("title"), locs, url, (x.get("posted_date") or "")[:10],
                             strip_html(x.get("description") or "") + " " + strip_html(x.get("qualifications") or ""), "Jibe"))
            got += len(js)
            page += 1
            if not js or got >= (d.get("totalCount") or 0):
                break
    ctx["note"] = f"Jibe: {len(out)} jobs (location={cfg.get('location', 'India')})"
    return out


def f_csod(s, cfg, ctx):
    corp, site = cfg["corp"], int(cfg["site"])
    t = chk(s.get(f"https://{corp}.csod.com/ux/ats/careersite/{site}/home?c={corp}", timeout=TIMEOUT)).text
    tok = re.search(r'"token":"([^"]+)"', t)
    cloud = re.search(r'"cloud":"([^"]+)"', t)
    if not tok or not cloud:
        raise Blocked("Cornerstone page did not expose its public search token")
    hd = {"Authorization": f"Bearer {tok.group(1)}", "Origin": f"https://{corp}.csod.com", "Referer": f"https://{corp}.csod.com/"}
    out, page = [], 1
    while page < 30:
        body = {"careerSiteId": site, "careerSitePageId": site, "pageNumber": page, "pageSize": 100, "cultureId": 1,
                "searchText": "", "cultureName": "en-US", "states": [], "countryCodes": [], "cities": [], "placeID": "",
                "radius": None, "postingsWithinDays": None, "customFieldCheckboxKeys": [], "customFieldDropdowns": [],
                "customFieldRadios": []}
        d = (post_json(s, f"{cloud.group(1)}rec-job-search/external/jobs", body, headers=hd).get("data") or {})
        reqs = d.get("requisitions") or []
        for q in reqs:
            ls = "; ".join(", ".join(str(l.get(k) or "") for k in ("city", "state", "country")) for l in (q.get("locations") or []))
            out.append(J(q.get("displayJobTitle"), ls,
                         f"https://{corp}.csod.com/ux/ats/careersite/{site}/home/requisition/{q.get('requisitionId')}?c={corp}",
                         (q.get("postingEffectiveDate") or "")[:10], q.get("externalDescription") or "", "Cornerstone"))
        if not reqs or page * 100 >= (d.get("totalCount") or 0):
            break
        page += 1
    n = 0
    for j in out:
        if eng_candidate(j["title"]) and (ncr(j["location"]) or ctx["any_location"]) and n < MAX_DETAILS:
            n += 1
            rid = re.search(r"/requisition/(\d+)", j["url"]).group(1)
            try:
                dd = get_json(s, f"https://{corp}.csod.com/services/x/job-requisition/v2/requisitions/{rid}/jobDetails",
                              params={"cultureId": 1}, headers=hd).get("data") or {}
                if isinstance(dd, list):
                    dd = dd[0] if dd else {}
                j["desc"] = strip_html(dd.get("externalDescription") or "")
            except Exception:
                pass
    ctx["note"] = f"Cornerstone: {len(out)} jobs listed worldwide"
    return out


def _dig(obj, path):
    if not path:
        return obj
    if "*" in str(path).split("."):
        head, _, tail = str(path).partition("*")
        base = _dig(obj, head.rstrip("."))
        vals = list(base.values()) if isinstance(base, dict) else (base or [])
        res = []
        for v in vals:
            x = _dig(v, tail.lstrip("."))
            res.extend(x if isinstance(x, list) else ([x] if x is not None else []))
        return res
    for k in str(path).split("."):
        if obj is None:
            return None
        if isinstance(obj, list):
            try:
                obj = obj[int(k)]
            except (ValueError, IndexError):
                return None
        else:
            obj = obj.get(k) if isinstance(obj, dict) else None
    return obj


def _fill(v, term, page, offset):
    if isinstance(v, str):
        return v.replace("{term}", term).replace("{page}", str(page)).replace("{offset}", str(offset))
    if isinstance(v, dict):
        return {k: _fill(x, term, page, offset) for k, x in v.items()}
    if isinstance(v, list):
        return [_fill(x, term, page, offset) for x in v]
    return v


def _num_fill(v):
    """'{page}' jaise placeholders ke baad numeric strings ko int bana do (JSON body ke liye)."""
    if isinstance(v, str) and re.fullmatch(r"-?\d+", v):
        return int(v)
    if isinstance(v, dict):
        return {k: _num_fill(x) for k, x in v.items()}
    if isinstance(v, list):
        return [_num_fill(x) for x in v]
    return v


def f_json_api(s, cfg, ctx):
    pg = cfg.get("page") or {}
    start, step, mx = int(pg.get("start", 1)), int(pg.get("step", 1)), int(pg.get("max", 1 if not pg else 20))
    flds = cfg.get("fields") or {}
    out, seen = [], set()
    for term in (cfg.get("terms") or [""]):
        for i in range(mx):
            page = start + i * step
            off = page
            url = _fill(cfg["url"], term, page, off)
            params = _fill(cfg.get("params") or {}, term, page, off)
            hd = cfg.get("headers") or {}
            if (cfg.get("method") or "GET").upper() == "POST" and cfg.get("data") is not None:
                frm = _fill(cfg.get("data") or {}, term, page, off)
                d = chk(s.post(url, data=frm, params=params, headers={"Accept": "application/json", **hd}, timeout=TIMEOUT)).json()
            elif (cfg.get("method") or "GET").upper() == "POST":
                body = _num_fill(_fill(cfg.get("json") or {}, term, page, off))
                d = post_json(s, url, body, params=params, headers=hd)
            else:
                d = get_json(s, url, params=params, headers=hd)
            items = _dig(d, cfg.get("items")) or []
            if isinstance(items, dict):
                items = list(items.values())
            new = 0
            for it in items:
                def g(k):
                    if isinstance(flds.get(k), list):
                        vals = []
                        for path in flds[k]:
                            v = _dig(it, path)
                            v = strip_html(str(v)) if v not in (None, "", []) else ""
                            if v and v not in vals:
                                vals.append(v)
                        return " – ".join(vals)
                    v = _dig(it, flds.get(k)) if flds.get(k) else None
                    if isinstance(v, list) and len(v) >= 3 and all(isinstance(x, int) for x in v[:3]) and v[0] > 1900:
                        return f"{v[0]:04d}-{v[1]:02d}-{v[2]:02d}"
                    if isinstance(v, list) and all(isinstance(x, str) for x in v):
                        return " ; ".join(v)
                    return "" if v is None else (json.dumps(v, ensure_ascii=False) if isinstance(v, (dict, list)) else str(v))
                url_ = g("url")
                if not url_ and cfg.get("url_template"):
                    url_ = re.sub(r"\{([\w.]+)\}", lambda m: str(_dig(it, m.group(1)) or ""), cfg["url_template"])
                key = url_ or g("title")
                if key in seen:
                    continue
                seen.add(key)
                new += 1
                posted = g("posted")
                if re.fullmatch(r"\d{10,13}", posted or ""):
                    posted = iso_from_ts(posted)
                jj = J(g("title"), g("location") or cfg.get("default_location", ""), url_, posted[:25],
                       strip_html(H.unescape(g("desc"))), cfg.get("label", "JSON API"))
                jj["exp_hint"] = g("experience")
                if flds.get("exp_min"):
                    try:
                        div = float(cfg.get("exp_divisor") or 1)
                        a, b = float(g("exp_min")) / div, (float(g("exp_max")) / div if g("exp_max") else None)
                        jj["exp_hint"] = f"{a:g}-{b:g} years" if b is not None else f"{a:g}+ years"
                    except ValueError:
                        pass
                out.append(jj)
            if not items or not new or not pg or _dig(d, pg.get("last_flag") or "__none__") is True:
                break
    ctx["note"] = f"{cfg.get('label', 'JSON API')}: {len(out)} jobs listed"
    return out


def f_html_list(s, cfg, ctx):
    pat = re.compile(cfg.get("link_pattern") or r"job|career|opening|vacanc", re.I)
    item_rx = re.compile(cfg["item_regex"], re.S | re.I) if cfg.get("item_regex") else None
    out, seen = [], set()
    for u in cfg.get("urls") or [cfg.get("url")]:
        r = s.get(u, timeout=TIMEOUT)
        if r.status_code == 404 and out:
            break
        t = chk(r).text
        if item_rx:
            for m in item_rx.finditer(t):
                gd = m.groupdict()
                title = strip_html(gd.get("title") or "")
                if title and title not in seen:
                    seen.add(title)
                    out.append(J(title, strip_html(gd.get("location") or "") or cfg.get("default_location", ""),
                                 urljoin(r.url, gd.get("url") or "") if gd.get("url") else r.url,
                                 strip_html(gd.get("posted") or ""), "", "Careers page"))
            continue
        for href, text in A_RE.findall(t):
            full = urljoin(r.url, H.unescape(href.strip()))
            title = strip_html(text)
            if not title and pat.search(full):
                title = re.sub(r"[-_]+", " ", urlparse(full).path.rstrip("/").split("/")[-1]).strip().title()
            if not title or len(title) > 160 or not pat.search(full) or full in seen:
                continue
            if cfg.get("title_filter", True) and not (ROLE_RE.search(title) or ENG_TITLE_RE.search(title)):
                continue
            seen.add(full)
            out.append(J(title, cfg.get("default_location", ""), full, "", "", "Careers page"))
    n = 0
    if cfg.get("fetch_details", True) and not item_rx:
        for j in out:
            if n >= MAX_DETAILS or not eng_candidate(j["title"]):
                continue
            n += 1
            try:
                t = chk(s.get(j["url"], timeout=TIMEOUT)).text
                j["desc"] = strip_html(t)[:20000]
                m = re.search(r'"datePosted"\s*:\s*"([^"]+)"|datetime="(20\d\d-\d\d-\d\d)', t)
                if m:
                    j["posted"] = (m.group(1) or m.group(2))[:10]
                lm = re.search(r"(?:Location|Job Location)\s*[:\-]\s*([A-Za-z ,\-]{3,60})", j["desc"])
                if lm and ncr(lm.group(1)):
                    j["location"] = lm.group(1).strip()
            except Exception:
                pass
    ctx["note"] = f"HTML page scan: {len(out)} engineering links"
    return out


A_RE = re.compile(r"<a\b[^>]*?href=[\"']([^\"'#]+)[\"'][^>]*>(.*?)</a>", re.I | re.S)
TAG_RE = re.compile(r"<[^>]+>")


def f_generic(s, cfg, ctx):
    return f_html_list(s, {"urls": [cfg["url"]], "link_pattern": r".", "fetch_details": True,
                           "default_location": "", "title_filter": True}, ctx)


def _gentle_get_json(s, url, params=None, gap=2.5):
    """Darwinbox jaise portals jaldi-jaldi requests pe 403 dete hain: har call ke beech gap, 403 pe ek baar ruk ke retry."""
    time.sleep(gap)
    try:
        return get_json(s, url, params=params)
    except Blocked as e:
        if "403" not in str(e) and "429" not in str(e):
            raise
        time.sleep(20)
        return get_json(s, url, params=params)


def f_darwinbox(s, cfg, ctx):
    host = cfg["host"]
    out, page = [], 1
    while page < 30:
        d = _gentle_get_json(s, f"https://{host}/ms/candidateapi/job", params={"page": page, "limit": 50}).get("message") or {}
        js = d.get("jobs") or []
        for x in js:
            j = J(x.get("title") or x.get("designation_display_name"), x.get("officelocation_show_arr") or x.get("tool_tip_locations"),
                  f"https://{host}/ms/candidate/careers/{x.get('id')}",
                  iso_from_ts(x.get("job_posting_on")) or (x.get("created_on") or "")[:10], "", "Darwinbox")
            ef, et = x.get("experience_from_num"), x.get("experience_to_num")
            j["exp_hint"] = f"{ef}-{et} years" if ef not in (None, "") and et not in (None, "") else ""
            if eng_candidate(j["title"]) and (ncr(j["location"]) or ctx["any_location"]):
                try:
                    dj = (_gentle_get_json(s, f"https://{host}/ms/candidateapi/job/{x.get('id')}").get("message") or {}).get("job") or [{}]
                    j["desc"] = strip_html(H.unescape(dj[0].get("jd") or ""))
                    j["exp_hint"] = dj[0].get("experience") or j["exp_hint"]
                except Exception:
                    pass
            out.append(j)
        page += 1
        if not js or len(out) >= int(d.get("jobscount") or 0):
            break
    ctx["note"] = f"Darwinbox: {len(out)} jobs listed"
    return out


def f_sf_classic_xml(s, cfg, ctx):
    t = chk(s.get(cfg["url"], timeout=TIMEOUT)).text
    out = []
    for blk in re.findall(r"<Job>(.*?)</Job>", t, re.S):
        def tag(n):
            m = re.search(rf"<{n}>(.*?)</{n}>", blk, re.S)
            return re.sub(r"^<!\[CDATA\[|\]\]>$", "", m.group(1).strip()) if m else ""
        rid = tag("ReqId")
        loc = re.search(r"<label>\s*Job Location\s*</label>\s*<value>(.*?)</value>", blk, re.S)
        url = (cfg.get("url_template") or "").replace("{ReqId}", rid)
        out.append(J(strip_html(tag("JobTitle")), strip_html(loc.group(1)) if loc else cfg.get("default_location", ""), url,
                     strip_html(tag("Posted-Date") or tag("PostedDate")), strip_html(H.unescape(tag("Job-Description"))),
                     "SuccessFactors (classic)"))
    ctx["note"] = f"SuccessFactors classic XML feed: {len(out)} jobs"
    return out


def f_manual(s, cfg, ctx):
    ctx["note"] = cfg.get("note") or "Manual check needed"
    ctx["status"] = cfg.get("status") or "Portal not found"
    return []


FETCHERS = {"workday": f_workday, "eightfold": f_eightfold, "smartrecruiters": f_smartrecruiters,
            "oracle": f_oracle, "lever": f_lever, "greenhouse": f_greenhouse, "talentbrew": f_talentbrew,
            "avature": f_avature, "sf_csb": f_sf_csb, "jibe": f_jibe, "csod": f_csod, "json_api": f_json_api,
            "html_list": f_html_list, "generic": f_generic, "manual": f_manual,
            "darwinbox": f_darwinbox, "sf_classic_xml": f_sf_classic_xml}
TYPE_LABEL = {"workday": "Workday", "eightfold": "Eightfold", "smartrecruiters": "SmartRecruiters", "oracle": "Oracle HCM",
              "lever": "Lever", "greenhouse": "Greenhouse", "talentbrew": "Radancy TalentBrew", "avature": "Avature",
              "sf_csb": "SAP SuccessFactors", "jibe": "iCIMS Jibe", "csod": "Cornerstone OnDemand", "json_api": "Custom (JSON API)",
              "html_list": "Custom careers page", "generic": "Custom careers page (auto-found)", "manual": "-",
              "darwinbox": "Darwinbox", "sf_classic_xml": "SAP SuccessFactors (classic)"}


def portal_url(cfg):
    if cfg.get("url"):
        return cfg["url"]
    if cfg.get("urls"):
        return cfg["urls"][0]
    t = cfg.get("type")
    if t == "eightfold":
        return f"https://{cfg['host']}/careers"
    if t == "smartrecruiters":
        return f"https://jobs.smartrecruiters.com/{cfg['id']}"
    if t == "oracle":
        return f"https://{cfg['host']}/hcmUI/CandidateExperience/en/sites/{cfg['site']}"
    if t == "lever":
        return f"https://jobs.lever.co/{cfg['id']}"
    if t == "greenhouse":
        return f"https://boards.greenhouse.io/{cfg['id']}"
    if t in ("sf_csb", "jibe"):
        return f"https://{cfg['host']}/"
    if t == "darwinbox":
        return f"https://{cfg['host']}/ms/candidate/careers"
    if t == "csod":
        return f"https://{cfg['corp']}.csod.com/ux/ats/careersite/{cfg['site']}/home?c={cfg['corp']}"
    return ""


# ===================== Auto-discovery (company yaml mein na ho to) =====================
PATTERNS = [
    ("workday", re.compile(r"https?://[a-z0-9-]+\.wd\d+\.myworkdayjobs\.com/(?:[a-z]{2}-[A-Z]{2}/)?[A-Za-z0-9_\-]+")),
    ("greenhouse", re.compile(r"https?://(?:boards|job-boards)(?:\.eu)?\.greenhouse\.io/(?:embed/job_board(?:/js)?\?for=)?([A-Za-z0-9_\-]+)")),
    ("lever", re.compile(r"https?://jobs\.lever\.co/([A-Za-z0-9_\-]+)")),
    ("smartrecruiters", re.compile(r"https?://(?:jobs|careers)\.smartrecruiters\.com/([A-Za-z0-9_\-]+)")),
    ("oracle", re.compile(r"https?://([a-z0-9.-]+\.oraclecloud\.com)/hcmUI/CandidateExperience/[a-z]{2}(?:-[A-Z]{2})?/sites/([A-Za-z0-9_]+)")),
]


def detect(url, html=""):
    blob = (url or "") + "\n" + (html or "")
    for typ, rx in PATTERNS:
        for m in rx.finditer(blob):
            if typ == "workday":
                if re.search(r"/(wday|en-US|job)$", m.group(0)):
                    continue
                return {"type": "workday", "url": m.group(0)}
            if typ == "oracle":
                return {"type": "oracle", "host": m.group(1), "site": m.group(2)}
            if m.group(1).lower() in ("embed", "api", "v1"):
                continue
            return {"type": typ, "id": m.group(1)}
    return None


def discover(s, website):
    w = re.sub(r"^https?://", "", (website or "").strip().strip("/"))
    if not w:
        return None
    host, _, path = w.partition("/")
    base = f"https://{host}" + (f"/{path}" if path else "")
    bare = host[4:] if host.startswith("www.") else host
    first_ok = None
    for u in [f"{base}/careers", f"https://www.{bare}/careers", f"https://careers.{bare}", f"https://jobs.{bare}",
              f"{base}/career", f"{base}/en/careers", f"{base}/careers.html"]:
        try:
            r = s.get(u, timeout=20, allow_redirects=True)
        except Exception:
            continue
        if r.status_code >= 400 or "html" not in r.headers.get("Content-Type", "html"):
            continue
        hit = detect(r.url, r.text)
        if hit:
            return hit
        first_ok = first_ok or r.url
    return {"type": "generic", "url": first_ok} if first_ok else None


# ===================== Analysis =====================
EXP_RE = re.compile(r"(?<![\d.])(\d{1,2})\s*(?:\+|plus)?\s*(?:(?:-|–|to)\s*(\d{1,2})\s*\+?)?\s*(?:years?|yrs?)(?!\s*old)", re.I)


def _exp_hint(hint):
    m = re.search(r"(\d{1,2}(?:\.\d)?)\s*(?:(?:-|–|to)\s*(\d{1,2}(?:\.\d)?)|\+)\s*(?:years|yrs)", hint or "", re.I)
    if m:
        return float(m.group(1)), (f"{m.group(1)}-{m.group(2)} years" if m.group(2) else f"{m.group(1)}+ years")
    return None, "Not stated"


def experience(desc, hint=""):
    if hint and re.search(r"\d", hint):
        r = _exp_hint(hint)
        if r[0] is not None:
            return r
    for m in EXP_RE.finditer(desc or ""):
        lo, hi = int(m.group(1)), m.group(2)
        pre = desc[max(0, m.start() - 45):m.start()].lower()
        post = desc[m.end():m.end() + 45].lower()
        if lo > 25 or re.search(r"\b(?:over|for|since|founded|nearly|almost|than)\s*$", pre) or \
                re.search(r"^\W*(?:of (?:computational|history|innovation|legacy|excellence)|old|ago|warranty)", post):
            continue
        if not (re.search(r"experience|exp\b|exp\.|exp:|minimum|min\.|at least|having|with\s*$|and\s*$|\+\s*$", pre)
                or re.search(r"^\W*(?:of\s+)?(?:[\w/&-]+\s+){0,6}?(?:experience|exp\b|work|industry|hands-on|relevant)", post)):
            continue
        return lo, (f"{lo}-{hi} years" if hi else f"{lo}+ years")
    return _exp_hint(hint)


def analyze(j, row):
    title = j["title"]
    m = re.match(r"(20\d\d-\d\d-\d\d)[T ]\d\d:\d\d[^ (]*\s*(.*)", j.get("posted") or "")
    if m:
        j["posted"] = (m.group(1) + " " + m.group(2)).strip()
    desc = FALSE_POS_RE.sub(" ", j.get("desc") or "")
    blob = title + " . " + desc
    found = [(n, w) for n, rx, w in SKILLS if rx.search(blob)]
    j["skills"] = ", ".join(n for n, _ in found[:14]) or "(description not read - open link)"
    lo, exp_txt = experience(desc, j.get("exp_hint", ""))
    j["experience"] = exp_txt
    # ---- score ----
    sc = sum(w for _, w in found)
    sc = min(sc, 14)
    if DEV_TITLE_RE.search(title) and not TEST_TITLE_RE.search(title):
        sc += 3
    if TEST_TITLE_RE.search(title):
        sc -= 2
    if HW_TITLE_RE.search(title) and not re.search(r"firmware|embedded|software", title, re.I):
        sc -= 4
    sector = f"{row.get('Sector', '')} {row.get('Embedded / firmware work', '')}"
    for rx, b in DOMAIN_BONUS:
        if rx.search(sector):
            sc += b
            break
    sc += PRIO_BONUS.get(row.get("Match for Aman"), 0)
    hard_senior = bool(re.search(r"\b(staff|principal|manager|director|head|architect|chief|vp|agm|dgm|gm)\b", title, re.I))
    if lo is not None:
        senior = lo >= 5 or (hard_senior and lo >= 4)
        if lo <= 3:
            sc += 3
        elif lo < 5:
            sc += 0
        else:
            sc -= 5
    else:
        senior = bool(SENIOR_TITLE_RE.search(title))
        if senior:
            sc -= 3
        elif re.search(r"\b(senior|sr\.?)\b", title, re.I):
            sc -= 2
    if senior:
        j["flags"].append("Senior – stretch")
    if not j.get("desc"):
        j["flags"].append("Description not read - check link")
    j["score"] = sc
    j["match"] = "High" if sc >= 14 else ("Med" if sc >= 8 else "Low")
    if senior and j["match"] == "High":
        j["match"] = "Med"
    top = [n for n, w in sorted(found, key=lambda x: -x[1]) if w >= 2][:4]
    bits = []
    if top:
        bits.append("matches your " + ", ".join(top))
    elif found:
        bits.append("mentions " + ", ".join(n for n, _ in found[:3]))
    for rx, b in DOMAIN_BONUS:
        if rx.search(sector):
            bits.append(["automotive", "semiconductor", "power/EV"][DOMAIN_BONUS.index((rx, b))] + " domain")
            break
    if lo is not None:
        bits.append(f"asks {exp_txt}" + (" (fits ~2 yrs)" if lo <= 3 else ""))
    if TEST_TITLE_RE.search(title):
        bits.append("testing/validation role")
    j["why"] = ("; ".join(bits) or "embedded keyword in title").capitalize()


def relevant(j):
    title = j["title"]
    if NOT_ROLE_RE.search(title) or re.search(r"embedded memor", title, re.I):
        return False
    title_hit = ROLE_RE.search(title) and not re.fullmatch(r"\s*(?:can|diagnostic\w*)\s*", ROLE_RE.search(title).group(0), re.I)
    if title_hit and re.search(r"embedded|firmware|\bbsp\b|autosar|rtos|driver|kernel|bootloader|micro|\bmcu\b|\becu\b|"
                               r"j1939|\buds\b|board support|low[- ]level|\bhil\b|\bbms\b", title, re.I):
        return True
    if NONFW_TITLE_RE.search(title) and not title_hit:
        return False
    d = FALSE_POS_RE.sub(" ", j.get("desc") or "")
    strong = set(m.group(0).lower() for m in STRONG_RE.finditer(d))
    core = set(m.group(0).lower() for m in CORE_RE.finditer(d))
    if title_hit:                                   # e.g. "CAN" / "Diagnostics" in title: need some JD support
        return len(core) >= 1 or not d
    if ENG_TITLE_RE.search(title):
        return len(core) >= 2 and len(strong | core) >= 3
    return len(core) >= 3 and len(strong | core) >= 4   # titles like "Cybersecurity, S/W Electronic control system"


OTHER_CITY_RE = re.compile(r"bengaluru|bangalore|pune|hyderabad|chennai|mumbai|kolkata|ahmedabad|coimbatore|kochi|"
                           r"trivandrum|mysuru|mysore|vadodara|nagpur|jaipur|indore|chandigarh|mohali|sanand|tirupati", re.I)


def located(j, any_location):
    if any_location:
        return True
    if ncr(j["location"]):
        return True
    if ncr(j.get("extra_loc", "")):
        if VAGUE_LOC_RE.search(j["location"]) or not j["location"].strip():
            cities = sorted(set(m.group(0).title() for m in LOC_RE.finditer(j["extra_loc"])
                                if m.group(0).lower() not in ("uttar pradesh", "haryana")))
            j["location"] = f"{j['location']} (incl. {', '.join(cities) or 'NCR'})".strip()
        return True
    if not j["location"].strip() or VAGUE_LOC_RE.search(j["location"]):
        # User rule: "Multiple locations"/unclear -> keep and mark "Location unclear",
        # unless the JD clearly places the job somewhere else (another Indian city, or not India at all).
        d = j.get("desc", "")
        if OTHER_CITY_RE.search(j.get("extra_loc", "")):    # portal's own detail names another city
            return False
        if not d or ncr(d[:6000]):
            j["flags"].append("Location unclear")
            return True
        india_hint = re.search(r"\bindia\b", j["location"] + " " + d[:6000], re.I)
        if india_hint and not OTHER_CITY_RE.search(j["title"] + " " + d[:6000]):
            j["flags"].append("Location unclear")
            return True
    return False


def city_of(j, rows):
    """Do rows (Noida + Gurgaon) wali company mein job kis row pe jaaye."""
    if len(rows) == 1:
        return rows[0]
    loc = f"{j['location']} {j.get('extra_loc', '')}"
    for r in rows:
        if r["City"].lower().startswith("gur") and GURGAON_RE.search(loc):
            return r
        if r["City"].lower().startswith("noi") and NOIDA_RE.search(loc):
            return r
    return rows[0]


# ===================== IO =====================
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
    hdr = [str(h or "").strip() for h in rows[0]]
    out = []
    for r in rows[1:]:
        d = {hdr[i]: (r[i] if i < len(r) else None) for i in range(len(hdr))}
        if d.get("Company"):
            d["Company"] = str(d["Company"]).strip()
            d["City"] = str(d.get("City") or "").strip()
            d["Website"] = str(d.get("Website") or "").strip()
            d["Match for Aman"] = str(d.get("Match for Aman") or "").strip()
            out.append(d)
    return out


# ===================== Excel =====================
def write_excel(jobs, status, companies, path, top5, run_time):
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
    from openpyxl.utils import get_column_letter
    F = "Arial"
    hdr_font, hdr_fill = Font(name=F, bold=True, color="FFFFFF", size=10), PatternFill("solid", fgColor="1F3864")
    body, bold = Font(name=F, size=10), Font(name=F, size=10, bold=True)
    link = Font(name=F, size=10, color="0563C1", underline="single")
    note = Font(name=F, italic=True, size=9, color="595959")
    thin = Side(style="thin", color="D9D9D9")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(vertical="top", wrap_text=True)
    fills = {"High": "C6EFCE", "Med": "FFEB9C", "Low": "EDEDED", "Jobs found": "C6EFCE", "No matching jobs": "EDEDED",
             "Blocked": "FFC7CE", "Portal not found": "FCE4D6", "NEW": "C6EFCE"}

    def header(ws, cols, widths):
        ws.append(cols)
        for i, c in enumerate(ws[1], 1):
            c.font, c.fill, c.border = hdr_font, hdr_fill, border
            c.alignment = Alignment(vertical="center", wrap_text=True)
            ws.column_dimensions[get_column_letter(i)].width = widths[i - 1]
        ws.row_dimensions[1].height = 30
        ws.freeze_panes = "A2"

    wb = Workbook()
    ws = wb.active
    ws.title = "Jobs"
    header(ws, ["Company", "City", "Job title", "Location", "Experience asked", "Posted date", "Key skills asked",
                "Match (High/Med/Low)", "Why it matches (one line)", "Direct apply link", "New since last run?"],
           [26, 9, 36, 24, 20, 22, 44, 11, 46, 52, 10])
    for j in jobs:
        loc = j["location"] + (" [Location unclear]" if "Location unclear" in j["flags"] else "")
        exp = j["experience"] + (" [Senior – stretch]" if "Senior – stretch" in j["flags"] else "")
        why = j["why"] + ("  ⚠ Description not read - check link" if "Description not read - check link" in j["flags"] else "")
        ws.append([j["company"], j["city"], j["title"], loc, exp, j["posted"] or "Not shown", j["skills"], j["match"], why,
                   j["url"], "NEW" if j.get("new") else ""])
        r = ws.max_row
        for c in ws[r]:
            c.font, c.alignment, c.border = body, wrap, border
        if j["url"]:
            ws.cell(r, 10).hyperlink = j["url"]
            ws.cell(r, 10).font = link
        ws.cell(r, 8).fill = PatternFill("solid", fgColor=fills[j["match"]])
        if j.get("new"):
            ws.cell(r, 11).fill = PatternFill("solid", fgColor=fills["NEW"])
    nj = max(ws.max_row, 2)
    ws.auto_filter.ref = f"A1:K{nj}"

    ps = wb.create_sheet("Portal Status")
    header(ps, ["Company", "City", "Careers portal URL", "Portal type", "Status", "# matching jobs", "Notes (what was checked)"],
           [34, 9, 52, 24, 18, 11, 70])
    for st in status:
        ps.append([st["company"], st["city"], st["url"], st["type"], st["status"], None, st["note"]])
        i = ps.max_row
        ps.cell(i, 6).value = f"=COUNTIFS(Jobs!$A$2:$A$2000,A{i},Jobs!$B$2:$B$2000,B{i})"
        for c in ps[i]:
            c.font, c.alignment, c.border = body, wrap, border
        ps.cell(i, 6).alignment = Alignment(horizontal="center", vertical="top")
        if st["url"].startswith("http"):
            ps.cell(i, 3).hyperlink = st["url"].split(" ")[0]
            ps.cell(i, 3).font = link
        ps.cell(i, 5).fill = PatternFill("solid", fgColor=fills.get(st["status"], "FFFFFF"))
    np_ = max(ps.max_row, 2)
    ps.auto_filter.ref = f"A1:G{np_}"

    sm = wb.create_sheet("Summary")
    for col, w in zip("ABC", [44, 18, 90]):
        sm.column_dimensions[col].width = w
    sm["A1"] = f"Embedded / firmware jobs – Delhi NCR – run {run_time}"
    sm["A1"].font = Font(name=F, bold=True, size=13, color="1F3864")
    lines = [
        ("Coverage", None, None),
        ("Rows in input file", len(companies), "Rows in the 'Companies' sheet of the input Excel."),
        ("Rows checked (Portal Status)", f"=COUNTA('Portal Status'!A2:A{np_})", None),
        ("Checked vs input", '=IF(B4=B3,"MATCH – every input row checked","MISMATCH")', None),
        ("Unique companies in input", len({c["Company"].split(" (")[0].lower() for c in companies}), "Same company in two cities counted once (by name before brackets)."),
        ("Status breakdown (rows)", None, None),
        ("Jobs found", f"=COUNTIF('Portal Status'!E2:E{np_},\"Jobs found\")", None),
        ("No matching jobs", f"=COUNTIF('Portal Status'!E2:E{np_},\"No matching jobs\")", None),
        ("Blocked", f"=COUNTIF('Portal Status'!E2:E{np_},\"Blocked\")", "Portal refused/failed - open these in a browser."),
        ("Portal not found", f"=COUNTIF('Portal Status'!E2:E{np_},\"Portal not found\")", "No job listing page (email/resume form only) - see Notes."),
        ("Jobs", None, None),
        ("Total matching jobs", f"=COUNTA(Jobs!A2:A{nj})", None),
        ("  High", f"=COUNTIF(Jobs!H2:H{nj},\"High\")", None),
        ("  Med", f"=COUNTIF(Jobs!H2:H{nj},\"Med\")", None),
        ("  Low", f"=COUNTIF(Jobs!H2:H{nj},\"Low\")", None),
        ("  NEW since last run", f"=COUNTIF(Jobs!K2:K{nj},\"NEW\")", None),
    ]
    r = 2
    for a, b, c in lines:
        sm.cell(r, 1, a).font = body if b is not None else Font(name=F, bold=True, size=11, color="1F3864")
        if b is not None:
            sm.cell(r, 2, b).font = Font(name=F, size=10, color="0000FF") if isinstance(b, int) else bold
            sm.cell(r, 2).alignment = Alignment(horizontal="left")
        if c:
            sm.cell(r, 3, c).font = note
        r += 1
    r += 1
    sm.cell(r, 1, "Top 5 jobs to apply first (auto-ranked by match score)").font = Font(name=F, bold=True, size=11, color="1F3864")
    r += 1
    for k, h in enumerate(["#  Company – Job title", "Match / Experience", "Why  /  link"]):
        c = sm.cell(r, k + 1, h)
        c.font, c.fill = hdr_font, hdr_fill
    for n, j in enumerate(top5, 1):
        r += 1
        sm.cell(r, 1, f"{n}. {j['company']} – {j['title']} ({j['location'][:40]})").font = bold
        sm.cell(r, 2, f"{j['match']} / {j['experience']}").font = body
        sm.cell(r, 3, j["why"]).font = body
        for c in (sm.cell(r, 1), sm.cell(r, 2), sm.cell(r, 3)):
            c.alignment = wrap
        r += 1
        sm.cell(r, 3, j["url"]).hyperlink = j["url"]
        sm.cell(r, 3).font = link
    r += 2
    for t in ["Match is a keyword-based score: your skills found in the JD (CAN/UDS/J1939/AUTOSAR/STM32/RTOS/BSP/drivers weigh most), "
              "company domain (automotive > semiconductor > power/EV), and experience asked (≤3 yrs boosts, >5 yrs = Senior – stretch). "
              "Always read the JD before applying.",
              "Only official company portals are checked (no LinkedIn/Naukri). 'Description not read' = only the title was available.",
              "Black numbers are formulas (update if you delete rows); blue numbers come from the input file."]:
        sm.cell(r, 1, t).font = note
        sm.merge_cells(start_row=r, start_column=1, end_row=r, end_column=3)
        sm.cell(r, 1).alignment = wrap
        sm.row_dimensions[r].height = 30
        r += 1
    wb.calculation.fullCalcOnLoad = True
    wb.save(path)


# ===================== Main run =====================
def job_key(j):
    return hashlib.sha1(f"{j['company']}|{j['url'] or j['title']}".encode()).hexdigest()[:16]


def run_once(args):
    companies = load_companies(args.companies)
    cfg_all = load_yaml()
    overrides = cfg_all.get("overrides") or {}
    skip = set(cfg_all.get("skip") or [])
    cache = {} if args.rediscover else load_json(CACHE_FILE)
    seen = load_json(SEEN_FILE)
    now = dt.datetime.now()
    if args.only:
        companies = [c for c in companies if args.only.lower() in c["Company"].lower()]

    # group input rows by portal (same company in 2 cities -> 1 fetch)
    groups = {}
    for c in companies:
        if c["Company"] in skip:
            continue
        ov = overrides.get(c["Company"])
        if ov and ov.get("same_as"):
            continue
        key = json.dumps(ov, sort_keys=True) if ov else f"site:{c['Website'].lower()}"
        groups.setdefault(key, {"cfg": ov, "rows": []})["rows"].append(c)
    print(f"\n[{now:%Y-%m-%d %H:%M}] {len(companies)} rows / {len(groups)} portals check ho rahe hain...\n")

    def work(key, g):
        s = session()
        row = g["rows"][0]
        ctx = {"any_location": args.any_location, "note": "", "status": None}
        cfg = g["cfg"] or cache.get(row["Company"])
        try:
            if not cfg:
                cfg = discover(s, row["Website"])
                if cfg and cfg["type"] != "generic":
                    cache[row["Company"]] = cfg
            if not cfg:
                return key, None, [], {"status": "Portal not found", "note": "Careers page nahi mila - portals.yaml mein add karo"}
            jobs = FETCHERS[cfg["type"]](s, cfg, ctx)
            return key, cfg, jobs, {"status": ctx["status"], "note": ctx["note"]}
        except Blocked as e:
            return key, cfg, [], {"status": "Blocked", "note": f"{e} - browser mein check karo"}
        except Exception as e:
            return key, cfg, [], {"status": "Blocked", "note": f"Error {type(e).__name__}: {str(e)[:160]} - browser mein check karo"}

    results = []
    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        futs = {ex.submit(work, k, g): k for k, g in groups.items()}
        for i, f in enumerate(as_completed(futs), 1):
            key, cfg, jobs, st = f.result()
            results.append((key, cfg, jobs, st))
            g = groups[key]
            print(f"  [{i:>2}/{len(groups)}] {g['rows'][0]['Company'][:38]:<38} {(cfg or {}).get('type', '-'):<15} "
                  f"{st['status'] or (str(len(jobs)) + ' listed')}")

    all_jobs, status_map = [], {}
    for key, cfg, jobs, st in results:
        rows = groups[key]["rows"]
        kept = {}
        for j in jobs:
            if not relevant(j) or not located(j, args.any_location):
                continue
            row = city_of(j, rows)
            j["company"], j["city"] = row["Company"], row["City"]
            analyze(j, row)
            kept[(j["company"], j["city"], j["url"] or j["title"])] = j
        for j in kept.values():
            k = job_key(j)
            j["new"] = k not in seen
            seen.setdefault(k, now.strftime("%Y-%m-%d %H:%M"))
            all_jobs.append(j)
        for row in rows:
            n = sum(1 for j in kept.values() if j["company"] == row["Company"] and j["city"] == row["City"])
            stt = st["status"] or ("Jobs found" if n else "No matching jobs")
            status_map[(row["Company"], row["City"])] = {
                "company": row["Company"], "city": row["City"], "url": portal_url(cfg or {}),
                "type": TYPE_LABEL.get((cfg or {}).get("type"), "-"), "status": stt, "note": st["note"] or ""}
    # same_as / skip rows
    for c in companies:
        k = (c["Company"], c["City"])
        if k in status_map:
            continue
        ov = overrides.get(c["Company"]) or {}
        if ov.get("same_as"):
            src = next((v for (n, _), v in status_map.items() if n == ov["same_as"]), None)
            status_map[k] = {"company": c["Company"], "city": c["City"], "url": (src or {}).get("url", ""),
                             "type": (src or {}).get("type", "-"), "status": "No matching jobs",
                             "note": f"Same portal as {ov['same_as']} - jobs are listed under that company."}
        else:
            status_map[k] = {"company": c["Company"], "city": c["City"], "url": "", "type": "-", "status": "Portal not found",
                             "note": "Skipped (portals.yaml skip list)"}
    status = [status_map[(c["Company"], c["City"])] for c in companies]

    order = {"High": 0, "Med": 1, "Low": 2}
    all_jobs.sort(key=lambda j: (order[j["match"]], -j["score"], j["company"].lower()))
    top5 = [j for j in all_jobs if "Senior – stretch" not in j["flags"]][:5]
    save_json(CACHE_FILE, cache)
    save_json(SEEN_FILE, seen)
    OUT_DIR.mkdir(exist_ok=True)
    out = OUT_DIR / f"Embedded_Jobs_{now:%Y-%m-%d_%H%M}.xlsx"
    write_excel(all_jobs, status, companies, out, top5, f"{now:%Y-%m-%d %H:%M}")

    new = [j for j in all_jobs if j["new"]]
    cnt = {k: sum(1 for s in status if s["status"] == k) for k in ("Jobs found", "No matching jobs", "Blocked", "Portal not found")}
    print(f"\nRows checked: {len(status)}/{len(companies)}  |  " + "  ".join(f"{k}: {v}" for k, v in cnt.items()))
    print(f"Matching jobs: {len(all_jobs)} (High {sum(j['match'] == 'High' for j in all_jobs)}, "
          f"Med {sum(j['match'] == 'Med' for j in all_jobs)}, Low {sum(j['match'] == 'Low' for j in all_jobs)})  |  NEW: {len(new)}")
    for j in top5:
        print(f"  TOP  {j['match']:<4} {j['company'][:24]:<24} {j['title'][:50]:<50} {j['experience']}")
    print(f"\nResult file: {out}\n")


# ===================== Self test =====================
def selftest():
    ok = True

    def check(name, cond):
        nonlocal ok
        ok &= bool(cond)
        print("PASS" if cond else "FAIL", name)
    check("detect workday", (detect("https://acme.wd1.myworkdayjobs.com/External_Careers/job/x") or {}).get("type") == "workday")
    check("detect lever", (detect("https://jobs.lever.co/acme") or {}).get("id") == "acme")
    check("detect oracle", (detect("https://ab.fa.ocs.oraclecloud.com/hcmUI/CandidateExperience/en/sites/CX_1/jobs") or {}).get("site") == "CX_1")
    check("exp 2-4", experience("B.Tech with 2-4 years of experience in firmware")[1] == "2-4 years")
    check("exp 4+", experience("Bachelor's degree and 4+ years of ASIC validation experience")[1] == "4+ years")
    check("exp ignores company age", experience("leader in EDA for over 25 years. Requires 3 years of experience")[1] == "3+ years")
    a = J("Firmware Engineer", "Noida", "u", "", "Embedded C on STM32 with FreeRTOS, CAN and UDS. 2-4 years experience.")
    check("relevant title", relevant(a))
    b = J("Senior Software Engineer", "Gurugram", "u", "", "microcontroller firmware, RTOS, device drivers, I2C, SPI")
    check("relevant via desc", relevant(b))
    c = J("Layout Design Engineer", "Noida", "u", "", "embedded memory compilers SRAM")
    check("embedded memory not relevant", not relevant(c))
    d = J("HR Executive", "Gurugram", "u", "", "")
    check("hr not relevant", not relevant(d))
    check("loc noida", located(J("x", "Noida, Uttar Pradesh, India", "u"), False))
    check("loc bangalore", not located(J("x", "Bengaluru, India", "u", "", "Bangalore office"), False))
    e = J("x", "2 Locations", "u")
    check("loc vague kept", located(e, False) and "Location unclear" in e["flags"])
    analyze(a, {"Sector": "Automotive electronics", "Match for Aman": "High"})
    check(f"score high ({a['score']})", a["match"] == "High")
    html = '<a href="/jobs/embedded-software-engineer/">Embedded Software Engineer</a><a href="/about">About</a>'
    check("A_RE", len([1 for h, t in A_RE.findall(html) if "jobs" in h]) == 1)
    check("_fill", _fill({"q": "{term}", "p": "{page}"}, "fw", 2, 0) == {"q": "fw", "p": "2"})
    check("_dig wildcard", _dig({"data": {"A": {"jobs": [1, 2]}, "B": {"jobs": [3]}}}, "data.*.jobs") == [1, 2, 3])
    check("exp 'Experience: 4+ years'", experience("Experience: 4+ years' experience in relevant technology")[1] == "4+ years")
    check("exp 'Exp: 7 to 12yrs'", experience("Embedded electronics Exp: 7 to 12yrs o Laboratory")[1] == "7-12 years")
    check("exp hint", experience("no numbers here", "12 - 15 Years")[1] == "12-15 years")
    check("RTL title not relevant", not relevant(J("RTL Design - Technical Lead", "Noida", "u", "", "embedded firmware i2c spi uart rtos")))
    check("S/W ECU title relevant via JD", relevant(J("Cybersecurity, S/W Electronic control system", "Gurugram", "u", "",
                                                       "ECU software, AUTOSAR, CAN bus, UDS diagnostics, embedded C on microcontrollers")))
    k = J("x", "India", "u", "", "Embedded firmware test for metering products")
    check("India-only kept as unclear", located(k, False) and "Location unclear" in k["flags"])
    check("Multiple Locations w/o NCR dropped", not located(J("x", "Multiple Locations", "u", "", "Prague, Czech Republic"), False))
    q = J("Lead Engineer - Validation", "Noida", "u", "", "Bachelor's degree and 4+ years of experience. firmware, RTOS, device drivers")
    analyze(q, {"Sector": "Semiconductor", "Match for Aman": "Medium"})
    check("Lead 4+ yrs not stretch", "Senior – stretch" not in q["flags"])
    print("\nSELFTEST", "OK" if ok else "FAILED")
    return ok


def main():
    ap = argparse.ArgumentParser(description="Embedded job checker v2 - Delhi NCR")
    ap.add_argument("--companies", default=str(DEFAULT_XLSX), help="Company list Excel ka path")
    ap.add_argument("--every", type=int, default=0, help="Har N minute mein dobara chalao")
    ap.add_argument("--workers", type=int, default=8, help="Ek saath kitne portals check hon")
    ap.add_argument("--only", default="", help="Sirf is naam wali company (debug)")
    ap.add_argument("--any-location", action="store_true", help="NCR location filter band")
    ap.add_argument("--rediscover", action="store_true", help="Auto-discovered portal cache reset")
    ap.add_argument("--selftest", action="store_true", help="Offline parser test")
    args = ap.parse_args()
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
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
