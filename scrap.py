
import csv
import json
import re
import sys
import time
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup

# ---------------- settings ----------------
OUTPUT_CSV = "partner_locations.csv"
MAX_PAGES_PER_SITE = 25
MAX_DEPTH = 2                  
WORKERS = 8
TIMEOUT = 15
DELAY = 0.2
USE_BROWSER_FALLBACK = True

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}


PARTNERS = [
   
]

SITE_OVERRIDES = {

}

# ---------------- patterns ----------------
STATE_ABBR = ("AL|AK|AZ|AR|CA|CO|CT|DE|DC|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|"
              "MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY")
STATE_NAMES = ("Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|"
               "Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|"
               "Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|"
               "New Hampshire|New Jersey|New Mexico|New York|North Carolina|North Dakota|Ohio|Oklahoma|"
               "Oregon|Pennsylvania|Rhode Island|South Carolina|South Dakota|Tennessee|Texas|Utah|"
               "Vermont|Virginia|Washington|West Virginia|Wisconsin|Wyoming")
# ", SC 29621"  /  " South Carolina 29621"
CSZ = re.compile(rf"(?:,|\s)\s*(?:{STATE_NAMES}|{STATE_ABBR})\.?,?\s+(\d{{5}})(?:-\d{{4}})?\b")

SUFFIX = (r"(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln|Way|Parkway|Pkwy|"
          r"Highway|Hwy|Court|Ct|Place|Pl|Circle|Cir|Pike|Trail|Trl|Plaza|Square|Sq|Terrace|Ter|"
          r"Loop|Center|Ctr|Expressway|Expy|Broadway|Route|Rte|Turnpike|Tpke|Crossing|Xing|"
          r"Commons|Point|Pt|Run|Row|Path|Walk|Mall|Freeway|Fwy)")
# street must have a number AND a real street suffix -> kills most junk matches
STREET = re.compile(rf"\b\d{{1,6}}[A-Za-z]?(?:-\d+)?\s+(?:[A-Za-z0-9.'#-]+\s+){{0,5}}?{SUFFIX}\b\.?")
PHONE = re.compile(r"\(?\b\d{3}\)?[\s.\-]?\d{3}[\s.\-]\d{4}\b")
ZIP_ONLY = re.compile(r"\b\d{5}(?:-\d{4})?\b")

NAME_JUNK = re.compile(
    r"^(contact( us)?|locations?|our locations?|address|directions|get directions|view map|map|"
    r"hours|office hours|visit us|find us|phone|fax|tel|main office|office|menu|home|quick links|"
    r"follow us|connect with us|location details|clinic locations?)$", re.I)
NAME_LABEL = re.compile(r"\b(phone|tel|fax|address|get directions|directions|view map)\b\s*:?", re.I)

# crawl rules
EXCLUDE_PATH = re.compile(
    r"/(blog|news|careers?|jobs?|physicians?|doctors?|providers?|our-team|team|events?|press|media|"
    r"patient-portal|portal|pay|bill-pay|billing|login|privacy|terms|sitemap|wp-content|wp-json|tag|"
    r"category|feed|search|clinical-trials|research|donate|give|foundation|insurance|forms?|"
    r"accessibility|non-discrimination|videos?)(/|$)", re.I)
LOC_PATH = re.compile(r"location|contact|office|find-us|directions|our-clinics|facilities", re.I)
LOC_TEXT = re.compile(r"\b(locations?|contact( us)?|offices?|find us|directions|our clinics|facilities)\b", re.I)
SKIP_EXT = (".pdf", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".zip",
            ".doc", ".docx", ".mp4", ".mp3", ".ics", ".xml")
COMMON_PATHS = ["locations", "contact-us", "contact"]


# ---------------- helpers ----------------
def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def tidy_address(a):
    a = clean(a)
    a = re.sub(r"\s+,", ",", a)
    a = re.sub(r",\s*,+", ",", a)
    return a.strip(" ,;-|")


def clean_name(s):
    s = PHONE.sub(" ", s or "")
    s = NAME_LABEL.sub(" ", s)
    s = clean(s).strip(" :-|•,·>")
    if not s or len(s) < 3 or len(s) > 80:
        return ""
    if NAME_JUNK.match(s) or len(re.findall(r"\d", s)) >= 3:
        return ""
    if s.isupper():
        s = s.title()
    return s


def pick_phone(text):
    for m in PHONE.finditer(text):
        before = text[max(0, m.start() - 12):m.start()].lower()
        if "fax" in before or "f:" in before:
            continue
        return m.group(0)
    return ""


def normalize(url):
    url = urldefrag(url)[0].split("?")[0]
    return url.rstrip("/").lower()


def get_override(url):
    for k, v in SITE_OVERRIDES.items():
        if k in url:
            return v
    return {}


# ---------------- extraction ----------------
def extract_jsonld(soup, page_url):
    out = []
    for s in soup.find_all("script", type="application/ld+json"):
        try:
            data = json.loads(s.string or s.get_text() or "")
        except Exception:
            continue
        stack = [data]
        while stack:
            node = stack.pop()
            if isinstance(node, list):
                stack.extend(node)
                continue
            if not isinstance(node, dict):
                continue
            addrs = node.get("address")
            if isinstance(addrs, dict):
                addrs = [addrs]
            if isinstance(addrs, list):
                for a in addrs:
                    if isinstance(a, dict) and a.get("streetAddress"):
                        tail = clean(f'{a.get("addressRegion", "")} {a.get("postalCode", "")}')
                        full = ", ".join(x for x in [clean(a.get("streetAddress")),
                                                     clean(a.get("addressLocality")), tail] if x)
                        out.append({"location_name": clean_name(str(node.get("name", ""))),
                                    "address": tidy_address(full),
                                    "phone": clean(str(node.get("telephone", ""))),
                                    "source_page": page_url, "method": "json-ld"})
            stack.extend(v for k, v in node.items() if k != "address" and isinstance(v, (dict, list)))
    return out


def address_blocks(soup):
    """Smallest elements (<= 400 chars of text) that contain a street + state + zip."""
    blocks, seen = [], set()
    for node in soup.find_all(string=ZIP_ONLY):
        el = node.parent
        while el is not None and el.name not in ("body", "html", "[document]"):
            txt = clean(el.get_text(" "))
            if len(txt) > 400:
                break
            if CSZ.search(txt) and STREET.search(txt):
                if id(el) not in seen:
                    seen.add(id(el))
                    blocks.append(el)
                break
            el = el.parent
    return blocks


HEADINGS = ["h1", "h2", "h3", "h4", "h5", "h6", "strong", "b"]


def heading_before(el):
    h = el.find_previous(HEADINGS)
    for _ in range(3):
        if h is None:
            return ""
        n = clean_name(h.get_text(" "))
        if n:
            return n
        h = h.find_previous(HEADINGS)
    return ""


def extract_blocks(soup, page_url):
    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()
    out = []
    for el in address_blocks(soup):
        text = clean(el.get_text(" "))
        tel = el.find("a", href=re.compile(r"^tel:", re.I))
        pos = 0
        for m in CSZ.finditer(text):
            street = next(STREET.finditer(text, pos, m.start()), None)   # FIRST street after previous address
            if not street or m.end() - street.start() > 160:
                pos = m.end()
                continue
            name = clean_name(text[pos:street.start()]) or heading_before(el)
            phone = pick_phone(text[m.end():m.end() + 120])
            if not phone and tel:
                phone = clean(tel.get_text(" ")) or tel["href"][4:]
            out.append({"location_name": name,
                        "address": tidy_address(text[street.start():m.end()]),
                        "phone": phone, "source_page": page_url, "method": "html"})
            pos = m.end()
    return out


def dedupe(records):
    """Same street number + same ZIP = same location."""
    out = {}
    for r in records:
        num = re.match(r"\s*(\d+)", r["address"])
        z = re.search(r"\b(\d{5})(?:-\d{4})?\s*$", r["address"])
        key = (r["company"], num.group(1), z.group(1)) if num and z else \
              (r["company"], re.sub(r"[^a-z0-9]", "", r["address"].lower()))
        cur = out.get(key)
        if cur is None:
            out[key] = r
            continue
        if r["method"] == "json-ld" and cur["method"] != "json-ld":   # structured data wins
            r["location_name"] = r["location_name"] or cur["location_name"]
            r["phone"] = r["phone"] or cur["phone"]
            out[key] = r
        else:
            cur["location_name"] = cur["location_name"] or r["location_name"]
            cur["phone"] = cur["phone"] or r["phone"]
    return list(out.values())


# ---------------- fetchers ----------------
def is_thin(html):
    text = BeautifulSoup(html, "lxml").get_text(" ")
    return len(clean(text)) < 500 or "enable javascript" in text.lower()


def make_requests_fetcher():
    session = requests.Session()
    session.headers.update(HEADERS)

    def fetch(url):
        time.sleep(DELAY)
        try:
            r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
            if r.status_code == 200 and "html" in r.headers.get("Content-Type", ""):
                return r.url, r.text
        except requests.RequestException:
            pass
        return None, None
    return fetch


# ---------------- crawler ----------------
def crawl_site(name, base_url, fetch, verbose=False):
    """Returns (records, pages_scraped, blocked)."""
    override = get_override(base_url)
    parsed = urlparse(base_url)
    netloc = parsed.netloc.lower().replace("www.", "")
    base_path = parsed.path.rstrip("/")            # e.g. /chesapeake-urology for United Urology

    def in_scope(link):
        p = urlparse(link)
        return (p.scheme in ("http", "https")
                and p.netloc.lower().replace("www.", "") == netloc
                and p.path.startswith(base_path)
                and not p.path.lower().endswith(SKIP_EXT)
                and not EXCLUDE_PATH.search(p.path))

    queue = deque([(base_url, 0)] + [(u, 1) for u in override.get("pages", [])])
    seen, records, pages = set(), [], 0

    while queue and pages < MAX_PAGES_PER_SITE:
        url, depth = queue.popleft()
        if normalize(url) in seen:
            continue
        seen.add(normalize(url))

        final, html = fetch(url)
        if not html:
            if depth == 0:
                return [], 0, True                 # homepage blocked -> use browser
            continue
        if depth == 0 and is_thin(html):
            return [], 0, True                     # JS-rendered homepage -> use browser
        seen.add(normalize(final))
        pages += 1

        soup = BeautifulSoup(html, "lxml")
        found = extract_jsonld(soup, final)
        # collect candidate links before extract_blocks strips tags
        anchors = [(urljoin(final, a["href"]), clean(a.get_text(" "))) for a in soup.find_all("a", href=True)]
        found += extract_blocks(soup, final)

        for r in found:
            r["company"], r["website"] = name, base_url
        records.extend(found)
        if verbose:
            print(f"  [{depth}] {final}  -> {len(found)} address(es)")
            for r in found:
                print(f"       {r['location_name'] or '-'} | {r['address']} | {r['phone']}")

        if depth >= MAX_DEPTH or override.get("no_crawl"):
            continue

        this_path = urlparse(final).path.rstrip("/")
        is_loc_index = bool(re.search(r"location|office|clinic", this_path, re.I))
        added = 0
        for link, text in anchors:
            if normalize(link) in seen or not in_scope(link):
                continue
            lp = urlparse(link).path.rstrip("/")
            if (LOC_PATH.search(lp)
                    or (len(text) <= 30 and LOC_TEXT.search(text))
                    or (is_loc_index and lp.startswith(this_path + "/"))):   # /locations/xyz
                queue.append((link, depth + 1))
                added += 1

        if depth == 0 and added == 0:              # no nav links found -> try usual paths
            root = f"{parsed.scheme}://{parsed.netloc}{base_path}/"
            queue.extend((urljoin(root, p), 1) for p in COMMON_PATHS)

    return dedupe(records), pages, False


# ---------------- browser pass ----------------
def browser_pass(sites, verbose=False):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright not installed - skipping browser pass")
        return {}

    results = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=HEADERS["User-Agent"])

        def fetch(url):
            try:
                resp = page.goto(url, timeout=25000, wait_until="domcontentloaded")
                try:
                    page.wait_for_load_state("networkidle", timeout=8000)
                except Exception:
                    pass
                if resp and resp.ok:
                    return page.url, page.content()
            except Exception:
                pass
            return None, None

        for name, url in sites:
            recs, pages, _ = crawl_site(name, url, fetch, verbose)
            results[(name, url)] = recs
            print(f"[browser] {name}: {len(recs)} locations ({pages} pages)")
        browser.close()
    return results


# ---------------- main ----------------
def main():
    only = sys.argv[1].lower() if len(sys.argv) > 1 else None
    partners = [p for p in PARTNERS if not only or only in p[1].lower() or only in p[0].lower()]
    verbose = only is not None
    if not partners:
        print("No partner matches", only)
        return

    results, needs_browser = {}, []

    with ThreadPoolExecutor(max_workers=1 if verbose else WORKERS) as ex:
        futures = {}
        for n, u in partners:
            if get_override(u).get("browser"):
                needs_browser.append((n, u))
                continue
            futures[ex.submit(crawl_site, n, u, make_requests_fetcher(), verbose)] = (n, u)
        for f in as_completed(futures):
            n, u = futures[f]
            try:
                recs, pages, blocked = f.result()
            except Exception as e:
                print(f"[error] {n}: {e}")
                recs, pages, blocked = [], 0, True
            results[(n, u)] = recs
            if blocked or not recs:
                needs_browser.append((n, u))
            print(f"{n}: {len(recs)} locations ({pages} pages){'  -> browser' if blocked or not recs else ''}")

    if needs_browser and USE_BROWSER_FALLBACK:
        print(f"\nBrowser pass for {len(needs_browser)} site(s)...")
        for k, v in browser_pass(needs_browser, verbose).items():
            if v:
                results[k] = v

    rows = [r for n, u in partners for r in results.get((n, u), [])]
    cols = ["company", "website", "location_name", "address", "phone", "source_page", "method"]
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as fh:   # utf-8-sig opens cleanly in Excel
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

    empty = [n for n, u in partners if not results.get((n, u))]
    print(f"\nSaved {len(rows)} locations to {OUTPUT_CSV}")
    if empty:
        print("Still empty:", ", ".join(empty))


if __name__ == "__main__":
    main()
