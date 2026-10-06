
import csv
import json
import re
import sys
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup

# ================= PASTE YOUR LIST HERE =================

# ========================================================

OUTPUT_CSV = "partner_locations.csv"
FOLLOW_SUBPAGES = False   # True = also open links directly under a listed page (e.g. /locations/xyz)
MAX_SUBPAGES = 40         # per listed URL, only used when FOLLOW_SUBPAGES = True
WORKERS = 8
TIMEOUT = 15
DELAY = 0.2
USE_BROWSER_FALLBACK = True

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
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
            if not phone:   # phone often sits in the next <p>/<div> after the address
                last = list(el.descendants)[-1] if el.contents else el
                after = " ".join(clean(t) for t in last.find_all_next(string=True, limit=8))
                nxt = CSZ.search(after)              # stop before the next address starts
                phone = pick_phone(after[:nxt.start() if nxt else 200][:200])
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


# ---------------- scraping ----------------
def scrape_page(final, html, company, url):
    soup = BeautifulSoup(html, "lxml")
    found = extract_jsonld(soup, final)
    anchors = [urljoin(final, a["href"]) for a in soup.find_all("a", href=True)]
    found += extract_blocks(soup, final)
    for r in found:
        r["company"], r["website"] = company, url
    return found, anchors


def scrape_url(company, url, fetch, verbose=False):
    """Returns (records, needs_browser)."""
    company = company.strip()
    final, html = fetch(url)
    if not html or is_thin(html):
        return [], True

    records, anchors = scrape_page(final, html, company, url)
    if verbose:
        print(f"  {final} -> {len(records)} address(es)")

    if FOLLOW_SUBPAGES:
        base = urlparse(final)
        base_path = base.path.rstrip("/")
        seen, count = {normalize(final), normalize(url)}, 0
        for link in anchors:
            p = urlparse(link)
            if (count >= MAX_SUBPAGES or normalize(link) in seen
                    or p.netloc != base.netloc
                    or not p.path.rstrip("/").startswith(base_path + "/")
                    or p.path.lower().endswith(SKIP_EXT)):
                continue
            seen.add(normalize(link))
            f2, h2 = fetch(link)
            if not h2:
                continue
            count += 1
            recs, _ = scrape_page(f2, h2, company, url)
            records.extend(recs)
            if verbose:
                print(f"    sub: {f2} -> {len(recs)} address(es)")

    return records, not records


def browser_pass(urls, verbose=False):
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

        for company, url in urls:
            recs, _ = scrape_url(company, url, fetch, verbose)
            results[(company, url)] = recs
            print(f"[browser] {company}: {len(recs)} address(es)")
        browser.close()
    return results


# ---------------- main ----------------
def main():
    only = sys.argv[1].lower() if len(sys.argv) > 1 else None
    partners = [(n.strip(), u.strip()) for n, u in PARTNERS if u and u.strip()]
    partners = list(dict.fromkeys(partners))          # drop exact repeats
    if only:
        partners = [(n, u) for n, u in partners if only in u.lower() or only in n.lower()]
    if not partners:
        print("Nothing to scrape - paste your list into PARTNERS at the top of the file.")
        return
    verbose = only is not None

    results, needs_browser = {}, []
    with ThreadPoolExecutor(max_workers=1 if verbose else WORKERS) as ex:
        futures = {ex.submit(scrape_url, n, u, make_requests_fetcher(), verbose): (n, u) for n, u in partners}
        for f in as_completed(futures):
            key = futures[f]
            try:
                recs, nb = f.result()
            except Exception as e:
                print(f"[error] {key[0]}: {e}")
                recs, nb = [], True
            results[key] = recs
            if nb:
                needs_browser.append(key)
            print(f"{key[0]} ({key[1]}): {len(recs)} address(es){'  -> browser' if nb else ''}")

    if needs_browser and USE_BROWSER_FALLBACK:
        print(f"\nBrowser pass for {len(needs_browser)} page(s)...")
        for key, recs in browser_pass(needs_browser, verbose).items():
            if recs:
                results[key] = recs

    # dedupe per company (same location on /contact and /locations = 1 row)
    rows = dedupe([r for key in partners for r in results.get(key, [])])
    order = {n: i for i, (n, _) in enumerate(partners)}  # keep your list order
    rows.sort(key=lambda r: (order.get(r["company"], 0), r["location_name"].lower(), r["address"]))

    cols = ["company", "location_name", "address", "phone", "website", "source_page", "method"]
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=cols, extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)

    empty = [f"{n} - {u}" for n, u in partners if not results.get((n, u))]
    print(f"\nSaved {len(rows)} locations to {OUTPUT_CSV}")
    if empty:
        print("No addresses found for:\n  " + "\n  ".join(empty))


if __name__ == "__main__":
    main()
