import csv
import json
import re
import time
from collections import deque
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup

# ---------------- settings ----------------
PARTNERS_URL = ""
OUTPUT_CSV = "partner_locations.csv"
MAX_PAGES_PER_SITE = 40          
USE_BROWSER_FALLBACK = True      
WORKERS = 6
TIMEOUT = 20
DELAY = 0.5                      

HEADERS = {
    "User-Agent": ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
                   "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"),
    "Accept-Language": "en-US,en;q=0.9",
}

FALLBACK_PARTNERS = [
    
]

COMMON_PATHS = ["contact", "contact-us", "locations", "our-locations", "location",
                "find-a-location", "offices", "our-offices"]
LOCATION_KEYWORDS = ("location", "contact", "office", "find-us", "directions", "map", "visit")
SKIP_EXT = (".pdf", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".zip",
            ".doc", ".docx", ".mp4", ".mp3", ".ics")

# ---------------- regexes ----------------
STATES = ("AL|AK|AZ|AR|CA|CO|CT|DE|DC|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|"
          "MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY")
CITY_STATE_ZIP = re.compile(
    rf"([A-Za-z][A-Za-z.'\- ]{{1,40}}?),?\s+({STATES})\.?,?\s+(\d{{5}}(?:-\d{{4}})?)\b")
STREET = re.compile(r"^\s*(\d{1,6}[A-Za-z]?\s+\S+|P\.?\s*O\.?\s*Box)", re.I)
SUITE = re.compile(r"\b(suite|ste\.?|floor|fl\.|building|bldg\.?|unit|entrance)\b|#\s*\d", re.I)
PHONE = re.compile(r"\(?\b\d{3}\)?[\s.\-]?\d{3}[\s.\-]\d{4}\b")


# ---------------- extraction ----------------
def clean(s):
    return re.sub(r"\s+", " ", s or "").strip(" ,")


def extract_jsonld(soup):
    """schema.org addresses from <script type=application/ld+json>."""
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
                        region_zip = clean(f'{a.get("addressRegion", "")} {a.get("postalCode", "")}')
                        full = ", ".join(x for x in [clean(a.get("streetAddress")),
                                                     clean(a.get("addressLocality")), region_zip] if x)
                        out.append({"location_name": clean(str(node.get("name", ""))),
                                    "address": full,
                                    "phone": clean(str(node.get("telephone", ""))),
                                    "method": "json-ld"})
            stack.extend(v for k, v in node.items() if k != "address" and isinstance(v, (dict, list)))
    return out


def extract_text(soup):
    """Addresses from visible text: find 'City, ST 12345' and pull street/suite lines above it."""
    for tag in soup(["script", "style", "noscript", "svg", "iframe"]):
        tag.decompose()
    lines = [clean(l) for l in soup.get_text("\n").split("\n")]
    lines = [l for l in lines if l]

    out = []
    for i, line in enumerate(lines):
        m = CITY_STATE_ZIP.search(line)
        if not m:
            continue

        if STREET.match(line):                      # whole address on one line
            parts = [line[:m.end()]]
            first = i
        else:                                       # address split over lines
            prev, first = [], i
            for j in range(i - 1, max(i - 4, -1), -1):
                if STREET.match(lines[j]) or SUITE.search(lines[j]):
                    prev.insert(0, lines[j])
                    first = j
                else:
                    break
            if not prev:
                continue                            # city/state/zip alone = probably noise
            parts = prev + [line[:m.end()]]

        # name = short text line right above the address (e.g. "Anderson Area Cancer Center")
        name = ""
        if first > 0:
            cand = lines[first - 1]
            if len(cand) <= 70 and not PHONE.search(cand) and not re.search(r"\d{3,}", cand):
                name = cand

        phone = ""
        for k in range(i, min(i + 5, len(lines))):
            p = PHONE.search(lines[k])
            if p:
                phone = p.group(0)
                break

        out.append({"location_name": name,
                    "address": clean(", ".join(p.strip(" ,") for p in parts)),
                    "phone": phone,
                    "method": "text"})
    return out


# ---------------- crawling ----------------
def normalize(url):
    url = urldefrag(url)[0]
    return url.rstrip("/")


def make_requests_fetcher():
    session = requests.Session()
    session.headers.update(HEADERS)

    def fetch(url):
        time.sleep(DELAY)
        try:
            r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
            if r.status_code == 200 and "text/html" in r.headers.get("Content-Type", ""):
                return r.url, r.text
        except requests.RequestException:
            pass
        return None, None
    return fetch


def crawl_site(name, base_url, fetch):
    parsed = urlparse(base_url)
    netloc = parsed.netloc.replace("www.", "")
    base_path = parsed.path if parsed.path not in ("", "/") else "/"
    prefix = f"{parsed.scheme}://{parsed.netloc}{base_path}"
    if not prefix.endswith("/"):
        prefix += "/"

    def in_scope(link):
        p = urlparse(link)
        return (p.scheme in ("http", "https")
                and p.netloc.replace("www.", "") == netloc
                and p.path.startswith(base_path.rstrip("/") or "/")   # keeps unitedurology.com/<practice>/ separate
                and not p.path.lower().endswith(SKIP_EXT))

    queue = deque([base_url] + [urljoin(prefix, p) for p in COMMON_PATHS])
    seen, records, pages = set(), [], 0

    while queue and pages < MAX_PAGES_PER_SITE:
        url = normalize(queue.popleft())
        if url in seen:
            continue
        seen.add(url)

        final_url, html = fetch(url)
        if not html:
            continue
        seen.add(normalize(final_url))
        pages += 1

        soup = BeautifulSoup(html, "lxml")
        links = soup.find_all("a", href=True)

        # collect links BEFORE extract_text (it removes tags)
        for a in links:
            link = normalize(urljoin(final_url, a["href"]))
            if link in seen or not in_scope(link):
                continue
            haystack = (a.get_text(" ") + " " + urlparse(link).path).lower()
            if any(k in haystack for k in LOCATION_KEYWORDS):
                queue.append(link)

        for rec in extract_jsonld(soup) + extract_text(soup):
            rec.update(company=name, website=base_url, source_page=final_url)
            records.append(rec)

    return dedupe(records), pages


def dedupe(records):
    best = {}
    for r in records:
        key = re.sub(r"[^a-z0-9]", "", r["address"].lower())
        if not key:
            continue
        if key not in best:
            best[key] = r
        else:   # fill missing name/phone from duplicates
            for f in ("location_name", "phone"):
                if not best[key][f] and r[f]:
                    best[key][f] = r[f]
    return list(best.values())


# ---------------- partner list ----------------
def get_partners():
    try:
        r = requests.get(PARTNERS_URL, headers=HEADERS, timeout=TIMEOUT)
        soup = BeautifulSoup(r.text, "lxml")
        partners, seen = [], set()
        for a in soup.select("a.BaseLink[href^='http']"):
            href = a["href"]
            if "oneoncology.com" in href or href in seen:
                continue
            seen.add(href)
            img = a.find("img")
            name = (a.get_text(strip=True)
                    or (img.get("alt", "").strip() if img else "")
                    or urlparse(href).netloc.replace("www.", ""))
            partners.append((name, href))
        if len(partners) >= 20:
            print(f"Got {len(partners)} partners from live page")
            return partners
    except Exception as e:
        print("Could not read partners page:", e)
    print(f"Using fallback list ({len(FALLBACK_PARTNERS)} partners)")
    return FALLBACK_PARTNERS


# ---------------- browser fallback ----------------
def browser_pass(partners):
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        print("Playwright not installed - skipping browser fallback")
        return {}

    results = {}
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True)
        page = browser.new_page(user_agent=HEADERS["User-Agent"])

        def fetch(url):
            try:
                resp = page.goto(url, timeout=30000, wait_until="domcontentloaded")
                page.wait_for_timeout(2500)          # let JS render
                if resp and resp.ok:
                    return page.url, page.content()
            except Exception:
                pass
            return None, None

        for name, url in partners:
            print(f"[browser] {name}")
            recs, pages = crawl_site(name, url, fetch)
            results[(name, url)] = recs
            print(f"   -> {len(recs)} locations from {pages} pages")
        browser.close()
    return results


# ---------------- main ----------------
def main():
    partners = get_partners()
    all_results = {}

    with ThreadPoolExecutor(max_workers=WORKERS) as ex:
        futures = {ex.submit(crawl_site, n, u, make_requests_fetcher()): (n, u) for n, u in partners}
        for f in as_completed(futures):
            n, u = futures[f]
            try:
                recs, pages = f.result()
            except Exception as e:
                print(f"[error] {n}: {e}")
                recs, pages = [], 0
            all_results[(n, u)] = recs
            print(f"{n}: {len(recs)} locations ({pages} pages)")

    empty = [k for k, v in all_results.items() if not v]
    if empty and USE_BROWSER_FALLBACK:
        print(f"\nRetrying {len(empty)} empty sites with a browser...")
        all_results.update({k: v for k, v in browser_pass(empty).items() if v})

    rows = [r for n, u in partners for r in all_results.get((n, u), [])]
    cols = ["company", "website", "location_name", "address", "phone", "source_page", "method"]
    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=cols)
        w.writeheader()
        w.writerows(rows)

    still_empty = [n for (n, u) in partners if not all_results.get((n, u))]
    print(f"\nSaved {len(rows)} locations to {OUTPUT_CSV}")
    if still_empty:
        print("No locations found for:", ", ".join(still_empty))


if __name__ == "__main__":
    main()
