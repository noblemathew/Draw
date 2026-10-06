
import csv
import html as htmllib
import json
import os
import re
import sys
import time
from collections import OrderedDict
from concurrent.futures import ThreadPoolExecutor, as_completed
from urllib.parse import urljoin, urlparse, urldefrag

import requests
from bs4 import BeautifulSoup, NavigableString


OUTPUT_CSV = "partner_locations.csv"
UNPARSED_CSV = "unparsed_addresses.csv"
FOLLOW_SUBPAGES = True    # open per-location sub-pages linked from your pages
MAX_SUBPAGES = 80         # per company
MAX_LIST_PAGES = 20       # pagination: page 2, 3, ... of a locations list
WORKERS = 8
TIMEOUT = 20
SAVE_PAGES = False        # True = save every downloaded page to ./pages for checking
PAGES_DIR = "pages"

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/124.0 Safari/537.36")
HEADERS = {"User-Agent": UA, "Accept-Language": "en-US,en;q=0.9"}

# ---------------- patterns ----------------
STATE_ABBR = ("AL|AK|AZ|AR|CA|CO|CT|DE|DC|FL|GA|HI|ID|IL|IN|IA|KS|KY|LA|ME|MD|MA|MI|MN|MS|MO|"
              "MT|NE|NV|NH|NJ|NM|NY|NC|ND|OH|OK|OR|PA|RI|SC|SD|TN|TX|UT|VT|VA|WA|WV|WI|WY")
STATE_NAMES = ("Alabama|Alaska|Arizona|Arkansas|California|Colorado|Connecticut|Delaware|Florida|"
               "Georgia|Hawaii|Idaho|Illinois|Indiana|Iowa|Kansas|Kentucky|Louisiana|Maine|Maryland|"
               "Massachusetts|Michigan|Minnesota|Mississippi|Missouri|Montana|Nebraska|Nevada|"
               "New Hampshire|New Jersey|New Mexico|New York|North Carolina|North Dakota|Ohio|Oklahoma|"
               "Oregon|Pennsylvania|Rhode Island|South Carolina|South Dakota|Tennessee|Texas|Utah|"
               "Vermont|Virginia|Washington|West Virginia|Wisconsin|Wyoming")
STATE_PAT = rf"(?:{STATE_NAMES}|{STATE_ABBR})"

# ", SC 29621" | " South Carolina 29621" | "SC , 29621" (tables) | line starting with "SC 29621"
CSZ = re.compile(rf"(?:^|,|\s)\s*{STATE_PAT}\.?\s*,?\s+(\d{{5}})(?:-\d{{4}})?\b")

SUFFIX = (r"(?:Street|St|Avenue|Ave|Road|Rd|Boulevard|Blvd|Drive|Dr|Lane|Ln|Way|Parkway|Pkwy|"
          r"Highway|Hwy|Court|Ct|Place|Pl|Circle|Cir|Pike|Trail|Trl|Plaza|Square|Sq|Terrace|Ter|"
          r"Loop|Center|Ctr|Expressway|Expy|Broadway|Route|Rte|Turnpike|Tpke|Crossing|Xing|"
          r"Commons|Point|Pt|Run|Row|Path|Walk|Mall|Freeway|Fwy|Park|Pkwy)")
STREET = re.compile(rf"\b\d{{1,6}}[A-Za-z]?(?:-\d+)?\s+(?:[A-Za-z0-9.'#-]+\s+){{0,5}}?{SUFFIX}\b\.?")
NUMSTART = re.compile(r"(?<![\w#$.,/-])\d{1,6}[A-Za-z]?(?:-\d+)?\s+(?:[NSEW]\.?\s+)?[A-Za-z]")
SUITE_BEFORE = re.compile(r"\b(suite|ste|unit|floor|fl|bldg|building|room|rm|apt|level|box)\.?\s*#?\s*$", re.I)
SUITE_LINE = re.compile(r"^\s*(suite|ste\.?|unit|floor|fl\.|bldg\.?|building|room|level|entrance|#\s*\d|"
                        r"\d+(st|nd|rd|th)\s+floor)\b", re.I)
PHONE = re.compile(r"\(?\b\d{3}\)?[\s.\-]?\d{3}[\s.\-]\d{4}\b")
ADDR_INLINE = re.compile(rf"(?<![\w$.-])\d{{1,6}}[A-Za-z]?\s+[A-Za-z][^\n<>{{}}\"|;]{{2,110}}?(?:,|\s)\s*"
                         rf"{STATE_PAT}\.?\s*,?\s+\d{{5}}(?:-\d{{4}})?\b")

NAME_JUNK = re.compile(
    r"^(contact( us)?|locations?|our locations?|all locations|address|directions|get directions|view map|map|"
    r"hours|office hours|visit us|find us|phone|fax|tel|main office|office|menu|home|quick links|"
    r"follow us|connect with us|location details|clinic locations?|view location|view details|"
    r"learn more|more info|details|open|closed|email|e-mail|mailing address|physical address)$", re.I)
NAME_BAD = re.compile(r"copyright|©|all rights|privacy|\b(a\.?m|p\.?m)\b|monday|mon\s*[-–]|"
                      r"appointment|call us|click|read more|directions|@|http", re.I)
NAME_LABEL = re.compile(r"\b(phone|tel|telephone|fax|address|location)\b\s*:", re.I)

LOC_DETAIL = re.compile(r"/(locations?|offices?|clinics?|centers?|facilities|our-locations|practices?)/[^/]+/?$", re.I)
EXCLUDE_PATH = re.compile(
    r"/(blog|news|careers?|jobs?|physicians?|doctors?|providers?|team|events?|press|media|portal|"
    r"pay|billing|login|privacy|terms|sitemap|wp-content|wp-json|tag|category|feed|search|"
    r"clinical-trials|research|donate|foundation|insurance|forms?|accessibility)(/|$)", re.I)
SKIP_EXT = (".pdf", ".jpg", ".jpeg", ".png", ".gif", ".svg", ".webp", ".zip", ".doc", ".docx",
            ".mp4", ".mp3", ".ics", ".xml", ".vcf")

BLOCK_TAGS = ["p", "div", "li", "ul", "ol", "h1", "h2", "h3", "h4", "h5", "h6", "tr", "table",
              "address", "section", "article", "header", "footer", "aside", "nav", "dt", "dd", "dl",
              "main", "form", "label", "figure", "figcaption", "blockquote", "pre", "button"]
HEADING_TAGS = {"h1", "h2", "h3", "h4", "h5", "h6"}


# ---------------- small helpers ----------------
def clean(s):
    return re.sub(r"\s+", " ", s or "").strip()


def strip_tags(s):
    s = re.sub(r"<br\s*/?>", ", ", str(s), flags=re.I)
    return clean(htmllib.unescape(re.sub(r"<[^>]+>", " ", s)))


def tidy_address(a):
    a = clean(a)
    a = re.sub(r"\s+,", ",", a)
    a = re.sub(r",(\s*,)+", ",", a)
    a = re.sub(r",(?=\S)", ", ", a)
    a = re.sub(rf"\b({STATE_ABBR}),\s+(\d{{5}})", r"\1 \2", a)      # "SC, 29650" -> "SC 29650"
    return a.strip(" ,;-|•")


def clean_name(s):
    s = PHONE.sub(" ", s or "")
    s = NAME_LABEL.sub(" ", s)
    s = clean(s).strip(" :-|•,·>–—")
    if not s or len(s) < 3 or len(s) > 80 or len(s.split()) > 10:
        return ""
    if NAME_JUNK.match(s) or NAME_BAD.search(s) or len(re.findall(r"\d", s)) >= 3:
        return ""
    if NUMSTART.match(s) or CSZ.search(s):
        return ""
    if s.isupper():
        s = s.title()
    return s


def pick_phone(text):
    for m in PHONE.finditer(text or ""):
        before = text[max(0, m.start() - 14):m.start()].lower()
        if "fax" in before or re.search(r"\bf\s*[:.]", before):
            continue
        return m.group(0)
    return ""


PAGE_PARAM = re.compile(r"(?:^|&)((?:page|paged|pg|p|pagenum|page_num|start|offset)=\d+)", re.I)
PAGE_PATH = re.compile(r"/page/\d+/?$", re.I)


def normalize(url):
    """Lower-case, no fragment, no query - except pagination params like ?page=2."""
    base, _, query = urldefrag(url)[0].partition("?")
    keep = "&".join(sorted(m.group(1).lower() for m in PAGE_PARAM.finditer(query)))
    return base.rstrip("/").lower() + ("?" + keep if keep else "")


def is_pagination(link, text, rel, page_url):
    """True if link is page 2/3/... (or 'Next') of the same listing page."""
    p, b = urlparse(link), urlparse(page_url)
    if p.netloc.lower().replace("www.", "") != b.netloc.lower().replace("www.", ""):
        return False
    same_list = PAGE_PATH.sub("", p.path).rstrip("/") == PAGE_PATH.sub("", b.path).rstrip("/")
    if not same_list:
        return False
    if "next" in rel or PAGE_PATH.search(p.path) or PAGE_PARAM.search(p.query):
        return True
    return bool(re.fullmatch(r"(next|next page|more|›|»|>|\d{1,2})", text.strip().lower())) and p.query != ""


def street_start(s, end):
    """Index where a street address begins inside s[:end], or -1."""
    phones = [(m.start(), m.end()) for m in PHONE.finditer(s)]
    cands = []
    m = STREET.search(s, 0, end)
    if m:
        cands.append(m.start())
    for mm in NUMSTART.finditer(s, 0, end):
        if SUITE_BEFORE.search(s[:mm.start()]):
            continue
        if any(a <= mm.start() < b for a, b in phones):
            continue
        cands.append(mm.start())
        break
    return min(cands) if cands else -1


# ---------------- HTML text -> lines ----------------
def page_lines(soup):
    for br in soup.find_all("br"):
        br.replace_with(NavigableString("\n"))
    for t in soup.find_all(["td", "th"]):          # table cells stay on one line, comma separated
        t.append(NavigableString(" , "))
    for t in soup.find_all(BLOCK_TAGS):
        t.insert_before(NavigableString("\n"))
        t.append(NavigableString("\n"))
    lines = [clean(l).strip(" ,") for l in soup.get_text(" ").split("\n")]
    lines = [l for l in lines if l]
    merged = []                                   # "Anderson, SC" + "29621"  -> one line
    for l in lines:
        if merged and re.fullmatch(r"\d{5}(?:-\d{4})?", l) and re.search(rf"{STATE_PAT}\.?,?$", merged[-1]):
            merged[-1] += " " + l
        else:
            merged.append(l)
    return merged


def name_and_phone_above(lines, idx, stop_idx):
    """Look up to 3 lines above idx for a location name (and a phone printed above the address)."""
    name, phone = "", ""
    for k in range(idx, max(idx - 3, stop_idx, -1), -1):
        line = lines[k]
        if CSZ.search(line):
            break
        if not phone:
            phone = pick_phone(line)
        n = clean_name(line)
        if n:
            name = n
            break
    return name, phone


def extract_lines(lines, page_url):
    recs, unparsed = [], []
    last_used = -1                                 # don't let lookback cross into previous address
    for i, line in enumerate(lines):
        pos = 0
        for m in CSZ.finditer(line):
            seg = line[pos:m.end()]
            csz_rel = m.start() - pos
            st = street_start(seg, csz_rel)
            name, phone, address, name_from = "", "", None, None

            if st >= 0:                            # street + city/state/zip on the same line
                address = seg[st:]
                name = clean_name(seg[:st])
                name_from = i - 1
            else:                                  # address split over several lines
                parts, j = [], i - 1
                if not seg[:csz_rel].strip(" ,") and j > last_used and not CSZ.search(lines[j]):
                    parts.insert(0, lines[j])      # line was just "SC 29621" -> previous is city
                    j -= 1
                found = False
                while j > last_used and i - j <= 5:
                    l = lines[j]
                    if CSZ.search(l):
                        break
                    s2 = street_start(l, len(l))
                    if s2 >= 0:
                        if s2 > 0:
                            name = clean_name(l[:s2])
                        parts.insert(0, l[s2:])
                        found = True
                        j -= 1
                        break
                    if SUITE_LINE.search(l) and len(l) <= 70:
                        parts.insert(0, l)
                        j -= 1
                        continue
                    break
                if found:
                    address = ", ".join(parts + [seg.strip(" ,")])
                    name_from = j

            if address:
                if not name and name_from is not None:
                    name, phone = name_and_phone_above(lines, name_from, last_used)
                after = line[m.end():]
                for k in range(i + 1, min(i + 5, len(lines))):
                    if CSZ.search(lines[k]):
                        break
                    after += " " + lines[k]
                phone = pick_phone(after) or phone
                recs.append({"location_name": name, "address": tidy_address(address),
                             "phone": phone, "source_page": page_url, "method": "html"})
                last_used = i
            else:
                unparsed.append({"source_page": page_url,
                                 "text": " | ".join(lines[max(0, i - 3):i + 1])})
            pos = m.end()
    return recs, unparsed


# ---------------- data (JSON / scripts / attributes) ----------------
STREET_KEYS = ("streetaddress", "street_address", "address1", "address_1", "address_line_1",
               "addressline1", "line1", "street", "address")
LINE2_KEYS = ("address2", "address_2", "address_line_2", "addressline2", "line2", "street2", "suite")
CITY_KEYS = ("city", "addresslocality", "locality", "town")
STATE_KEYS = ("state", "addressregion", "region", "state_code", "province")
ZIP_KEYS = ("zip", "zipcode", "zip_code", "postalcode", "postal_code", "postcode", "postal")
NAME_KEYS = ("name", "location_name", "locationname", "title", "store", "label")
PHONE_KEYS = ("telephone", "phone", "phone_number", "phonenumber", "tel", "main_phone")


def sval(v):
    if isinstance(v, dict):
        v = v.get("rendered") or v.get("value") or ""
    if isinstance(v, (str, int, float)) and not isinstance(v, bool):
        return strip_tags(v)
    return ""


def first(low, keys):
    for k in keys:
        if k in low:
            s = sval(low[k])
            if s:
                return s
    return ""


def walk_json(obj, out, method="json"):
    stack, steps = [(obj, "", "")], 0
    while stack and steps < 300000:
        steps += 1
        node, ctx_name, ctx_phone = stack.pop()
        if isinstance(node, list):
            stack.extend((x, ctx_name, ctx_phone) for x in node)
            continue
        if not isinstance(node, dict):
            if isinstance(node, str) and CSZ.search(node):
                scan_text(node, out, method)
            continue
        low = {str(k).lower().replace("-", "_").replace(" ", "_"): v for k, v in node.items()}
        name = clean_name(first(low, NAME_KEYS)) or ctx_name
        phone = first(low, PHONE_KEYS) or ctx_phone
        street = first(low, STREET_KEYS)
        if street and re.match(r"\s*\d", street):
            if CSZ.search(street):
                full = street
            else:
                city, state, z = first(low, CITY_KEYS), first(low, STATE_KEYS), first(low, ZIP_KEYS)
                full = ", ".join(x for x in [street, first(low, LINE2_KEYS), city, clean(f"{state} {z}")] if x) \
                    if city and (state or z) else ""
            if full:
                out.append({"location_name": name, "address": tidy_address(full),
                            "phone": pick_phone(phone) or phone, "method": method})
        for k, v in low.items():
            if isinstance(v, (dict, list)):
                stack.append((v, name, phone))
            elif isinstance(v, str) and len(v) > 15 and k not in STREET_KEYS and CSZ.search(v):
                scan_text(v, out, method)


def scan_text(text, out, method="script"):
    s = (text.replace("\\/", "/").replace('\\"', '"').replace("\\n", " ")
             .replace("\\u0026", "&").replace("\\u003c", "<").replace("\\u003e", ">"))
    s = strip_tags(s)
    for m in ADDR_INLINE.finditer(s):
        out.append({"location_name": "", "address": tidy_address(m.group(0)),
                    "phone": pick_phone(s[m.end():m.end() + 60]), "method": method})


def json_blobs(text, limit=200):
    """Find JSON objects/arrays embedded inside JavaScript (e.g. var locations = [{...}];)."""
    out, i, n = [], 0, len(text)
    while i < n and len(out) < limit:
        m = re.compile(r'[\[{]\s*[{"]').search(text, i)
        if not m:
            break
        start, depth, j, in_str, esc = m.start(), 0, m.start(), False, False
        while j < n:
            c = text[j]
            if in_str:
                if esc:
                    esc = False
                elif c == "\\":
                    esc = True
                elif c == '"':
                    in_str = False
            elif c == '"':
                in_str = True
            elif c in "[{":
                depth += 1
            elif c in "]}":
                depth -= 1
                if depth == 0:
                    break
            j += 1
        blob = text[start:j + 1]
        try:
            out.append(json.loads(blob))
            i = j + 1
        except Exception:
            i = start + 1
    return out


def try_json(text):
    t = (text or "").strip()
    if not t or t[0] not in "[{":
        return None
    try:
        return json.loads(t)
    except Exception:
        return None


# ---------------- page parsing ----------------
def parse_page(final, html, jsons):
    soup = BeautifulSoup(html, "lxml")
    anchors = [(urljoin(final, a["href"]), clean(a.get_text(" ")),
                [r.lower() for r in (a.get("rel") or [])]) for a in soup.find_all("a", href=True)]
    anchors += [(urljoin(final, l["href"]), "", ["next"])                       # <link rel="next">
                for l in soup.find_all("link", href=True) if "next" in [r.lower() for r in (l.get("rel") or [])]]
    data_recs = []

    for s in soup.find_all("script"):
        txt = s.string or s.get_text() or ""
        if not txt.strip():
            continue
        parsed = try_json(txt)
        if parsed is not None:
            walk_json(parsed, data_recs, "json")
            continue
        if re.search(r"zip|postal|address|street", txt, re.I):
            for blob in json_blobs(txt):
                walk_json(blob, data_recs, "json")
        if CSZ.search(txt):
            scan_text(txt, data_recs, "script")

    for tag in soup.find_all(True):
        for k, v in tag.attrs.items():
            if not isinstance(v, str) or len(v) < 15:
                continue
            parsed = try_json(htmllib.unescape(v))
            if parsed is not None:
                walk_json(parsed, data_recs, "json")
            elif CSZ.search(v):
                scan_text(v, data_recs, "attribute")

    for j in jsons:
        parsed = try_json(j)
        if parsed is not None:
            walk_json(parsed, data_recs, "network-json")
        elif CSZ.search(j):
            scan_text(j, data_recs, "network-json")

    for tag in soup(["script", "style", "noscript", "svg", "template", "iframe"]):
        tag.decompose()
    html_recs, unparsed = extract_lines(page_lines(soup), final)

    for r in data_recs:
        r["source_page"] = final
    return html_recs + data_recs, unparsed, anchors


# ---------------- dedupe ----------------
DIRS = r"(?:N|S|E|W|NE|NW|SE|SW|North|South|East|West)\.?"
PRIORITY = {"html": 4, "json": 3, "network-json": 3, "attribute": 2, "script": 1}


def addr_key(a):
    num = re.match(r"\s*(\d+)", a)
    word = re.match(rf"\s*\d+[A-Za-z]?(?:-\d+)?\s+(?:{DIRS}\s+)?([A-Za-z]+)", a)
    zips = re.findall(r"\b(\d{5})(?:-\d{4})?\b", a)
    if num and zips:
        return (num.group(1), word.group(1).lower()[:4] if word else "", zips[-1])
    return (re.sub(r"[^a-z0-9]", "", a.lower()),)


def dedupe(records):
    best = OrderedDict()
    for r in records:
        key = (r["company"],) + addr_key(r["address"])
        cur = best.get(key)
        if cur is None:
            best[key] = dict(r)
            continue
        if PRIORITY.get(r["method"], 0) > PRIORITY.get(cur["method"], 0):
            keep, other = dict(r), cur
        else:
            keep, other = cur, r
        keep["location_name"] = keep["location_name"] or other["location_name"]
        keep["phone"] = keep["phone"] or other["phone"]
        best[key] = keep
    return list(best.values())


# ---------------- fetchers ----------------
def requests_fetcher():
    session = requests.Session()
    session.headers.update(HEADERS)

    def fetch(url, light=False):
        try:
            r = session.get(url, timeout=TIMEOUT, allow_redirects=True)
            if r.status_code == 200 and "html" in r.headers.get("Content-Type", ""):
                return r.url, r.text, []
        except requests.RequestException:
            pass
        return None, None, []
    fetch.close = lambda: session.close()
    return fetch


def is_thin(html):
    text = clean(BeautifulSoup(html, "lxml").get_text(" "))
    return len(text) < 500 or "enable javascript" in text.lower()


# ---------------- per company ----------------
def save_page(company, n, html):
    os.makedirs(PAGES_DIR, exist_ok=True)
    fn = re.sub(r"[^A-Za-z0-9]+", "_", company)[:50] + f"_{n}.html"
    with open(os.path.join(PAGES_DIR, fn), "w", encoding="utf-8") as fh:
        fh.write(html)


def scrape_company(company, urls, fetch, verbose=False):
    """Returns (records, unparsed, pages, thin)."""
    records, unparsed, pages, thin = [], [], 0, False
    seen = set()
    sub_queue = []
    list_queue = []      # page 2, 3 ... of listed pages

    def handle(url, light, listed):
        nonlocal pages, thin
        final, html, jsons = fetch(url, light=light)
        if not html:
            if verbose:
                print(f"   FAILED {url}")
            return
        if listed and is_thin(html):
            thin = True
        seen.add(normalize(final))
        pages += 1
        if SAVE_PAGES:
            save_page(company, pages, html)
        recs, unp, anchors = parse_page(final, html, jsons)
        for r in recs:
            r["company"] = company
        records.extend(recs)
        unparsed.extend(dict(u, company=company) for u in unp)
        if verbose:
            print(f"   {'sub ' if not listed else ''}{final} -> {len(recs)} address(es), {len(unp)} unparsed")
        if listed:
            for link, text, rel in anchors:
                if normalize(link) not in seen and is_pagination(link, text, rel, final):
                    list_queue.append(link)
        if listed and FOLLOW_SUBPAGES:
            base = urlparse(final)
            base_path = PAGE_PATH.sub("", base.path).rstrip("/")
            for link, text, rel in anchors:
                p = urlparse(link)
                if is_pagination(link, text, rel, final):
                    continue
                lp = p.path.rstrip("/")
                if (p.netloc.lower().replace("www.", "") != base.netloc.lower().replace("www.", "")
                        or lp.lower().endswith(SKIP_EXT) or EXCLUDE_PATH.search(lp)
                        or normalize(link) in seen):
                    continue
                under_page = bool(base_path) and lp.startswith(base_path + "/")
                if under_page or LOC_DETAIL.search(lp):
                    sub_queue.append(link)

    for u in urls:
        if normalize(u) not in seen:
            seen.add(normalize(u))
            handle(u, light=False, listed=True)

    list_count = 0
    while list_queue and list_count < MAX_LIST_PAGES:   # pagination; new pages can add more
        link = list_queue.pop(0)
        if normalize(link) in seen:
            continue
        seen.add(normalize(link))
        handle(link, light=False, listed=True)
        list_count += 1

    count = 0
    for link in sub_queue:
        if count >= MAX_SUBPAGES:
            break
        if normalize(link) in seen:
            continue
        seen.add(normalize(link))
        handle(link, light=True, listed=False)
        count += 1

    records = dedupe(records)
    # drop "unparsed" lines that were recovered another way (same street number + zip)
    keys = {(re.match(r"\s*(\d+)", r["address"]) or [None, ""])[1] + "|" +
            (re.findall(r"\b(\d{5})\b", r["address"]) or [""])[-1] for r in records}
    clean_unp, seen_txt = [], set()
    for u in unparsed:
        zips = re.findall(r"\b(\d{5})\b", u["text"])
        nums = re.findall(r"\b(\d{1,6})\b", u["text"])
        if zips and any(f"{n}|{zips[-1]}" in keys for n in nums):
            continue
        if u["text"] not in seen_txt:
            seen_txt.add(u["text"])
            clean_unp.append(u)
    return records, clean_unp, pages, thin


def run_company(company, urls, verbose):
    fetch = requests_fetcher()
    try:
        return scrape_company(company, urls, fetch, verbose)
    finally:
        fetch.close()


# ---------------- main ----------------
def main():
    only = sys.argv[1].lower() if len(sys.argv) > 1 else None
    groups = OrderedDict()
    for n, u in PARTNERS:
        if u and u.strip():
            groups.setdefault(n.strip(), [])
            if u.strip() not in groups[n.strip()]:
                groups[n.strip()].append(u.strip())
    if only:
        groups = OrderedDict((n, us) for n, us in groups.items()
                             if only in n.lower() or any(only in u.lower() for u in us))
    if not groups:
        print("Nothing to scrape - paste your list into PARTNERS at the top of the file.")
        return
    verbose = only is not None

    results = {}
    workers = 1 if verbose else WORKERS
    t0 = time.time()
    with ThreadPoolExecutor(max_workers=workers) as ex:
        futs = {ex.submit(run_company, n, us, verbose): n for n, us in groups.items()}
        for f in as_completed(futs):
            n = futs[f]
            try:
                results[n] = f.result()
            except Exception as e:
                print(f"[error] {n}: {e}")
                results[n] = ([], [], 0, True)
            recs, unp, pages, thin = results[n]
            note = "  (page looks JavaScript-loaded - addresses may not be in the HTML)" if thin else ""
            print(f"{n}: {len(recs)} locations, {len(unp)} unparsed ({pages} pages){note}")

    rows, unparsed_rows = [], []
    for n in groups:
        recs, unp, _, _ = results.get(n, ([], [], 0, False))
        rows.extend(sorted(recs, key=lambda r: (r["location_name"].lower(), r["address"])))
        unparsed_rows.extend(unp)

    with open(OUTPUT_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["company", "location_name", "address", "phone",
                                           "source_page", "method"], extrasaction="ignore")
        w.writeheader()
        w.writerows(rows)
    with open(UNPARSED_CSV, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=["company", "source_page", "text"], extrasaction="ignore")
        w.writeheader()
        w.writerows(unparsed_rows)

    print(f"\nSaved {len(rows)} locations to {OUTPUT_CSV}  ({time.time() - t0:.0f}s)")
    if unparsed_rows:
        print(f"{len(unparsed_rows)} address-like lines could not be parsed -> {UNPARSED_CSV}")
    empty = [n for n in groups if not results.get(n, ([],))[0]]
    if empty:
        print("No locations found for: " + ", ".join(empty))


if __name__ == "__main__":
    main()
