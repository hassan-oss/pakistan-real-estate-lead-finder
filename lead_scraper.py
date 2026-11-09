#!/usr/bin/env python3
"""
Real-estate lead finder (Pakistan) v2  ->  leads.xlsx (import into Google Sheets)

    pip install -U requests beautifulsoup4 openpyxl ddgs
    python lead_finder_v2.py --target 100 --out leads.xlsx

No seed data needed. It searches every major Pakistani city (and big housing areas), harvests agency sites
from "top agencies" list pages, crawls each site in parallel, and keeps only companies where at most ONE of
(decision-maker, email, phone) is missing and every field has its own source link. Emails are never guessed.

Optional:  --cities Lahore Karachi      limit cities (default: 14 cities)
           --max-missing 0              require all three fields
           --min-confidence high        keep only High rows
           --google                     add Google Places (GOOGLE_API_KEY) - biggest candidate pool
           --serpapi-key KEY            fill missing decision-makers from public LinkedIn search snippets
           --js                         re-render JavaScript-only sites (pip install playwright; playwright install chromium)
"""
import argparse, base64, csv, json, os, re, sys, threading, time
import urllib.parse as up
import urllib.robotparser as rp
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import date

import requests
from bs4 import BeautifulSoup, NavigableString
from openpyxl import Workbook
from openpyxl.formatting.rule import FormulaRule
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation

UA = ("Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/124.0 Safari/537.36")
DELAY = 0.4
TIMEOUT = (5, 12)
MAX_PAGES = 14
SITE_BUDGET = 30      # max seconds spent crawling one company
SKIP_SITEMAP = False
USE_JS = False

# --------------------------------------------------------------------------- discovery config
DEFAULT_CITIES = ["Lahore", "Karachi", "Islamabad", "Rawalpindi", "Faisalabad", "Multan", "Peshawar",
                  "Gujranwala", "Sialkot", "Hyderabad", "Quetta", "Bahawalpur", "Sargodha", "Abbottabad"]
AREAS = {
    "Lahore": ["DHA", "Bahria Town", "Gulberg", "Johar Town", "Model Town", "Cantt", "Wapda Town", "Valencia"],
    "Karachi": ["DHA", "Clifton", "Gulshan-e-Iqbal", "North Nazimabad", "Bahria Town", "Scheme 33"],
    "Islamabad": ["F-7", "F-10", "E-11", "G-13", "DHA", "Bahria Enclave", "Blue Area", "B-17"],
    "Rawalpindi": ["Bahria Town", "Saddar", "DHA", "Chaklala"],
}
TEMPLATES = ["{niche}", "property dealers", "real estate developers", "real estate consultants", "estate agents",
             "property advisors", "real estate company contact us", "real estate agency about us team"]
LIST_HINT = re.compile(r"\b(top|best|list|directory|companies|agencies|agents|dealers|developers)\b", re.I)

EXCLUDE_DOMAINS = (
    "facebook.com", "instagram.com", "linkedin.com", "twitter.com", "x.com", "youtube.com", "zoominfo.com",
    "realmadrid.com", "zameen.com", "graana.com", "olx.com.pk", "lamudi.pk", "zameen.net", "pinterest.com",
    "wikipedia.org", "tiktok.com", "yellowpages.com.pk", "pakistanyellowpages.com", "tripadvisor.com", "yelp.com",
    "google.com", "bing.com", "duckduckgo.com", "realtor.com", "zillow.com", "tumblr.com", "medium.com",
    "reddit.com", "quora.com", "blogspot.com", "wordpress.org", "wordpress.com", "w3.org", "cloudflare.com",
    "apple.com", "microsoft.com", "whatsapp.com", "wa.me", "t.me", "snapchat.com", "github.com", "zhihu.com",
    "f6s.com", "ensun.io", "urdupoint.com", "wix.com", "godaddy.com", "bit.ly", "goo.gl", "maps.app.goo.gl")

# --------------------------------------------------------------------------- crawl config
PAGE_PRIORITY = [("ceo", 0), ("message", 0), ("founder", 0), ("owner", 0), ("chairman", 0), ("team", 1),
                 ("leadership", 1), ("management", 1), ("director", 1), ("board", 1), ("people", 1), ("staff", 1),
                 ("about", 2), ("who-we-are", 2), ("whoweare", 2), ("our-story", 2), ("contact", 2),
                 ("company", 3), ("profile", 3)]
FALLBACK_PATHS = ("/about", "/about-us", "/aboutus", "/team", "/our-team", "/management", "/leadership",
                  "/contact", "/contact-us", "/contactus", "/ceo-message", "/message-from-ceo", "/who-we-are",
                  "/company", "/our-company", "/about-company")
GENERIC_LOCALS = {
    "info", "contact", "sales", "admin", "support", "hello", "office", "enquiry", "inquiry", "enquiries",
    "inquiries", "mail", "marketing", "help", "team", "hr", "careers", "career", "service", "services",
    "customercare", "query", "accounts", "billing", "webmaster", "feedback", "booking", "bookings", "reception",
    "helpline", "helpdesk", "complaint", "complaints", "care", "general", "president", "presidencia", "press", "media", "pr", "business", "enquire", "contactus", "ask"}
BAD_EMAIL_PARTS = ("example.", "sentry", "wixpress", "domain.com", "yourdomain", "email.com", "yourname", "user@")
IMG_EXT = (".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js", ".woff")
FREE_MAIL = {"gmail.com", "yahoo.com", "hotmail.com", "outlook.com", "live.com", "icloud.com", "yahoo.co.uk", "ymail.com"}

TITLES = [  # priority order, best first
    ("CEO", r"chief executive officer|\bceo\b"),
    ("Founder", r"co-?founder|\bfounder\b"),
    ("Owner", r"\bowner\b|proprietor"),
    ("Managing Director", r"managing director|\bmd\b"),
    ("Chairman", r"chairman|chairperson"),
    ("President", r"\bpresident\b"),
    ("Director", r"\bdirector\b"),
    ("General Manager", r"general manager"),
    ("Sales Manager", r"sales manager|marketing manager|sales head|head of sales"),
    ("Manager", r"\bmanager\b"),
]
RE_KEYWORDS = ("real estate", "real-estate", "realestate", "realtor", "estate agen", "property", "properties",
               "housing", "plots", "estate consult")
STRONG_HEAD = ("real estate", "realestate", "real-estate", "realtor", "estate agen", "estate advis", "estate consult",
               "estate marketing", "property dealer", "property consult", "property advis", "property management",
               "properties", "developers", "builders", "housing", "plots", "estate")
SALE_WORDS = ("for sale", "buy", "sell", "rent", "plots", "apartments", "commercial", "residential", "marla",
              "kanal", "listing")
META_KEYS = ("description", "og:title", "og:description", "og:site_name", "keywords", "twitter:title")

PLACEHOLDER = {"john doe", "jane doe", "lorem ipsum", "your name", "first last", "name surname", "full name",
               "john smith", "jane smith", "demo user", "test user"}
NAME_BAD = set("""
team about contact our the estate real property properties message from meet us home services company group
agency realtors builders developers read more welcome chief executive officer managing director founder owner ceo
manager chairman and of for with get in touch call email phone who press highlights amenity amenities lands land
achievements vision mission projects project gallery news blog careers testimonials why choose awards board
leadership management founders city town valley park residency plaza heights homes associates pvt ltd limited
marketing consultants consultant solutions international enterprises investment investments sales support customer
service office head hr dealer dealers agents agent advisors advisor society housing phase block sector road street
lahore karachi islamabad rawalpindi faisalabad pakistan dha bahria gulberg note words speaks speech desk profile
biography bio story history values commitment quality trust excellence professional experienced expert experts best
top leading premium luxury new latest opportunity investors clients customers buy sell rent view all my your we are
is a an to by at on be do you can will talk inquire enquire now today free consultation schedule book visit site
dream partner happy started learn sign log privacy policy terms conditions follow quick links menu search subscribe
newsletter copyright rights reserved developed designed powered share apply join hiring career jobs portfolio
listing listings featured popular recent related similar other less next previous back skip main navigation footer
header director's ceo's chairman's president vice
ask anything personal information details enter name email submit send form select required optional first last full
address number mobile comment comments chat bot assistant help helpline question questions faq reply click here
download brochure brochures register login registration account cart wishlist compare filter sort price prices area
areas size bedrooms beds baths sqft marla kanal acre plot file files booking payment plan plans installments
installment possession location map video videos photos images tour virtual
""".split())
HONORIFICS = re.compile(r"^(?:(?:mr|mrs|ms|dr|engr|eng|sir|haji|hajji|prof|capt|col|maj|brig|gen|syed|sayed|sheikh|"
                        r"rana|ch|chaudhry|chaudhary|mian|khawaja|sardar|hafiz|malik|mirza|raja)\.?\s+)+", re.I)

EMAIL_RE = re.compile(r"[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)*\.[A-Za-z]{2,}")
PHONE_RE = re.compile(r"(?:\+|00)?\d[\d\s\-\(\)\.]{8,16}\d")
LANDLINE_AREA = {"21", "42", "51", "91", "41", "61", "52", "55", "71", "81", "22", "25", "99", "92", "44", "47",
                 "48", "49", "58", "53", "54", "56", "57", "62", "63", "64", "65", "66", "67", "68", "69", "72",
                 "74", "76", "77", "82", "83", "84", "85", "86", "87", "88", "89", "94", "95", "96", "97", "98"}

# --------------------------------------------------------------------------- http helpers
_tl = threading.local()
_robots = {}


def S():
    """Thread-local requests session."""
    if not hasattr(_tl, "s"):
        s = requests.Session()
        s.headers.update({"User-Agent": UA, "Accept-Language": "en",
                          "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"})
        _tl.s = s
    return _tl.s


def root_domain(url):
    host = up.urlparse(url if "//" in url else "//" + url).netloc.lower().split(":")[0]
    return host[4:] if host.startswith("www.") else host


def norm_url(url):
    url = (url or "").strip()
    if not url:
        return ""
    if not url.startswith("http"):
        url = "https://" + url
    p = up.urlparse(url)
    return f"{p.scheme}://{p.netloc}/"


NON_BUSINESS_SUFFIX = (".gov.pk", ".edu.pk", ".mil.pk", ".ac.pk", ".gov", ".edu", ".mil", ".gop.pk", ".gob.pk")


def excluded(dom):
    return (not dom or dom.endswith(NON_BUSINESS_SUFFIX)
            or any(dom == x or dom.endswith("." + x) for x in EXCLUDE_DOMAINS))


def allowed(url):
    base = norm_url(url)
    if base not in _robots:
        r = rp.RobotFileParser()
        try:
            resp = S().get(base + "robots.txt", timeout=(5, 10))
            r.parse(resp.text.splitlines() if resp.status_code == 200 else [])
        except Exception:
            r.parse([])
        _robots[base] = r
    return _robots[base].can_fetch("*", url)


def fetch(url, xml=False):
    try:
        if not allowed(url):
            return None
        time.sleep(DELAY)
        r = S().get(url, timeout=TIMEOUT)
        ct = r.headers.get("content-type", "text/html")
        if r.status_code == 200 and (xml or "html" in ct):
            return r.text[:2_500_000]
    except Exception:
        pass
    return None


def fetch_js(url):
    """Optional: render JavaScript pages with Playwright (only used with --js)."""
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return None
    try:
        with sync_playwright() as p:
            b = p.chromium.launch()
            pg = b.new_page(user_agent=UA)
            pg.goto(url, wait_until="networkidle", timeout=25000)
            html = pg.content()
            b.close()
            return html
    except Exception:
        return None


# --------------------------------------------------------------------------- extraction
def cf_decode(h):
    key = int(h[:2], 16)
    return "".join(chr(int(h[i:i + 2], 16) ^ key) for i in range(2, len(h), 2))


def extract_emails(soup):
    found = []
    for a in soup.select("a[href^='mailto:']"):
        found.append(up.unquote(a["href"][7:]).split("?")[0].strip())
    for el in soup.select("[data-cfemail]"):
        try:
            found.append(cf_decode(el["data-cfemail"]))
        except Exception:
            pass
    text = soup.get_text(" ")
    text = re.sub(r"\s*[\[\(\{]\s*at\s*[\]\)\}]\s*", "@", text, flags=re.I)       # name [at] domain
    text = re.sub(r"\s*[\[\(\{]\s*dot\s*[\]\)\}]\s*", ".", text, flags=re.I)       # domain [dot] com
    found += EMAIL_RE.findall(text)
    out = []
    for e in found:
        e = e.strip().strip(".,;:").lower()
        if (not EMAIL_RE.fullmatch(e) or e.endswith(IMG_EXT) or any(b in e for b in BAD_EMAIL_PARTS)):
            continue
        if e not in out:
            out.append(e)
    return out


def format_pk_phone(digits):
    if not digits.startswith("92"):
        return None
    rest = digits[2:]
    if rest.startswith("3") and len(rest) == 10:
        return f"+92 {rest[:3]} {rest[3:]}", "Mobile"
    if 9 <= len(rest) <= 10 and rest[:2] in LANDLINE_AREA:
        return f"+92 {rest[:2]} {rest[2:]}", "Landline"
    return None


def to_digits(raw):
    d = re.sub(r"\D", "", raw)
    if d.startswith("0092"):
        d = d[2:]
    elif d.startswith("0") and not d.startswith("00"):
        d = "92" + d[1:]
    return d


def extract_phones(soup):
    out = []

    def add(raw, force=None):
        res = format_pk_phone(to_digits(raw))
        if res and res[0] not in [o[0] for o in out]:
            out.append((res[0], force or res[1]))

    for a in soup.select("a[href^='tel:']"):
        add(up.unquote(a["href"][4:]))
    for a in soup.select("a[href*='wa.me'], a[href*='api.whatsapp.com'], a[href*='whatsapp.com/send']"):
        m = re.search(r"(?:wa\.me/|phone=)(\+?\d+)", a["href"])
        if m:
            add(m.group(1), "WhatsApp")
    for m in PHONE_RE.findall(soup.get_text(" ")):
        add(m)
    return out


def looks_like_name(s):
    s0 = re.sub(r"\s+", " ", s).strip(" -–—|,:•·\u00a0\"'")
    if s0.isupper():
        s0 = s0.title()
    core = HONORIFICS.sub("", s0).strip()
    if s0.lower() in PLACEHOLDER or core.lower() in PLACEHOLDER:
        return None
    words = core.split()
    if not (2 <= len(words) <= 4) or len(core) > 40:
        return None
    real = 0
    for w in words:
        if re.fullmatch(r"[A-Z]\.?", w):                      # initial like "M."
            continue
        if (not re.fullmatch(r"[A-Z][a-zA-Z'’\-]{2,}\.?", w)) or w.lower().strip(".'’") in NAME_BAD:
            return None
        real += 1
    return s0 if real >= 1 else None


NOT_OWNER = re.compile(r"\b(?:property|home|house|land|plot|flat|shop|business|apartment|unit|villa)\s+owners?\b", re.I)
FORMISH = re.compile(r"personal information|your name|enter your|submit|required field|select |full name|first name|"
                     r"last name|captcha|subscribe|ask anything|type your|your message|your email", re.I)


def titles_in(text):
    low = NOT_OWNER.sub(" ", text.lower())
    labels = [lab for lab, pat in TITLES if re.search(pat, low)]
    if "Managing Director" in labels:
        labels = [x for x in labels if x != "Director"]
    if any(x != "Manager" and x.endswith("Manager") for x in labels):
        labels = [x for x in labels if x != "Manager"]
    return labels


def title_of(text):
    low = NOT_OWNER.sub(" ", text.lower())
    for rank, (label, pat) in enumerate(TITLES):
        if re.search(pat, low):
            return rank, label
    return None


def jsonld_people(soup):
    out = []

    def walk(o):
        if isinstance(o, list):
            for x in o:
                walk(x)
        elif isinstance(o, dict):
            for key in ("founder", "founders"):
                v = o.get(key)
                for p in (v if isinstance(v, list) else [v] if v else []):
                    nm = p.get("name") if isinstance(p, dict) else p
                    n = looks_like_name(nm) if isinstance(nm, str) else None
                    if n:
                        out.append((1, n, "Founder", 0))
            ty = o.get("@type")
            if ty == "Person" or (isinstance(ty, list) and "Person" in ty):
                jt, nm = o.get("jobTitle"), o.get("name")
                t = title_of(jt) if isinstance(jt, str) else None
                n = looks_like_name(nm) if isinstance(nm, str) else None
                if n and t:
                    out.append((t[0], n, jt[:50], 0))
            for v in o.values():
                walk(v)

    for sc in soup.find_all("script", type="application/ld+json"):
        try:
            walk(json.loads(sc.string or sc.get_text()))
        except Exception:
            pass
    return out


TESTIMONIAL = re.compile(r"testimonial|what our clients|what clients|happy client|reviews?\b|customer says|clients say", re.I)


def extract_dms(soup):
    """Returns [(rank, name, title, distance)]. Scans text line by line: name on the same line as a title,
    or within 3 lines before/after it (handles cards, 'Message from the CEO', signature blocks)."""
    results = jsonld_people(soup)
    for t in soup(["script", "style", "noscript"]):
        t.decompose()
    for img in soup.find_all("img"):
        alt = (img.get("alt") or img.get("title") or "").strip()
        if alt and len(alt) < 80:
            img.insert_after(NavigableString("\n" + alt + "\n"))
    lines, prev = [], None
    for l in soup.get_text("\n").split("\n"):
        l = re.sub(r"\s+", " ", l).strip()
        if l and l != prev:
            lines.append(l)
        prev = l
    for i, l in enumerate(lines):
        if len(l) > 70 or len(l.split()) > 9:
            continue
        t = title_of(l)
        if not t:
            continue
        if any(TESTIMONIAL.search(x) or FORMISH.search(x) for x in lines[max(0, i - 3):i + 4]):
            continue
        rank, label = t
        name, dist = None, 0
        m = re.search("|".join(p for _, p in TITLES), l, re.I)
        parts = ([l[:m.start()], l[m.end():]] if m else []) + re.split(r"\s[-–—|/]\s|\s*[,:()]\s*", l)
        for p in parts:
            name = looks_like_name(p) if p.strip() else None
            if name:
                break
        if not name:
            for d in (-1, 1, -2, 2, -3, 3):
                j = i + d
                if 0 <= j < len(lines) and not title_of(lines[j]):
                    name = looks_like_name(lines[j])
                    if name:
                        dist = abs(d)
                        break
        if name:
            results.append((rank, name, " & ".join(titles_in(l)[:2]) or label, dist))
    return results


def email_matches_person(email, name):
    if not name:
        return False
    local = re.sub(r"[^a-z]", "", email.split("@")[0])
    toks = [t.lower() for t in re.findall(r"[A-Za-z]+", name) if len(t) > 2]
    return any(t in local for t in toks)


def email_type(email, site_domain=""):
    """Generic = role mailbox (info@, sales@...) or a mailbox named after the company.
    Personal = any other named mailbox."""
    if not email:
        return "Not found"
    local = email.split("@")[0]
    letters = re.sub(r"[^a-z]", "", local)
    if local in GENERIC_LOCALS or letters in GENERIC_LOCALS:
        return "Generic"
    label = re.sub(r"[^a-z]", "", site_domain.split(".")[0]) if site_domain else ""
    if label and len(letters) >= 4 and (letters in label or label in letters):
        return "Generic"
    return "Personal"


def email_belongs(email, site):
    d = email.split("@")[1]
    return d in FREE_MAIL or d == site or d.endswith("." + site) or site.endswith("." + d)


def page_priority(href, text=""):
    hay = (href + " " + text).lower()
    scores = [p for k, p in PAGE_PRIORITY if k in hay]
    return min(scores) if scores else None


def candidate_links(base, soup):
    links = {}
    for a in soup.find_all("a", href=True):
        href = up.urljoin(base, a["href"]).split("#")[0]
        if root_domain(href) != root_domain(base) or href.rsplit(".", 1)[-1].lower() in ("pdf", "jpg", "png", "zip"):
            continue
        sc = page_priority(a["href"], a.get_text(" "))
        if sc is not None and (href not in links or sc < links[href]):
            links[href] = sc
    return links


def sitemap_links(base):
    found = {}
    xml = fetch(base + "sitemap.xml", xml=True)
    if not xml:
        return found
    locs = re.findall(r"<loc>\s*(.*?)\s*</loc>", xml)
    children = [l for l in locs if l.endswith(".xml") and "page" in l.lower()][:2]
    for c in children:
        sub = fetch(c, xml=True)
        if sub:
            locs += re.findall(r"<loc>\s*(.*?)\s*</loc>", sub)
    for l in locs:
        if root_domain(l) == root_domain(base) and not l.endswith(".xml"):
            sc = page_priority(l)
            if sc is not None and sc <= 2:
                found[l] = sc
    return found


def clean_name(n, site):
    n = re.sub(r"^(welcome to|home\s*[-–|:]\s*)", "", n.strip(), flags=re.I)
    n = re.sub(r"\s*[-–—|:•]\s*(home|official website|official site)$", "", n, flags=re.I).strip(" -–—|:•,")
    if len(n) < 3 or n.lower() in ("property dealer", "real estate", "home", "welcome", "property", "estate agent"):
        n = re.sub(r"[^a-z0-9]", "", site.split(".")[0]).title()
    return n


def head_text(soup):
    bits = [soup.title.get_text(" ") if soup.title else ""]
    for m in soup.find_all("meta"):
        if (m.get("name") or m.get("property") or "").lower() in META_KEYS:
            bits.append(m.get("content", ""))
    bits += [h.get_text(" ") for h in soup.find_all(["h1", "h2"])]
    return " ".join(bits).lower()


def company_name_from(soup, base):
    dom_label = re.sub(r"[^a-z]", "", root_domain(base).split(".")[0])
    cands = []
    for m in soup.find_all("meta"):
        if (m.get("property") or "").lower() == "og:site_name" and m.get("content", "").strip():
            cands.append(m["content"].strip())
    if soup.title:
        cands.append(soup.title.get_text(strip=True))
    for c in cands:
        segs = [x.strip() for x in re.split(r"\s*(?:\|\||[|–—•:]|\s-\s)\s*", c) if x.strip()]
        for seg in segs:                                           # segment that matches the domain wins
            letters = re.sub(r"[^a-z]", "", seg.lower())
            if len(letters) >= 4 and (letters in dom_label or dom_label in letters):
                return seg[:60]
    for c in cands:
        first = re.split(r"\s*(?:\|\||[|–—•:]|\s-\s)\s*", c)[0].strip()
        if first and not re.search(r"real estate|property|construction|company|best|top|leading|pakistan", first, re.I):
            return first[:60]
    return dom_label.title() or root_domain(base)


# --------------------------------------------------------------------------- crawl one company
COLS = ["Company Name", "Website URL", "City", "Homepage Title", "Decision-Maker Name", "Title", "Email", "Email Type",
        "Email Matches DM", "Phone",
        "Phone Type", "Source Link: Email", "Source Link: Phone", "Source Link: Decision-Maker",
        "Other Emails Found", "Date Verified", "Confidence", "Notes", "LinkedIn Search (manual check)"]


def harvest(pages, site):
    emails, phones, dms = [], [], []
    for url, soup in pages.items():
        for e in extract_emails(soup):
            if email_belongs(e, site) and e not in [x[0] for x in emails]:
                emails.append((e, url))
        for p in extract_phones(soup):
            if p[0] not in [x[0][0] for x in phones]:
                phones.append((p, url))
        for rank, name, title, dist in extract_dms(soup):
            dms.append((rank, name, title, url, dist))
    return emails, phones, dms


def crawl_company(company, website, city=""):
    """Row dict, or None if unreachable / not a Pakistani real-estate business."""
    base = norm_url(website)
    if not base or excluded(root_domain(base)):
        return None
    site = root_domain(base)
    t0 = time.time()
    home = fetch(base)
    if home is None:
        print(f"   x {site}: unreachable / blocked")
        return None
    pages = {base: BeautifulSoup(home, "html.parser")}
    links = candidate_links(base, pages[base])
    if not SKIP_SITEMAP:
        for u, sc in sitemap_links(base).items():
            links.setdefault(u, sc)
    for path in FALLBACK_PATHS:
        links.setdefault(up.urljoin(base, path), page_priority(path) or 4)
    order = sorted(links, key=lambda u: links[u])
    for u in order:
        if len(pages) >= MAX_PAGES or time.time() - t0 > SITE_BUDGET:
            break
        h = fetch(u)
        if h:
            pages[u] = BeautifulSoup(h, "html.parser")

    head = head_text(pages[base])                                   # relevance gate (before soups are modified)
    body = " ".join(p.get_text(" ", strip=True).lower() for p in pages.values())
    hits = sum(body.count(k) for k in RE_KEYWORDS)
    sale = sum(body.count(w) for w in SALE_WORDS)
    if excluded(site) or not any(k in head for k in STRONG_HEAD) or hits < 8 or sale < 5:
        print(f"   x {site}: not a real-estate site")
        return None
    home_title = (pages[base].title.get_text(" ", strip=True) if pages[base].title else "")[:90]
    name = clean_name(company or company_name_from(pages[base], base), site)

    emails, phones, dms = harvest(pages, site)
    if USE_JS and (not dms or not emails):                          # optional JavaScript re-render
        js_pages = {}
        for u in [base] + order[:4]:
            h = fetch_js(u)
            if h:
                js_pages[u + "#js"] = BeautifulSoup(h, "html.parser")
        e2, p2, d2 = harvest(js_pages, site)
        emails += [x for x in e2 if x[0] not in [y[0] for y in emails]]
        phones += [x for x in p2 if x[0][0] not in [y[0][0] for y in phones]]
        dms += [(r, n, t, u.replace("#js", ""), d) for r, n, t, u, d in d2]

    if not (site.endswith(".pk") or phones or "pakistan" in body):
        print(f"   x {site}: no Pakistan indicators")
        return None

    row = {k: "" for k in COLS}
    row.update({"Company Name": name, "Website URL": base, "City": city, "Homepage Title": home_title,
                "Email Type": "Not found",
                "Date Verified": date.today().isoformat()})

    # decision-maker: best title, prefer leadership/about pages over the home page (testimonials live there)
    dms.sort(key=lambda d: (d[0] + (3 if d[3] == base else 0), d[4]))
    dm = dms[0] if dms else None
    if dm:
        row["Decision-Maker Name"], row["Title"], row["Source Link: Decision-Maker"] = dm[1], dm[2], dm[3]

    if emails:
        def erank(eu):
            e = eu[0]
            if dm and email_matches_person(e, dm[1]) and email_type(e, site) == "Personal":
                return 0
            return 1 if email_type(e, site) == "Personal" else 2
        ranked = sorted(emails, key=erank)
        e, eurl = ranked[0]
        row["Email"], row["Source Link: Email"], row["Email Type"] = e, eurl, email_type(e, site)
        row["Other Emails Found"] = " | ".join(f"{x} ({email_type(x, site)})" for x, _ in ranked[1:6])

    if phones:
        mobiles = [x for x in phones if x[0][1] in ("Mobile", "WhatsApp")]
        (ph, ptype), purl = (mobiles or phones)[0]
        row["Phone"], row["Phone Type"], row["Source Link: Phone"] = ph, ptype, purl
    return finish_row(row)


def missing_count(row):
    return sum(1 for k in ("Decision-Maker Name", "Email", "Phone") if not row[k])


def sources_ok(row):
    pairs = (("Decision-Maker Name", "Source Link: Decision-Maker"), ("Email", "Source Link: Email"),
             ("Phone", "Source Link: Phone"))
    return all(row[s] for f, s in pairs if row[f])


def finish_row(row):
    em, dmn = row["Email"], row["Decision-Maker Name"]
    row["Email Matches DM"] = ("Yes" if (dmn and email_matches_person(em, dmn)) else "No") if em else ""
    miss = missing_count(row)
    if miss == 0 and row["Email Matches DM"] == "Yes":
        row["Confidence"] = "High"
    elif miss <= 1:
        row["Confidence"] = "Medium"
    else:
        row["Confidence"] = "Low"
    q = up.quote(f'site:linkedin.com "{row["Company Name"]}" CEO OR Founder OR Owner')
    row["LinkedIn Search (manual check)"] = f"https://www.google.com/search?q={q}"
    return row


# --------------------------------------------------------------------------- discovery
SEARCH_BACKENDS = ["duckduckgo", "bing", "brave", "yahoo", "mojeek", "google"]
_bad_backend = {}
_warned = set()


def warn_once(key, msg):
    if key not in _warned:
        _warned.add(key)
        print("  !", msg)


def _bing_unwrap(href):
    if "bing.com/ck/a" in href:
        m = re.search(r"[?&]u=a1([A-Za-z0-9_\-]+)", href)
        if m:
            b = m.group(1)
            try:
                return base64.urlsafe_b64decode(b + "=" * (-len(b) % 4)).decode()
            except Exception:
                return ""
    return href


def search(query, n=30):
    """[(url, title)] from ddgs (several engines), falling back to Bing HTML."""
    out, productive = [], 0
    DDGS = None
    for mod in ("ddgs", "duckduckgo_search"):
        try:
            DDGS = __import__(mod).DDGS
            break
        except Exception:
            continue
    if DDGS:
        for be in SEARCH_BACKENDS:
            if _bad_backend.get(be, 0) >= 3:
                continue
            try:
                try:
                    res = DDGS(timeout=10).text(query, max_results=n, backend=be)
                except TypeError:
                    res = DDGS().text(query, max_results=n)
                got = [(r.get("href") or r.get("url"), r.get("title", "")) for r in (res or []) if r]
                if got:
                    out += got
                    _bad_backend[be] = 0
                    productive += 1
                    if productive >= 2:
                        break
                else:
                    _bad_backend[be] = _bad_backend.get(be, 0) + 1
            except Exception as e:
                _bad_backend[be] = _bad_backend.get(be, 0) + 1
                warn_once(be, f"search engine '{be}' failing ({type(e).__name__}); will skip it if it keeps failing")
            time.sleep(0.7)
    if not out:                                                      # Bing HTML fallback
        try:
            time.sleep(1.5)
            r = S().get("https://www.bing.com/search", params={"q": query, "count": 30, "cc": "pk"}, timeout=TIMEOUT)
            for a in BeautifulSoup(r.text, "html.parser").select("li.b_algo h2 a"):
                out.append((_bing_unwrap(a["href"]), a.get_text(" ")))
        except Exception as e:
            warn_once("bing", f"Bing fallback failed: {type(e).__name__}")
    return out


def harvest_links(url):
    """Collect outbound website links from a 'top agencies' list page."""
    html = fetch(url)
    if not html:
        return []
    soup = BeautifulSoup(html, "html.parser")
    me, seen, out = root_domain(url), set(), []
    for a in soup.find_all("a", href=True):
        href = up.urljoin(url, a["href"])
        dom = root_domain(href)
        if not href.startswith("http") or dom == me or excluded(dom) or dom in seen:
            continue
        seen.add(dom)
        out.append(norm_url(href))
    return out[:60]


def discover_web(niche, city, max_queries):
    queries = [f"{t.format(niche=niche)} {city} Pakistan" for t in TEMPLATES]
    for area in AREAS.get(city, []):
        queries += [f"{niche} {area} {city}", f"property dealers {area} {city}"]
    out, seen = [], set()
    for i, q in enumerate(queries[:max_queries], 1):
        res = search(q)
        added = 0
        for url, title in res:
            if not url:
                continue
            dom = root_domain(url)
            deep = up.urlparse(url).path.strip("/") != ""
            if (deep or excluded(dom)) and LIST_HINT.search(title or "") and not any(x in dom for x in ("facebook", "linkedin", "youtube", "google")):
                for w in harvest_links(url):
                    if root_domain(w) not in seen:
                        seen.add(root_domain(w))
                        out.append({"company": "", "website": w, "city": city})
                        added += 1
            if excluded(dom) or dom in seen:
                continue
            seen.add(dom)
            out.append({"company": "", "website": norm_url(url), "city": city})
            added += 1
        print(f"  [{city}] query {i}/{min(len(queries), max_queries)}: +{added} sites ({q})")
    return out


def discover_places(city, queries):
    key = os.environ.get("GOOGLE_API_KEY")
    if not key:
        sys.exit("Set GOOGLE_API_KEY to use --google.")
    mask = "places.displayName,places.websiteUri,nextPageToken"
    out = []
    for q in queries:
        token = None
        for _ in range(3):
            body = {"textQuery": f"{q} in {city}, Pakistan", "pageSize": 20}
            if token:
                body["pageToken"] = token
            r = S().post("https://places.googleapis.com/v1/places:searchText", json=body, timeout=TIMEOUT,
                         headers={"X-Goog-Api-Key": key, "X-Goog-FieldMask": mask})
            if r.status_code != 200:
                print("  Places API error:", r.status_code, r.text[:150])
                break
            data = r.json()
            for p in data.get("places", []):
                out.append({"company": p["displayName"]["text"], "website": p.get("websiteUri", ""), "city": city})
            token = data.get("nextPageToken")
            if not token:
                break
            time.sleep(2)
    return out


def discover_osm(niche, city):
    q = (f'[out:json][timeout:50];area["name"="{city}"]["boundary"="administrative"]->.a;'
         '(nwr["office"="estate_agent"](area.a);nwr["shop"="estate_agent"](area.a););out center tags;')
    for url in ("https://overpass-api.de/api/interpreter", "https://overpass.kumi.systems/api/interpreter"):
        try:
            r = S().post(url, data={"data": q}, timeout=(10, 60), headers={"User-Agent": "lead-research-script/1.0"})
            if r.status_code == 200 and r.text.strip().startswith("{"):
                return [{"company": e["tags"]["name"], "city": city,
                         "website": e["tags"].get("website") or e["tags"].get("contact:website", "")}
                        for e in r.json().get("elements", []) if e.get("tags", {}).get("name")]
        except Exception:
            continue
    return []


def linkedin_lookup(company, serp_key):
    try:
        r = S().get("https://serpapi.com/search.json", timeout=TIMEOUT, params={
            "engine": "google", "api_key": serp_key, "num": 5,
            "q": f'site:linkedin.com/in "{company}" (CEO OR Founder OR Owner OR "Managing Director")'})
        results = r.json().get("organic_results", [])
    except Exception:
        return None
    key = (re.sub(r"[^a-z0-9 ]", "", company.lower()).split() or [""])[0]
    for res in results:
        title = res.get("title", "").replace("| LinkedIn", "").strip()
        parts = [p.strip() for p in re.split(r"\s[-–|]\s", title)]
        name = looks_like_name(parts[0]) if parts else None
        blob = (title + " " + res.get("snippet", "")).lower()
        t = title_of(blob)
        if name and t and key and key in blob and "linkedin.com/in" in res.get("link", ""):
            return name, t[1], res["link"]
    return None


def read_seeds(path):
    with open(path, newline="", encoding="utf-8-sig") as f:
        return [{"company": r.get("company", ""), "website": r.get("website", ""), "city": r.get("city", "")}
                for r in csv.DictReader(f)]


# --------------------------------------------------------------------------- output
WIDTHS = {"Company Name": 28, "Website URL": 30, "City": 12, "Homepage Title": 36, "Email Matches DM": 12, "Decision-Maker Name": 24, "Title": 20, "Email": 30,
          "Email Type": 12, "Phone": 18, "Phone Type": 12, "Source Link: Email": 38, "Source Link: Phone": 38,
          "Source Link: Decision-Maker": 38, "Other Emails Found": 40, "Date Verified": 13, "Confidence": 12,
          "Notes": 36, "LinkedIn Search (manual check)": 30}


def col(name):
    return get_column_letter(COLS.index(name) + 1)


def write_xlsx(rows, path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Leads"
    ws.append(COLS)
    for r in rows:
        ws.append([r.get(c, "") for c in COLS])
    n = max(len(rows) + 1, 2)
    for row in ws.iter_rows(min_row=1, max_row=n):
        for c in row:
            c.font = Font(name="Arial", size=10, bold=(c.row == 1), color="FFFFFF" if c.row == 1 else "000000")
            c.alignment = Alignment(vertical="top")
            if c.row == 1:
                c.fill = PatternFill("solid", fgColor="1F4E78")
    for name, w in WIDTHS.items():
        ws.column_dimensions[col(name)].width = w
    ws.freeze_panes = "B2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(COLS))}{n}"
    fill = lambda c: PatternFill("solid", bgColor=c, fgColor=c)
    et, cf, web = col("Email Type"), col("Confidence"), col("Website URL")
    ws.conditional_formatting.add(f"{col('Email')}2:{et}{n}", FormulaRule(formula=[f'${et}2="Generic"'], fill=fill("FFF2CC")))
    ws.conditional_formatting.add(f"A2:A{n}", FormulaRule(formula=[f'AND(${web}2<>"",COUNTIF(${web}$2:${web}$1000,${web}2)>1)'], fill=fill("F8CBAD")))
    ws.conditional_formatting.add(f"{cf}2:{cf}{n}", FormulaRule(formula=[f'${cf}2="Medium"'], fill=fill("FCE4D6")))
    ws.conditional_formatting.add(f"{cf}2:{cf}{n}", FormulaRule(formula=[f'${cf}2="High"'], fill=fill("C6EFCE")))
    for name, opts in (("Email Type", '"Personal,Generic,Not found"'), ("Email Matches DM", '"Yes,No"'), ("Phone Type", '"Mobile,Landline,WhatsApp"'),
                       ("Confidence", '"High,Medium,Low"')):
        dv = DataValidation(type="list", formula1=opts, allow_blank=True)
        ws.add_data_validation(dv)
        dv.add(f"{col(name)}2:{col(name)}1000")

    s = wb.create_sheet("Summary")
    s["A1"], s["A1"].font = "Delivery summary", Font(name="Arial", size=12, bold=True)
    rng = lambda name: f"Leads!{col(name)}2:{col(name)}1000"
    items = [
        ("Total companies", f"=COUNTA({rng('Company Name')})"),
        ("With named decision-maker", f'=COUNTIF({rng("Decision-Maker Name")},"?*")'),
        ("Personal (named) emails", f'=COUNTIF({rng("Email Type")},"Personal")'),
        ("Generic emails (flagged)", f'=COUNTIF({rng("Email Type")},"Generic")'),
        ("No email", f'=COUNTIF({rng("Email Type")},"Not found")'),
        ("With phone", f'=COUNTIF({rng("Phone")},"?*")'),
        ("Email matches decision-maker's name", f'=COUNTIF({rng("Email Matches DM")},"Yes")'),
        ("High confidence", f'=COUNTIF({rng("Confidence")},"High")'),
        ("Medium confidence", f'=COUNTIF({rng("Confidence")},"Medium")'),
        ("Emails with source link", f'=COUNTIF({rng("Source Link: Email")},"?*")'),
    ]
    for i, (k, f) in enumerate(items, 3):
        s.cell(i, 1, k).font = Font(name="Arial", size=10)
        s.cell(i, 2, f).font = Font(name="Arial", size=10)
    s.column_dimensions["A"].width = 38
    s.cell(len(items) + 4, 1, "Every field is copied from its linked source page; emails are never guessed. "
           "Rows with 2+ of (decision-maker, email, phone) missing are excluded.").font = Font(name="Arial", size=9, italic=True)
    wb.save(path)


# --------------------------------------------------------------------------- main
def iter_batches(a):
    if a.seeds:
        yield read_seeds(a.seeds)
    for city in (a.cities or DEFAULT_CITIES):
        print(f"\n=== Discovering: {city} ===")
        batch = []
        if a.osm:
            batch += discover_osm(a.niche, city)
        if a.google:
            batch += discover_places(city, [a.niche, "property dealer", "real estate developers"])
        if not a.no_web:
            batch += discover_web(a.niche, city, a.max_queries)
        yield batch


def main():
    global USE_JS, MAX_PAGES, DELAY, SKIP_SITEMAP
    ap = argparse.ArgumentParser()
    ap.add_argument("--niche", default="real estate agency")
    ap.add_argument("--cities", nargs="*", default=[], help="default: 14 major Pakistani cities")
    ap.add_argument("--target", type=int, default=100)
    ap.add_argument("--out", default="leads.xlsx")
    ap.add_argument("--max-missing", type=int, default=1, choices=[0, 1])
    ap.add_argument("--min-confidence", default="medium", choices=["high", "medium"])
    ap.add_argument("--workers", type=int, default=8, help="parallel site crawlers")
    ap.add_argument("--max-queries", type=int, default=12, help="search queries per city")
    ap.add_argument("--fast", action="store_true", help="6 pages/site, no sitemap, 16 workers, 6 queries/city")
    ap.add_argument("--seeds")
    ap.add_argument("--osm", action="store_true")
    ap.add_argument("--google", action="store_true")
    ap.add_argument("--no-web", action="store_true")
    ap.add_argument("--js", action="store_true")
    ap.add_argument("--serpapi-key", default=os.environ.get("SERPAPI_KEY"))
    a = ap.parse_args()
    USE_JS = a.js
    if a.fast:
        MAX_PAGES, DELAY, SKIP_SITEMAP = 6, 0.15, True
        a.workers, a.max_queries = max(a.workers, 16), min(a.max_queries, 6)

    rows, seen_dom, seen_phone, seen_name, seen_email = [], set(), set(), set(), set()
    rejected = 0
    with ThreadPoolExecutor(max_workers=a.workers) as ex:
        for batch in iter_batches(a):
            fresh = []
            for c in batch:
                dom = root_domain(c["website"]) if c.get("website") else ""
                if not dom or excluded(dom) or dom in seen_dom:
                    continue
                seen_dom.add(dom)
                fresh.append(c)
            fresh.sort(key=lambda c: 0 if root_domain(c["website"]).endswith(".pk") else 1)
            print(f"-- {len(fresh)} new candidate sites; crawling with {a.workers} workers "
                  f"({len(rows)}/{a.target} kept so far)")
            futs = {ex.submit(crawl_company, c["company"], c["website"], c.get("city", "")): c for c in fresh}
            for f in as_completed(futs):
                try:
                    row = f.result()
                except Exception as e:
                    print("   x crawl error:", type(e).__name__, str(e)[:80])
                    row = None
                if row is None:
                    rejected += 1
                    continue
                if a.serpapi_key and not row["Decision-Maker Name"]:
                    li = linkedin_lookup(row["Company Name"], a.serpapi_key)
                    if li:
                        row["Decision-Maker Name"], row["Title"], row["Source Link: Decision-Maker"] = li
                        row["Notes"] = "DM from public LinkedIn snippet; verify"
                        row = finish_row(row)
                nm = re.sub(r"[^a-z0-9]", "", row["Company Name"].lower())
                ph = re.sub(r"\D", "", row["Phone"])
                dup = nm in seen_name or (ph and ph in seen_phone) or (row["Email"] and row["Email"] in seen_email)
                ok = (missing_count(row) <= a.max_missing and sources_ok(row)
                      and (row["Confidence"] == "High" or a.min_confidence == "medium"))
                if dup or not ok:
                    rejected += 1
                    print(f"   - rejected {row['Company Name']}: {'duplicate' if dup else str(missing_count(row)) + '/3 fields missing'}")
                    continue
                seen_name.add(nm)
                if ph:
                    seen_phone.add(ph)
                if row["Email"]:
                    seen_email.add(row["Email"])
                rows.append(row)
                write_xlsx(rows, a.out)                              # file is always up to date
                print(f"   + KEPT {len(rows)}/{a.target} [{row['Confidence']}] {row['Company Name']} | "
                      f"{row['Decision-Maker Name'] or '-'} | {row['Email'] or '-'} | {row['Phone'] or '-'}")
                if len(rows) >= a.target:
                    for x in futs:
                        x.cancel()
                    break
            write_xlsx(rows, a.out)                                  # save progress after every batch
            if len(rows) >= a.target:
                break
    write_xlsx(rows, a.out)
    print(f"\nWrote {len(rows)} qualified companies to {a.out} ({rejected} candidates rejected).")
    if len(rows) < a.target:
        print("Below target: try --google, --js, or --max-missing 1 (default).")


if __name__ == "__main__":
    main()