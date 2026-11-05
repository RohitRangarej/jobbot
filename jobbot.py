#!/usr/bin/env python3
"""jobbot: polls company job feeds and job boards, scores openings against the
profile in config.json, and logs every match to a per-day sheet in an Excel
workbook. Each region in config.json gets its own workbook.

    python3 jobbot.py start    run in the background until stopped
    python3 jobbot.py stop     stop the background run
    python3 jobbot.py status   show whether it is running and what it has found
    python3 jobbot.py once     run a single pass in the foreground
    python3 jobbot.py setup-email   connect the Gmail inbox that receives job alerts
    python3 jobbot.py email-check   show what it can read from those alerts, saving nothing

On first run, config.example.json is copied to config.json; edit that copy.
It reads public job listings and, once connected, job-alert emails (read-only: messages
are never changed or marked as read). It does not submit applications.
Standard library only; nothing to install.
"""
import email
import email.utils
import hashlib
import html
import imaplib
import json
import os
import re
import shutil
import signal
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
import zipfile
from concurrent.futures import ThreadPoolExecutor
from datetime import date, datetime, timezone
from email.header import decode_header
from html.parser import HTMLParser
from urllib.parse import quote
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

HERE = os.path.dirname(os.path.abspath(__file__))
CONFIG = os.path.join(HERE, "config.json")
EXAMPLE_CONFIG = os.path.join(HERE, "config.example.json")
STATE = os.path.join(HERE, ".jobbot_state.json")
PIDFILE = os.path.join(HERE, ".jobbot.pid")
LOGFILE = os.path.join(HERE, "jobbot.log")
EMAIL_FILE = os.path.join(HERE, ".jobbot_email")  # the Gmail address only; never the password
KEYCHAIN_SERVICE = "jobbot-gmail"  # macOS Keychain entry that holds the app password
PASSWORD_ENV = "JOBBOT_EMAIL_PASSWORD"  # alternative to Keychain, and the only option off macOS

HEADERS = ["Found at", "Posted", "Company", "Role", "Location", "Fit", "Experience asked", "Salary",
           "Why it matched", "Apply link", "Status", "Notes"]
WIDTHS = [17, 18, 18, 44, 28, 6, 16, 18, 48, 60, 14, 30]
PERMIT_HEADER, PERMIT_WIDTH = "Work-permit notes", 46  # extra column for regions with work_permit
STATUSES = ["To apply", "Applied", "Interview", "Offer", "Rejected", "Not interested"]
BOARDS_SHEET, TODAY_SHEET, FOLLOW_SHEET = "Job boards", "Apply today", "Follow up"
STRETCH_SHEET = "Stretch roles"
TODAY_ROWS, FOLLOW_UP_DAYS = 15, 7
DETAILS = {}  # SmartRecruiters posting details, fetched once per posting while the bot runs


def log(msg):
    print(f"{datetime.now():%Y-%m-%d %H:%M:%S}  {msg}", flush=True)


def load_json(path, default):
    try:
        with open(path) as f:
            return json.load(f)
    except (OSError, ValueError):
        return default


def load_config():
    if not os.path.exists(CONFIG) and os.path.exists(EXAMPLE_CONFIG):
        shutil.copyfile(EXAMPLE_CONFIG, CONFIG)
        log("created config.json from config.example.json - edit it to describe the jobs you want")
    with open(CONFIG) as f:
        return json.load(f)


def utc_day(seconds):
    """YYYY-MM-DD for a Unix timestamp in seconds."""
    return datetime.fromtimestamp(seconds, timezone.utc).date().isoformat()


def load_state(cfg):
    """Matches found so far, as {region key: {job id: job}}."""
    state = load_json(STATE, {})
    if "jobs" in state:  # written before regions existed: all of it belongs to the first region
        state = {cfg["regions"][0]["key"]: state["jobs"]}
    for region in cfg["regions"]:
        state.setdefault(region["key"], {})
    return state


def save_json(path, data):
    tmp = path + ".tmp"
    with open(tmp, "w") as f:
        json.dump(data, f, indent=1)
    os.replace(tmp, path)


# ---------------------------------------------------------------- feeds

def fetch(url, as_json=True, tries=3):
    """GET a URL, retrying temporary failures (server errors, rate limits, timeouts, resets)."""
    for attempt in range(tries):
        try:
            req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0 jobbot"})
            with urllib.request.urlopen(req, timeout=30) as r:
                return json.load(r) if as_json else r.read().decode("utf-8", "replace")
        except urllib.error.HTTPError as e:
            if (e.code < 500 and e.code != 429) or attempt == tries - 1:
                raise
        except OSError:  # timeouts, connection resets, DNS
            if attempt == tries - 1:
                raise
        time.sleep(5 * (attempt + 1))


def strip_html(s):
    return re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", html.unescape(s or "")))


def jobs_from(feed, wanted=()):
    """Return normalised jobs for one feed: id, company, title, location, url, text, posted.
    wanted: location keywords of all regions, used to skip costly detail requests."""
    ats, slug = feed["ats"], feed["slug"]
    out = []

    def add(i, title, locations, url, text, company=None, **extra):
        out.append({"id": f"{ats}:{slug}:{i}", "company": company or feed["name"],
                    "title": title.strip(), "url": url, "text": text,
                    "location": " / ".join(dict.fromkeys(l.strip() for l in locations if l)),
                    **extra})

    if ats == "greenhouse":
        data = fetch(f"https://boards-api.greenhouse.io/v1/boards/{slug}/jobs?content=true")
        for j in data.get("jobs", []):
            add(j["id"], j["title"], [j["location"]["name"]], j["absolute_url"],
                strip_html(j.get("content")), posted=(j.get("first_published") or "")[:10])
    elif ats == "lever":
        for j in fetch(f"https://api.lever.co/v0/postings/{slug}?mode=json"):
            cat = j["categories"]
            text = " ".join([j.get("descriptionPlain") or "", j.get("additionalPlain") or ""] +
                            [l.get("text", "") + " " + strip_html(l.get("content"))
                             for l in j.get("lists", [])])
            add(j["id"], j["text"], [cat.get("location")] + (cat.get("allLocations") or []),
                j["hostedUrl"], text,
                posted=utc_day(j["createdAt"] / 1000) if j.get("createdAt") else "")
    elif ats == "ashby":
        data = fetch(f"https://api.ashbyhq.com/posting-api/job-board/{slug}")
        for j in data.get("jobs", []):
            add(j["id"], j["title"],
                [j.get("location")] + [s.get("location") for s in j.get("secondaryLocations") or []],
                j["jobUrl"], j.get("descriptionPlain") or "", posted=(j.get("publishedAt") or "")[:10])
    elif ats == "kalibrr":  # job board (Indonesia, Philippines); slug is the city to search
        country = quote(feed.get("country", "Indonesia"))
        offset = 0
        while offset < 1500:
            data = fetch(f"https://www.kalibrr.com/kjs/job_board/search?limit=100&text=&country={country}"
                         f"&location={quote(slug)}&offset={offset}")
            if data.get("from_alternative"):  # nothing in this city; these are from elsewhere
                break
            for j in data["jobs"]:
                place = (j.get("google_location") or {}).get("address_components") or {}
                add(j["id"], j["name"],
                    [", ".join(place[k] for k in ("city", "region", "country") if place.get(k))],
                    f"https://www.kalibrr.com/c/{j['company']['code']}/jobs/{j['id']}/{j['slug']}",
                    strip_html((j.get("description") or "") + " " + (j.get("qualifications") or "")),
                    j["company"]["name"], posted=(j.get("activation_date") or "")[:10])
            offset += len(data["jobs"])
            if not data["jobs"] or offset >= data.get("count", 0):
                break
    elif ats == "dealls":  # Indonesian job board; slug is its city id. Listings carry skills, not a description
        for page in range(1, 11):
            data = fetch("https://api.sejutacita.id/v1/explore-job/job?limit=50&published=true"
                         "&status=active&sortParam=mostRelevant&sortBy=asc&search="
                         f"&cityIds={slug}&page={page}")["data"]
            for j in data["docs"]:
                add(j["id"], j["role"],
                    [", ".join(x["name"] for x in (j.get("city"), j.get("country")) if x and x.get("name"))],
                    f"https://dealls.com/loker/{j['slug']}~{j['company'].get('slug', '')}",
                    ", ".join(s.get("name", "") for s in j.get("skills") or []),
                    j["company"]["name"], posted=(j.get("publishedAt") or "")[:10])
            if page >= data.get("totalPages", 1):
                break
    elif ats == "cutshort":  # Indian job board; slug is a listing page, newest jobs first
        page = re.search(r'<script id="__NEXT_DATA__"[^>]*>(.*?)</script>',
                         fetch(f"https://cutshort.io/jobs/{slug}", as_json=False), re.S)
        if not page:
            raise ValueError("listing page has no job data; its layout may have changed")
        data = json.loads(page.group(1))["props"]["pageProps"]["dehydratedState"]["queries"][0]
        for j in (data["state"]["data"]["data"].get("pageData") or {}).get("jobs") or []:
            company = (j.get("companyDetails") or {}).get("name")
            agency = (j.get("companyId") or {}).get("name")
            pay = j.get("salaryRange") or {}
            remote = ["Remote (India)"] if j.get("remoteType") in ("remote_only", "remote_okay") else []
            add(j["_id"], j["headline"], (j.get("locations") or []) + remote, j["publicUrl"],
                strip_html(j.get("sanitizedComment")) + " " + ", ".join(j.get("allSkills") or []),
                f"{company} (via {agency})" if j.get("hiringForClient") and company and agency
                else company or agency,
                id=f"cutshort:{j['_id']}",  # the same job appears on several listing pages
                years=(j.get("expRange") or {}).get("min"),
                salary=j.get("salaryRangeText") or "",
                posted=((j.get("jobFactSummary") or {}).get("postedDate") or "")[:10],
                salary_max_lpa=pay["maxVanity"] / 100000
                if pay.get("currency") == "INR" and pay.get("maxVanity") else None)
    elif ats == "smartrecruiters":
        offset = 0
        while offset < 2000:
            data = fetch(f"https://api.smartrecruiters.com/v1/companies/{slug}/postings"
                         f"?limit=100&offset={offset}")
            for j in data["content"]:
                where = (j.get("location") or {}).get("fullLocation") or ""
                if wanted and not any(k in where.lower() for k in wanted):
                    continue
                if j["id"] not in DETAILS:
                    DETAILS[j["id"]] = fetch(j["ref"])
                d = DETAILS[j["id"]]
                sections = ((d.get("jobAd") or {}).get("sections") or {}).values()
                add(j["id"], j["name"], [where], d.get("postingUrl") or j["ref"],
                    strip_html(" ".join((s or {}).get("text") or "" for s in sections)),
                    (j.get("company") or {}).get("name"), posted=(j.get("releasedDate") or "")[:10])
            offset += len(data["content"])
            if not data["content"] or offset >= data.get("totalFound", 0):
                break
    elif ats == "workable":  # listings carry no description, only department and level
        for j in fetch(f"https://apply.workable.com/api/v1/widget/accounts/{slug}")["jobs"]:
            places = [", ".join(x for x in (l.get("city"), l.get("country")) if x)
                      for l in j.get("locations") or []]
            add(j["shortcode"], j["title"], places or [f"{j.get('city') or ''}, {j.get('country') or ''}"],
                j["url"], " ".join(str(j.get(k) or "") for k in ("department", "function", "experience")),
                posted=j.get("published_on") or "")
    elif ats == "himalayas":  # remote board; slug is a search term, optional country keeps roles open there
        country = feed.get("country", "")
        seniority = feed.get("seniority", ["Entry-level", "Mid-level"])
        for page in range(1, 6):
            data = fetch(f"https://himalayas.app/jobs/api/search?q={quote(slug)}"
                         + (f"&country={quote(country)}" if country else "") + f"&page={page}&limit=20")
            for j in data.get("jobs", []):
                levels = j.get("seniority") or []
                if seniority and levels and not any(l in seniority for l in levels):
                    continue
                pay = (f"{j.get('currency')} {j.get('minSalary') or 0:,} - {j['maxSalary']:,} / "
                       f"{j.get('salaryPeriod') or 'year'}" if j.get("maxSalary") else "")
                add(j["guid"], j["title"], [f"Remote - open to {country}" if country else "Remote"],
                    j["applicationLink"], strip_html(j.get("description")), j.get("companyName"),
                    id="himalayas:" + j["guid"].split("/companies/")[-1],  # same job, several searches
                    salary=pay, posted=utc_day(j["pubDate"]) if j.get("pubDate") else "")
            if len(data.get("jobs", [])) < 20:
                break
    return out


# --------------------------------------------------------------- job-alert emails

class AlertParser(HTMLParser):
    """Turn an alert email into a flat list of (text, link) pieces in reading order."""

    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.pieces, self.link, self.skip = [], None, 0

    def handle_starttag(self, tag, attrs):
        if tag in ("style", "script", "head", "title"):
            self.skip += 1
        elif tag == "a":
            self.link = dict(attrs).get("href")

    def handle_endtag(self, tag):
        if tag in ("style", "script", "head", "title"):
            self.skip = max(0, self.skip - 1)
        elif tag == "a":
            self.link = None

    def handle_data(self, data):
        text = re.sub(r"[\x00-\x20\x7f\xa0\u200b-\u200d]+", " ", data).strip()
        if text and not self.skip:
            self.pieces.append((text, self.link))


NOT_A_JOB = re.compile(r"^(?:view|see|show|browse|more|all|apply|search|explore|find|get|"
                       r"unsubscribe|manage|update|\d+\+? (?:new )?jobs?\b)", re.I)
NOT_COMPANY = re.compile(r"^(?:(?:easy(?:ly)? apply|actively|promoted|new|apply|view|save|just posted|null|"
                         r"not disclosed|\d+ (?:days?|hours?|weeks?|connections?)|be an early|early applicant|"
                         r"posted|applied|premium|verified)\b|\d(?:\.\d)?$|\W*$)", re.I)  # \d.\d is a company rating
PAY = re.compile(r"₹|\bLPA\b|\blakhs?\b|\blacs?\b|\bIDR\b|\bRp\.?\s?\d|\$\s?\d", re.I)
EXPERIENCE = re.compile(r"^\(?\d+\s*(?:-|–|to)\s*\d+\s*(?:yrs?|years?)\)?$", re.I)
APPLIED_MAIL = re.compile(r"^re:|application was sent|application successful|you applied for|job applies", re.I)
CARD_END = re.compile(r"apply\b|view all|see all|land your next|try premium|this email was|you are receiving", re.I)
SIMILAR = re.compile(r"similar job|recommended jobs", re.I)  # where a confirmation email's suggestions start


def clean_link(url):
    """Shorten a tracked job link: LinkedIn to its job page, anything very long to its path."""
    url = (url or "").strip().replace(" ", "%20")
    found = re.search(r"linkedin\.com/(?:comm/)?jobs/view/(\d+)", url)
    if found:
        return f"https://www.linkedin.com/jobs/view/{found.group(1)}/"
    return url.split("?")[0] if len(url) > 1800 else url


def jobs_in_alert(body, cfg, default_location="", subject=""):
    """Find jobs in an alert's HTML: a link whose text is a job title, then the company and
    location that follow it. Works from the text rather than one board's exact layout.
    Jobs a confirmation email says were applied to come back with applied=True."""
    parser = AlertParser()
    parser.feed(body)
    pieces, jobs = parser.pieces, {}
    subject = " ".join(subject.split())
    applied_until = 0
    if APPLIED_MAIL.search(subject):
        applied_until = next((i for i, (text, _) in enumerate(pieces) if SIMILAR.search(text)), len(pieces))
    named = re.match(r"application successful:\s*(.+?)\s+-\s+(.+?)(?:\s*\(.*\))?$", subject, re.I)
    if named:  # Hirist names the job only in its subject line
        jobs[title_key(named[2], named[1])] = {"title": named[2], "company": named[1], "applied": True}

    def is_title(text, link):
        lowered = text.lower()
        return (link and link.lower().startswith("http") and len(text) <= 100
                and not NOT_A_JOB.match(text) and "unsubscribe" not in link.lower()
                and any(has_word(k, lowered) for k in cfg["title_include"]))

    for i, (text, link) in enumerate(pieces):
        if not is_title(text, link):
            continue
        details, raw = [], [text]
        for following, follow_link in pieces[i + 1:i + 14]:
            if following == text:
                continue
            if is_title(following, follow_link) or CARD_END.match(following):
                break
            raw.append(following)
            details += [x.strip(" -–—") for x in re.split(r"\s*[·•|]\s*", following) if x.strip()]
        salary = next((d for d in details if PAY.search(d) and len(d) <= 40), "")
        details = [d for d in details if not NOT_COMPANY.match(d) and not EXPERIENCE.match(d) and d != salary]
        if not details:  # a bare list of titles says nothing about who is hiring or where
            continue
        company = details[0]
        location = details[details.index("Location") + 1] if "Location" in details[:-1] else ""
        location = location or (details[1] if len(details) > 1 else "")
        if not location or "remote" in location.lower():  # a named city stays, so Pune doesn't pass as India
            location = " / ".join(x for x in (location, default_location) if x)
        key = title_key(text, company)
        if key not in jobs:
            jobs[key] = {"title": text, "company": company, "url": clean_link(link), "salary": salary,
                         "salary_max_lpa": max_lpa(salary), "location": location, "text": " ".join(raw),
                         "applied": i < applied_until}
    return list(jobs.values())


def title_key(title, company):
    """The same job from two boards should give the same key."""
    return re.sub(r"[^a-z0-9]+", " ", f"{title}|{company}".lower()).strip()


def email_address():
    try:
        return open(EMAIL_FILE).read().strip()
    except OSError:
        return ""


def email_password(address):
    """The Gmail app password: from the JOBBOT_EMAIL_PASSWORD environment variable if set,
    otherwise from macOS Keychain."""
    password = os.environ.get(PASSWORD_ENV, "")
    if not password and sys.platform == "darwin":
        password = subprocess.run(["security", "find-generic-password", "-a", address, "-s", KEYCHAIN_SERVICE,
                                   "-w"], capture_output=True, text=True, timeout=20).stdout
    return password.strip().replace(" ", "")


def header_text(value):
    return "".join(part.decode(enc or "utf-8", "replace") if isinstance(part, bytes) else part
                   for part, enc in decode_header(value or ""))


def html_of(message):
    for part in message.walk():
        if part.get_content_type() == "text/html":
            return part.get_payload(decode=True).decode(part.get_content_charset() or "utf-8", "replace")
    return ""


def read_alerts(cfg, after_uid=0, days=None, applied=None):
    """Read job-alert emails newer than after_uid. Read-only. Returns (jobs, newest uid, stats).
    Jobs the emails confirm were applied to are added to the applied set and left out of jobs."""
    applied = set() if applied is None else applied
    address = email_address()
    password = email_password(address)
    if not password:
        raise RuntimeError(f"no app password found; run: python3 jobbot.py setup-email (macOS) "
                           f"or set {PASSWORD_ENV}")
    boards = cfg["alert_senders"]
    jobs, stats, newest = [], {}, after_uid
    box = imaplib.IMAP4_SSL("imap.gmail.com", timeout=60)
    try:
        box.login(address, password)
        listing = box.list()[1]
        all_mail = next((re.search(r'"([^"]+)"$', line.decode("utf-8", "replace")).group(1) for line in listing
                         if line and b"\\All" in line), "INBOX")
        box.select(f'"{all_mail}"', readonly=True)
        query = f'from:({" OR ".join(b[1] for b in boards)}) newer_than:{days or cfg.get("alert_days", 14)}d'
        found = box.uid("SEARCH", "X-GM-RAW", f'"{query}"')[1]
        uids = [int(u) for u in (found[0] or b"").split() if int(u) > after_uid]
        for uid in uids[-300:]:
            raw = box.uid("FETCH", str(uid), "(BODY.PEEK[])")[1]
            message = email.message_from_bytes(raw[0][1])
            sender = header_text(message.get("From")).lower()
            name, default = next(((b[0], b[2]) for b in boards if b[1] in sender), ("Email", ""))
            sent = email.utils.parsedate_to_datetime(message.get("Date"))
            found_jobs = jobs_in_alert(html_of(message), cfg, default, header_text(message.get("Subject")))
            applied.update(title_key(j["title"], j["company"]) for j in found_jobs if j["applied"])
            found_jobs = [j for j in found_jobs if not j["applied"]]
            row = stats.setdefault(name, [0, 0])
            row[0], row[1] = row[0] + 1, row[1] + len(found_jobs)
            for j in found_jobs:
                jobs.append(dict(j, id="alert:" + hashlib.sha1(title_key(j["title"], j["company"]).encode()).hexdigest()[:16],
                                 board=name, posted=sent.date().isoformat()))
            newest = max(newest, uid)
    finally:
        try:
            box.logout()
        except Exception:
            pass
    return [j for j in jobs if title_key(j["title"], j["company"]) not in applied], newest, stats


# -------------------------------------------------------------- scoring

YEARS = re.compile(  # "tahun"/"thn" (Indonesian) and "sampai"/"s/d" (to) catch Indonesian postings
    r"(\d{1,2})\s*\+?\s*(?:-|–|—|to|sampai|s/d)\s*\d{1,2}\s*\+?\s*(?:years|yrs|tahun|thn)"
    r"|(\d{1,2})\s*\+\s*(?:years|yrs|tahun|thn)"
    r"|(?:minimum|minimal|at least|min\.?)\s*(?:of\s*)?(\d{1,2})\s*(?:years|yrs|tahun|thn)", re.I)
# "Go" the language, without matching the English word: "Go, Python", "(Go)", "Go and Rust", golang
GO = re.compile(r"(?i:\bgolang\b)|[,/(]\s*Go\b|\bGo\s*(?=[,/)])|\bGo\s+(?:or|and|lang|programming|services)\b")
GO_SKILLS = ("go", "golang")

# What a posting says that matters to an applicant who needs a work permit, for regions
# that set "work_permit" in config.json.
VISA_TERMS = ["visa sponsor", "visa support", "work visa", "working visa", "work permit", "relocat",
              "expat", "foreigner"]
PERMIT_RULES = {}  # compiled rules, per work_permit setting


def any_of(phrases):
    """One case-insensitive pattern matching any of the phrases at the start of a word
    ("relocat" finds "relocation"), or None if there are none."""
    phrases = [re.escape(p) for p in phrases if p]
    if not phrases:
        return None
    return re.compile(r"\b(?:" + "|".join(phrases) + ")", re.I)


def permit_rules(region):
    """Compiled work-permit rules for a region, or None if the region doesn't ask for them.
    region["work_permit"] keys (all optional): country, nationality, language,
    citizens_only_terms, permit_terms, language_terms, local_words."""
    wp = region.get("work_permit")
    if not wp:
        return None
    wp = wp if isinstance(wp, dict) else {}
    key = json.dumps(wp, sort_keys=True)
    if key in PERMIT_RULES:
        return PERMIT_RULES[key]
    country, nationality, language = (re.escape(wp.get(k) or "") for k in ("country", "nationality", "language"))
    citizens = [re.escape(t) for t in wp.get("citizens_only_terms", [])]
    if nationality:
        citizens += [rf"{nationality} (?:citizen|national)s? only",
                     rf"must be (?:an? )?{nationality} (?:citizen|national)",
                     rf"only (?:for|open to) {nationality}"]
    signals = []
    if country:
        signals.append((re.compile(rf"(?:right|eligib\w+|authori[sz]\w+) to work in {country}", re.I),
                        f"Asks for an existing right to work in {wp['country']}"))
    signals.append((any_of(VISA_TERMS + wp.get("permit_terms", [])),
                    "Mentions visa, work permit or relocation - read the posting"))
    speaks = [re.escape(t) for t in wp.get("language_terms", [])]
    if language:
        speaks += [rf"\b{language} (?:language|speaker)", rf"(?:fluent|proficient|native)[^.]{{0,40}}\b{language}\b"]
    if speaks:
        signals.append((re.compile("|".join(speaks), re.I), f"Asks for {wp.get('language') or 'the local language'}"))
    words = any_of(wp.get("local_words", []))
    rules = {"citizens_only": re.compile("|".join(rf"\b{c}" for c in citizens), re.I) if citizens else None,
             "signals": signals,
             "local_words": re.compile(words.pattern + r"\b", re.I) if words else None,
             "language": wp.get("language") or "the local language"}
    PERMIT_RULES[key] = rules
    return rules


def required_years(text):
    found = [int(next(g for g in m.groups() if g)) for m in YEARS.finditer(text)]
    return min(found) if found else None


def age_days(day):
    """Days since a YYYY-MM-DD date; 0 when unknown so undated jobs are never treated as old."""
    try:
        return (date.today() - date.fromisoformat(day)).days
    except (TypeError, ValueError):
        return 0


def max_lpa(salary):
    """Top of an INR salary range such as "₹8L - ₹12L / yr" or "10-15 Lacs PA", in lakhs; None if not stated."""
    found = re.findall(r"₹\s*([\d.]+)\s*L|([\d.]+)\s*(?:lacs?|lakhs?|lpa)\b", salary or "", re.I)
    found = [x for pair in found for x in pair if x]
    return max(float(x) for x in found) if found else None


def has_word(word, text):
    """True if word appears in text with no letter directly before or after it ("sde" is in
    "sde2" but "go developer" is not in "django developer")."""
    return re.search(r"(?<![a-z])" + re.escape(word) + r"(?![a-z])", text) is not None


def permit_notes(text, rules):
    notes = [label for pattern, label in rules["signals"] if pattern.search(text)]
    if rules["local_words"] and len(rules["local_words"].findall(text)) >= 8:
        notes.append(f"Posting is written in {rules['language']}")
    return "; ".join(notes) or "No mention of visa sponsorship - ask before applying"


def score(job, cfg, region):
    """Return (score, years, reason) for a job in a region, or None if it is filtered out."""
    title = job["title"].lower()
    if not any(k in job["location"].lower() for k in region["locations"]):
        return None
    if not any(has_word(k, title) for k in cfg["title_include"]):
        return None
    if any(has_word(k, title) for k in cfg["title_exclude"]):
        return None
    text = job["text"]
    rules = permit_rules(region)
    if rules and rules["citizens_only"] and rules["citizens_only"].search(text):
        return None
    years = job.get("years")  # boards that state experience; otherwise read it from the text
    if years is None:
        years = required_years(text)
    if years is not None and years > cfg.get("stretch_years", cfg["max_required_years"]):
        return None
    lpa = job.get("salary_max_lpa")  # only known when the listing states a salary in INR
    if lpa is not None and lpa < cfg.get("min_salary_lpa", 0):
        return None
    if age_days(job.get("posted")) > cfg.get("max_posting_age_days", 10 ** 6):
        return None
    blob = (job["title"] + " " + text).lower()
    total, hits = (45 if job.get("board") else 20), []  # an alert already matched a saved search
    for skill, pts in cfg["skills"].items():
        if skill.lower() in GO_SKILLS:
            matched = bool(GO.search(job["title"] + " " + text))
        else:
            matched = re.search(r"\b" + re.escape(skill.lower()), blob) is not None
        if matched:
            total += pts
            hits.append("Go" if skill.lower() in GO_SKILLS else skill)
    if years is not None:
        total += 15 if years <= 1 else 8
    return min(total, 100), years, ", ".join(hits)


# ----------------------------------------------------------------- xlsx

NS = {"m": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
      "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
      "p": "http://schemas.openxmlformats.org/package/2006/relationships"}


def col(n):
    return chr(ord("A") + n)


def read_user_edits(path):
    """Read Status and Notes the user typed, as {apply link: [(status, notes), ...]} -
    one entry per sheet the job appears on."""
    edits = {}
    if not os.path.exists(path):
        return edits
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            for si in ET.fromstring(z.read("xl/sharedStrings.xml")).findall("m:si", NS):
                shared.append("".join(t.text or "" for t in si.iter("{%s}t" % NS["m"])))
        for name in z.namelist():
            if not re.match(r"xl/worksheets/sheet\d+\.xml$", name):
                continue
            rows = []
            for row in ET.fromstring(z.read(name)).iter("{%s}row" % NS["m"]):
                cells = {}
                for c in row.findall("m:c", NS):
                    letters = re.match(r"[A-Z]+", c.get("r")).group()
                    v = c.find("m:v", NS)
                    if c.get("t") == "s" and v is not None:
                        val = shared[int(v.text)]
                    elif c.get("t") == "inlineStr":
                        val = "".join(t.text or "" for t in c.iter("{%s}t" % NS["m"]))
                    else:
                        val = v.text if v is not None else ""
                    cells[letters] = val or ""
                rows.append(cells)
            if not rows:
                continue
            head = {v: k for k, v in rows[0].items()}
            if not all(h in head for h in ("Apply link", "Status", "Notes")):
                continue
            for r in rows[1:]:
                link = r.get(head["Apply link"], "")
                if link:
                    edits.setdefault(link, []).append((r.get(head["Status"], ""), r.get(head["Notes"], "")))
    return edits


def sheet_xml(rows, widths, link_col, status_col=None):
    """Build one worksheet and its hyperlink relationships from a list of rows."""
    cols = "".join(f'<col min="{i + 1}" max="{i + 1}" width="{w}" customWidth="1"/>'
                   for i, w in enumerate(widths))
    body, links, rels = [], [], []
    for ri, row in enumerate(rows, 1):
        cells = []
        for ci, val in enumerate(row):
            ref = f"{col(ci)}{ri}"
            style = 1 if ri == 1 else (2 if ci == link_col else 0)
            if isinstance(val, (int, float)):
                cells.append(f'<c r="{ref}" s="{style}"><v>{val}</v></c>')
            else:
                cells.append(f'<c r="{ref}" t="inlineStr" s="{style}"><is><t xml:space="preserve">'
                             f'{escape(str(val))}</t></is></c>')
            if ri > 1 and ci == link_col and val:
                rid = f"rId{len(rels) + 1}"
                links.append(f'<hyperlink ref="{ref}" r:id="{rid}"/>')
                rels.append(f'<Relationship Id="{rid}" Type="{NS["r"]}/hyperlink" '
                            f'Target="{escape(str(val), {chr(34): "&quot;"})}" TargetMode="External"/>')
        body.append(f'<row r="{ri}">{"".join(cells)}</row>')
    xml = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<worksheet xmlns="{NS["m"]}" xmlns:r="{NS["r"]}">'
           f'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" '
           f'activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
           f'<cols>{cols}</cols><sheetData>{"".join(body)}</sheetData>'
           + (f'<dataValidations count="1"><dataValidation type="list" allowBlank="1" '
              f'sqref="{col(status_col)}2:{col(status_col)}5000"><formula1>"{",".join(STATUSES)}"'
              f'</formula1></dataValidation></dataValidations>' if status_col is not None else "")
           + (f'<hyperlinks>{"".join(links)}</hyperlinks>' if links else "") + '</worksheet>')
    rel = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           f'<Relationships xmlns="{NS["p"]}">{"".join(rels)}</Relationships>')
    return xml, rel


def write_workbook(path, sheets):
    """sheets: list of (name, rows, widths, link_col, status_col)."""
    ct = "".join(f'<Override PartName="/xl/worksheets/sheet{i}.xml" ContentType="application/'
                 f'vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
                 for i in range(1, len(sheets) + 1))
    wb_sheets = "".join(f'<sheet name="{escape(n[:31])}" sheetId="{i}" r:id="rId{i}"/>'
                        for i, (n, *_) in enumerate(sheets, 1))
    wb_rels = "".join(f'<Relationship Id="rId{i}" Type="{NS["r"]}/worksheet" '
                      f'Target="worksheets/sheet{i}.xml"/>' for i in range(1, len(sheets) + 1))
    styles_id = len(sheets) + 1
    tmp = path + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml",
                   '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
                   '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
                   '<Default Extension="xml" ContentType="application/xml"/>'
                   '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
                   '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
                   + ct + '</Types>')
        z.writestr("_rels/.rels",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<Relationships xmlns="{NS["p"]}"><Relationship Id="rId1" '
                   f'Type="{NS["r"]}/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr("xl/workbook.xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<workbook xmlns="{NS["m"]}" xmlns:r="{NS["r"]}"><sheets>{wb_sheets}</sheets></workbook>')
        z.writestr("xl/_rels/workbook.xml.rels",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<Relationships xmlns="{NS["p"]}">{wb_rels}<Relationship Id="rId{styles_id}" '
                   f'Type="{NS["r"]}/styles" Target="styles.xml"/></Relationships>')
        z.writestr("xl/styles.xml",
                   f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                   f'<styleSheet xmlns="{NS["m"]}">'
                   '<fonts count="3"><font><sz val="11"/><name val="Calibri"/></font>'
                   '<font><b/><sz val="11"/><name val="Calibri"/></font>'
                   '<font><u/><sz val="11"/><color rgb="FF0563C1"/><name val="Calibri"/></font></fonts>'
                   '<fills count="2"><fill><patternFill patternType="none"/></fill>'
                   '<fill><patternFill patternType="gray125"/></fill></fills>'
                   '<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>'
                   '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
                   '<cellXfs count="3"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
                   '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
                   '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/></cellXfs>'
                   '</styleSheet>')
        for i, (_, rows, widths, link_col, status_col) in enumerate(sheets, 1):
            xml, rel = sheet_xml(rows, widths, link_col, status_col)
            z.writestr(f"xl/worksheets/sheet{i}.xml", xml)
            z.writestr(f"xl/worksheets/_rels/sheet{i}.xml.rels", rel)
    os.replace(tmp, path)


def is_target(job, cfg):
    """Company is on the target_employers list in config.json (the employers you most want)."""
    return any(has_word(t.lower(), job["company"].lower()) for t in cfg.get("target_employers", []))


def priority(job, cfg):
    """Order for the Apply today sheet: fit, plus bonuses for a target employer, a stated INR
    salary of at least priority_salary_lpa, and a posting from the last week."""
    lpa = max_lpa(job.get("salary"))
    good_pay = lpa is not None and lpa >= cfg.get("priority_salary_lpa", float("inf"))
    return (job["score"] + (15 if is_target(job, cfg) else 0) + (10 if good_pay else 0)
            + (10 if job.get("posted") and age_days(job["posted"]) <= 7 else 0))


def export(jobs, region, cfg):
    """Rebuild a region's workbook from its jobs, keeping Status/Notes typed by the user."""
    path = os.path.join(HERE, region["excel_file"])
    headers, widths = list(HEADERS), list(WIDTHS)
    if region.get("work_permit"):
        at = headers.index("Apply link")
        headers.insert(at, PERMIT_HEADER)
        widths.insert(at, PERMIT_WIDTH)
    try:
        edits = read_user_edits(path)
    except Exception as e:  # unreadable workbook: keep a copy rather than lose edits
        backup = path + f".unreadable-{datetime.now():%Y%m%d-%H%M%S}"
        os.replace(path, backup)
        log(f"could not read {region['excel_file']} ({e}); kept a copy at {os.path.basename(backup)}")
        edits = {}
    today = date.today().isoformat()
    for job in jobs.values():
        for i, field in enumerate(("status", "notes")):
            # a job can be on two sheets; whichever copy differs from last time is the new edit
            changed = [seen[i] for seen in edits.get(job["url"], []) if seen[i] != job.get(field, "")]
            if changed:
                job[field] = changed[0]
        if job.get("status") == "Applied" and not job.get("applied_on"):
            job["applied_on"] = today

    def row(j, extra=()):
        cells = {"Found at": j["found"],
                 "Posted": (j.get("posted") or "") + (" (closed)" if j.get("closed") else ""),
                 "Company": ("★ " if is_target(j, cfg) else "") + j["company"],
                 "Role": j["title"], "Location": j["location"],
                 "Fit": j["score"],
                 "Experience asked": "" if j["years"] is None else f"{j['years']}+ years",
                 "Salary": j.get("salary", ""), "Why it matched": j["why"],
                 PERMIT_HEADER: j.get("permit", ""), "Apply link": j["url"],
                 "Status": j.get("status", ""), "Notes": j.get("notes", "")}
        return [cells[h] for h in headers] + list(extra)

    link_col, status_col = headers.index("Apply link"), headers.index("Status")
    sheets = []
    max_age = cfg.get("max_posting_age_days", 10 ** 6)
    main = {k: j for k, j in jobs.items() if not j.get("stretch")}
    todo = [j for j in main.values() if j.get("status", "") in ("", "To apply")
            and not j.get("closed") and age_days(j.get("posted")) <= max_age]
    todo = sorted(todo, key=lambda j: -priority(j, cfg))[:TODAY_ROWS]
    sheets.append((TODAY_SHEET, [headers] + [row(j) for j in todo], widths, link_col, status_col))
    due = [j for j in jobs.values() if j.get("status") == "Applied" and not j.get("closed")
           and age_days(j.get("applied_on")) >= FOLLOW_UP_DAYS]
    if due:
        sheets.append((FOLLOW_SHEET, [headers + ["Applied on"]] +
                       [row(j, [j["applied_on"]]) for j in sorted(due, key=lambda j: j["applied_on"], reverse=True)],
                       widths + [12], link_col, status_col))
    stretch = [j for j in jobs.values() if j.get("stretch") and not j.get("closed")]
    if stretch:
        sheets.append((STRETCH_SHEET, [headers] + [row(j) for j in sorted(stretch, key=lambda j: -j["score"])],
                       widths, link_col, status_col))
    days = {}
    for job in main.values():
        days.setdefault(job["found"][:10], []).append(job)
    for day in sorted(days, reverse=True):
        rows = [headers] + [row(j) for j in sorted(days[day], key=lambda j: -j["score"])]
        sheets.append((day, rows, widths, link_col, status_col))
    boards = [["Where to check daily", "Link"]] + [list(b) for b in region["boards"]]
    sheets.append((BOARDS_SHEET, boards, [42, 80], 1, None))
    write_workbook(path, sheets)


# ------------------------------------------------------------- run loop

def notify(message):
    """Show a macOS notification; does nothing elsewhere or if it fails."""
    if sys.platform != "darwin":
        return
    text = message.replace("\\", "").replace('"', "'")
    try:
        subprocess.run(["osascript", "-e", f'display notification "{text}" with title "jobbot"'],
                       timeout=10, capture_output=True)
    except Exception:
        pass


def run_once(cfg, state):
    try:
        socket.getaddrinfo("boards-api.greenhouse.io", 443)
    except OSError:
        raise RuntimeError("no internet connection (Mac asleep or offline?); skipped this pass")
    wanted = sorted({k for region in cfg["regions"] for k in region["locations"]})
    failed = set()

    def safe(feed):
        try:
            return jobs_from(feed, wanted)
        except Exception as e:
            log(f"feed failed: {feed['name']} ({e})")
            failed.add(f"{feed['ats']}:{feed['slug']}")
            return []
    with ThreadPoolExecutor(8) as pool:
        openings = [j for batch in pool.map(safe, cfg["feeds"]) for j in batch]
    listed = {j["id"]: j for j in openings}
    if email_address():
        try:
            applied = set(state.get("_applied_keys", []))
            alerts, state["_email_last_uid"], stats = read_alerts(cfg, state.get("_email_last_uid", 0), applied=applied)
            state["_applied_keys"] = sorted(applied)
            openings += [j for j in alerts if title_key(j["title"], j["company"]) not in applied]
            if stats:
                log("email alerts: " + ", ".join(f"{k} {n} emails, {c} jobs" for k, (n, c) in stats.items()))
        except Exception as e:
            log(f"email alerts skipped: {e}")
    checked = {f"{f['ats']}:{f['slug']}" for f in cfg["feeds"]} - failed
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    totals, strong = [], []
    for region in cfg["regions"]:
        jobs, new = state[region["key"]], 0
        have = {title_key(j["title"], j["company"]) for j in jobs.values()}
        for job_id, job in jobs.items():
            # Cutshort pages show only the newest 50 jobs, so leaving one doesn't mean closed
            source = ":".join(job_id.split(":")[:2])
            if not job_id.startswith("cutshort:") and source in checked:
                job["closed"] = job_id not in listed
            if not job.get("posted") and job_id in listed:  # matches found before dates were kept
                job["posted"] = listed[job_id].get("posted", "")
        for job in openings:
            if job["id"] in jobs or (job.get("board") and title_key(job["title"], job["company"]) in have):
                continue
            result = score(job, cfg, region)
            if result is None or result[0] < cfg["min_score"]:
                continue
            stretch = result[1] is not None and result[1] > cfg["max_required_years"]
            have.add(title_key(job["title"], job["company"]))
            board = f"{job['board']} alert" if job.get("board") else ""
            rules = permit_rules(region)
            permit = permit_notes(job["text"], rules) if rules else ""
            jobs[job["id"]] = {
                "found": now, "posted": job.get("posted", ""), "company": job["company"],
                "title": job["title"], "location": job["location"], "url": job["url"],
                "score": result[0], "years": result[1],
                "why": ": ".join(x for x in (board, result[2]) if x), "stretch": stretch,
                "salary": job.get("salary", ""), "status": "", "notes": "",
                "permit": "From an email alert - read the posting for visa terms"
                if board and permit else permit}
            new += 1
            if result[0] >= cfg.get("notify_min_fit", 101) and not stretch:
                strong.append(f"{job['company']} - {job['title']}")
            log(f"match {result[0]:>3}  {job['company']} - {job['title']} ({job['location']})")
        export(jobs, region, cfg)
        totals.append(f"{new} new, {len(jobs)} total in {region['excel_file']}")
    save_json(STATE, state)
    log(f"pass done: scanned {len(openings)} openings; " + "; ".join(totals))
    if strong:
        notify(f"{len(strong)} strong new match{'es' if len(strong) > 1 else ''}: {strong[0]}"
               + (" and more" if len(strong) > 1 else ""))


def run_forever():
    stop = {"now": False}
    signal.signal(signal.SIGTERM, lambda *_: stop.update(now=True))
    signal.signal(signal.SIGINT, lambda *_: stop.update(now=True))
    while not stop["now"]:
        wait = 600  # retry sooner if this pass fails, e.g. on a typo in config.json
        try:
            cfg = load_config()
            run_once(cfg, load_state(cfg))
            wait = cfg["poll_minutes"] * 60
        except Exception as e:
            log(f"pass failed: {e}")
        wake = time.time() + wait
        while time.time() < wake and not stop["now"]:
            time.sleep(1)
    log("stopped")


def running_pid():
    try:
        pid = int(open(PIDFILE).read())
        os.kill(pid, 0)
        return pid
    except (OSError, ValueError):
        return None


def setup_email():
    """Connect Gmail. On macOS the app password goes straight into Keychain via its own hidden
    prompt; elsewhere it is read from the JOBBOT_EMAIL_PASSWORD environment variable."""
    address = input("Gmail address that receives your job alerts: ").strip()
    if sys.platform == "darwin":
        print("\nPaste your 16-character Google app password when asked, then paste it again to confirm.\n"
              "Nothing shows while you paste. It is stored in macOS Keychain, not in this folder.\n")
        subprocess.run(["security", "add-generic-password", "-U", "-a", address, "-s", KEYCHAIN_SERVICE,
                        "-T", "/usr/bin/security", "-w"], check=True)
    with open(EMAIL_FILE, "w") as f:
        f.write(address)
    os.chmod(EMAIL_FILE, 0o600)
    if sys.platform != "darwin" and not os.environ.get(PASSWORD_ENV):
        print(f"\nSaved the address. Now put your 16-character Google app password in the {PASSWORD_ENV}\n"
              f"environment variable (for example in your shell profile) and run: python3 jobbot.py email-check")
        return
    email_check()


def email_check():
    """Show what the alert reader finds, without saving anything."""
    if not email_address():
        print("Gmail is not connected yet. Run: python3 jobbot.py setup-email")
        return
    cfg = load_config()
    try:
        jobs, _, stats = read_alerts(cfg)
    except Exception as e:
        print(f"Could not read Gmail: {e}\nCheck the address, the app password, and that the "
              "account has 2-Step Verification on. To redo it: python3 jobbot.py setup-email")
        return
    print(f"Connected. Alert emails in the last {cfg.get('alert_days', 14)} days:")
    for name, (emails, count) in sorted(stats.items()):
        print(f"  {name}: {emails} emails, {count} jobs read" + ("   <- no job listings in these" if not count else ""))
    if not stats:
        print("  none yet - create the job alerts first; they arrive over the next day or two")
    for region in cfg["regions"]:
        matched = [j for j in jobs if (score(j, cfg, region) or (0,))[0] >= cfg["min_score"]]
        print(f"\n{region['name']}: {len(matched)} would match. Examples:")
        for j in matched[:8]:
            print(f"  {j['title']} | {j['company']} | {j['location']} | {j['board']}")


def main():
    cmd = sys.argv[1] if len(sys.argv) > 1 else "status"
    if cmd == "start":
        if running_pid():
            print(f"already running (pid {running_pid()})")
            return
        with open(LOGFILE, "a") as out:
            p = subprocess.Popen([sys.executable, os.path.abspath(__file__), "run"],
                                 stdout=out, stderr=out, stdin=subprocess.DEVNULL,
                                 start_new_session=True, cwd=HERE)
        with open(PIDFILE, "w") as f:
            f.write(str(p.pid))
        print(f"started (pid {p.pid}). Log: jobbot.log")
    elif cmd == "stop":
        pid = running_pid()
        if not pid:
            print("not running")
            return
        os.kill(pid, signal.SIGTERM)
        for _ in range(100):
            if not running_pid():
                break
            time.sleep(0.1)
        os.path.exists(PIDFILE) and os.remove(PIDFILE)
        print("stopped")
    elif cmd == "status":
        cfg = load_config()
        state = load_state(cfg)
        pid = running_pid()
        print(f"running (pid {pid})" if pid else "not running")
        for region in cfg["regions"]:
            n = len(state[region["key"]])
            print(f"  {region['name']}: {n} match{'' if n == 1 else 'es'} logged in "
                  f"{region['excel_file']}")
    elif cmd == "once":
        cfg = load_config()
        run_once(cfg, load_state(cfg))
    elif cmd == "setup-email":
        setup_email()
    elif cmd == "email-check":
        email_check()
    elif cmd == "run":
        run_forever()
    else:
        print(__doc__)


if __name__ == "__main__":
    main()
