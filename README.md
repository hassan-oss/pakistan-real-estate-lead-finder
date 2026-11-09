# Pakistan Real Estate Lead Finder

A Python tool that finds real estate agencies in Pakistan and builds a verified lead sheet:
company, website, decision-maker, email, phone, and a source link for every field.

## Features
- Automatic discovery across 14 Pakistani cities (no seed list needed); optional Google Places
- Parallel crawling of Home / About / Team / Contact pages (robots.txt respected, rate-limited)
- Filters out non-real-estate, government and education sites
- Decision-maker extraction with junk-name rejection (chatbots, forms, testimonials)
- Emails are only taken from published pages, **never guessed**; generic emails are flagged
- Quality gate: at most one of (decision-maker, email, phone) may be missing; each field has its own source link
- Confidence score: High = all three fields present and the email matches the decision-maker's name
- Deduplication by domain, phone, email and company name
- Excel output (Leads + Summary sheets) that imports into Google Sheets

## Install
```bash
python3 -m venv venv && source venv/bin/activate
pip install -r requirements.txt
```

## Usage
```bash
python lead_finder_v2.py --fast --cities Lahore Islamabad Karachi --target 100 --out leads.xlsx
```

| Flag | Meaning |
|---|---|
| `--cities` | Cities to search (default: 14 major cities) |
| `--target` | Number of qualified companies to collect |
| `--fast` | Fewer pages per site, more workers |
| `--max-missing 0` | Require decision-maker, email and phone |
| `--min-confidence high` | Keep only High-confidence rows |
| `--google` | Add Google Places (needs `GOOGLE_API_KEY`) |
| `--serpapi-key` | Fill missing decision-makers from public LinkedIn search snippets |
| `--js` | Render JavaScript-only sites (needs Playwright) |

## Limitations
- Many agencies publish no owner name or personal email, so results can fall short of the target.
- LinkedIn is not scraped. The optional lookup uses public search-result snippets only.
- Spot-check Medium-confidence rows before delivering them to a client.

## Responsible use
Collect only publicly listed business contact details, follow each site's terms and robots.txt,
and comply with applicable data-protection and anti-spam rules when contacting leads.
