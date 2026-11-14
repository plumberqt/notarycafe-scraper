# NotaryCafe Scraper - Complete Technical Specification
## For AntiGravity / External System Integration

---

## 🎯 Overview

This document describes the **exact scraping workflow**, **every click/action**, **data schema**, and **Excel output structure** for the NotaryCafe adaptive scraper. Use this to replicate, audit, or feed into AntiGravity or any downstream system.

---

## 🔧 System Architecture

```
┌─────────────────────────────────────────────────────────────────┐
│                    NOTARYCAFE SCRAPER PIPELINE                  │
├─────────────────────────────────────────────────────────────────┤
│  1. LOAD CITIES          →  Excel (All_Cities_Raw sheet)       │
│  2. RESUME PROGRESS      →  JSON (scraping_progress_*.json)    │
│  3. LAUNCH CHROME        →  Persistent Context (Profile 3)     │
│  4. FOR EACH CITY:                                            │
│     ├─ SEARCH            →  Click search box → Type "City, ST" → Enter
│     ├─ WAIT LISTINGS     →  Detect table/div/card elements
│     ├─ PAGINATE          →  Click "Next" or page numbers
│     ├─ FOR EACH LISTING:                                      │
│     │   ├─ CLICK         →  Navigate to detail page
│     │   ├─ SCROLL        →  Human-like scroll (5 steps)
│     │   ├─ EXTRACT       →  Regex + JSON-LD + HTML attributes
│     │   └─ STORE         →  Append to all_listings[]
│     └─ SAVE BATCH        →  Every 10 listings → Excel + CSV
│  5. FINAL SAVE           →  Complete Excel + Progress JSON     │
└─────────────────────────────────────────────────────────────────┘
```

---

## 📋 Complete Click-by-Click Workflow

### Phase 1: Browser Initialization

```python
# EXACT SEQUENCE - Do not reorder
1. playwright.chromium.launch_persistent_context(
     user_data_dir=r"C:\Users\nishu\AppData\Local\Google\Chrome\User Data",
     channel="chrome",           # Use installed Chrome, not bundled
     headless=False,             # Visible browser (required for login state)
     args=[
         "--disable-blink-features=AutomationControlled",
         "--no-first-run",
         "--no-default-browser-check"
     ],
     viewport={"width": 1920, "height": 1080},
     locale="en-US",
     timezone_id="America/New_York"
   )

2. context.add_init_script("""
     Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
     window.chrome = {runtime: {}};
   """)

3. page = context.pages[0] or await context.new_page()
```

### Phase 2: City Search (Per City)

```python
# URL: https://www.notarycafe.com/search
await page.goto("https://www.notarycafe.com/search", wait_until='domcontentloaded', timeout=60000)
await asyncio.sleep(random.uniform(2, 4))

# FIND SEARCH INPUT - Try selectors in ORDER until one works:
SEARCH_SELECTORS = [
    'input[name="search"]',
    'input[name="q"]',
    'input[type="search"]',
    'input[placeholder*="city" i]',
    'input[placeholder*="location" i]',
    'input[name*="city" i]',
    'input[id*="search" i]',
    'input[class*="search" i]',
    'form input[type="text"]'
]
# Fallback: any visible input[type="text"] or input[type="search"]

# CLICK SEQUENCE:
await search_input.click()                    # Focus
await search_input.fill("")                   # Clear
await asyncio.sleep(0.3)
await search_input.type(f"{city}, {state}", delay=random.randint(50, 150))  # Human typing
await asyncio.sleep(0.5)
await search_input.press('Enter')             # Submit
await page.wait_for_load_state('domcontentloaded', timeout=60000)
await asyncio.sleep(random.uniform(2, 4))
```

### Phase 3: Wait for Listings & Detect Structure

```python
# WAIT FOR LISTINGS - Try each indicator until found:
LISTING_INDICATORS = [
    'table tbody tr',
    'div[class*="listing"]',
    'div[class*="result"]',
    'a[href*="/listing/"]',
    'a[href*="/profile/"]',
    'div[class*="card"]'
]

# For each selector:
await page.wait_for_selector(selector, timeout=5000)
elements = await page.query_selector_all(selector)
if elements and len(elements) > 0: return True

# CHECK FOR NO RESULTS:
body_text = await page.inner_text('body').lower()
NO_RESULT_INDICATORS = [
    'no results', 'no listings', 'no matches',
    '0 results', '0 listings', 'nothing found',
    'no notaries found', 'no notaries in'
]
if any(indicator in body_text for indicator in NO_RESULT_INDICATORS):
    return "NO_RESULTS"
```

### Phase 4: Pagination Detection

```python
# FIND PAGINATION CONTAINER:
PAGINATION_SELECTORS = [
    'nav.pagination', '.pagination', '.pager',
    '[class*="pagination"]', '[class*="pager"]',
    'nav[aria-label="pagination"]'
]

# FIND TOTAL PAGES:
page_links = await page.query_selector_all('a[href*="page"], a[href*="p="], a[href*="pg="]')
page_numbers = [int(await link.inner_text()) for link in page_links if (await link.inner_text()).isdigit()]
total_pages = max(page_numbers) if page_numbers else 1

# FIND NEXT BUTTON:
NEXT_SELECTORS = [
    'a:has-text("Next")', 'a:has-text("next")',
    'a[rel="next"]', '.pagination .next a',
    '.pager .next a', 'a[aria-label="Next"]'
]
# First visible/enabled one wins

# CLICK NEXT PAGE:
await next_button.click()
await page.wait_for_load_state('domcontentloaded', timeout=30000)
await asyncio.sleep(random.uniform(2, 4))
```

### Phase 5: Find Listing Elements on Page

```python
# ADAPTIVE STRATEGY - Try in ORDER:
LISTING_STRATEGIES = [
    ('table tbody tr', 'table rows'),
    ('div[class*="listing"]', 'listing divs'),
    ('div[class*="result"]', 'result divs'),
    ('div[class*="card"]', 'card divs'),
    ('a[href*="/listing/"]', 'listing links'),
    ('a[href*="/profile/"]', 'profile links'),
    ('tr[class*="row"]', 'row classes'),
    ('li[class*="item"]', 'list items'),
]

# For each strategy:
elements = await page.query_selector_all(selector)
valid_elements = []
for el in elements:
    text = await el.inner_text()
    if len(text.strip()) > 10:  # Meaningful content filter
        valid_elements.append(el)
if valid_elements: return valid_elements
```

### Phase 6: Extract Summary from Listing Row (Before Click)

```python
# FROM EACH LISTING ELEMENT:
data = {
    'city_searched': city,
    'state_searched': state,
    'page_number': current_page,
    'position_on_page': idx,
    'listing_name': '',
    'listing_url': '',
    'name': '',
    'email': '',
    'phone': '',
    'city': '',
    'state': '',
    'zip': '',
    'address': '',
    'raw_text': ''
}

# GET LINK:
link = await element.query_selector('a[href]')
if link:
    href = await link.get_attribute('href')
    data['listing_url'] = href if href.startswith('http') else f"https://www.notarycafe.com{href}"
    data['listing_name'] = (await link.inner_text()).strip()

# EXTRACT FROM TEXT:
full_text = await element.inner_text()
data['raw_text'] = full_text[:500]

lines = [l.strip() for l in full_text.split('\n') if l.strip()]
if lines: data['listing_name'] = lines[0]

# REGEX EXTRACTION:
# Phone: \(\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}
# Email: [\w\.-]+@[\w\.-]+\.\w+
# Address: ([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5})
```

### Phase 7: Click Listing → Detail Page

```python
# NAVIGATE:
await page.goto(listing_url, wait_until='domcontentloaded', timeout=60000)
await asyncio.sleep(random.uniform(2, 4))

# HUMAN-LIKE SCROLL (5 steps):
await page.evaluate("""
    () => {
        const height = document.body.scrollHeight;
        const steps = 5;
        for (let i = 1; i <= steps; i++) {
            window.scrollTo(0, height * i / steps);
            await new Promise(r => setTimeout(r, Math.random() * 500 + 200));
        }
        window.scrollTo(0, 0);
    }
""")
await asyncio.sleep(random.uniform(0.5, 1.5))
```

### Phase 8: Extract ALL Fields from Detail Page

```python
# GET FULL PAGE TEXT & HTML:
full_text = await page.inner_text('body')
raw_html = await page.content()

# 1. NAME (from H1):
h1 = await page.query_selector('h1')
name = (await h1.inner_text()).strip() if h1 else ''

# 2. EMAIL - Multiple strategies:
# Strategy A: Regex on text
email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', full_text)
# Strategy B: mailto: links in HTML
mailto_match = re.search(r'mailto:([\w\.-]+@[\w\.-]+\.\w+)', raw_html)
# Strategy C: JSON-LD structured data
json_ld = re.findall(r'<script type="application/ld\+json">(.*?)</script>', raw_html, re.DOTALL)
for match in json_ld:
    data = json.loads(match.strip())
    if 'email' in data: email = data['email']

# 3. PHONE - Multiple patterns:
PHONE_PATTERNS = [
    r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',
    r'\d{3}-\d{3}-\d{4}',
    r'\(\d{3}\)\s*\d{3}-\d{4}',
]
# Also check tel: links and JSON-LD telephone

# 4. ADDRESS - Multiple patterns:
ADDR_PATTERNS = [
    r'(\d+\s+[A-Za-z0-9\s]+(?:Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Lane|Ln|Boulevard|Blvd|Court|Ct|Circle|Cir|Way|Place|Pl)\.?),\s*([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)',
    r'(\d+\s+[A-Za-z0-9\s]+),\s*([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5})',
]
# Also parse JSON-LD address object

# 5. COMMISSION:
COMMISSION_PATTERNS = [
    r'[Cc]ommission\s*(?:number|#|no\.?)\s*[:\-]?\s*([A-Z0-9\-]+)',
    r'[Cc]ommission\s*[Ee]xpir(?:ation|y)\s*[:\-]?\s*(\d{1,2}/\d{1,2}/\d{4})',
]

# 6. SEMANTIC FIELD EXTRACTION (Label: Value):
SEMANTIC_FIELDS = {
    'services_offered': r'(?:[Ss]ervices|[Oo]ffered?)\s*[:\-]\s*([^\n]{20,200})',
    'languages_spoken': r'[Ll]anguages?\s*[:\-]\s*([^\n]{5,100})',
    'business_hours': r'(?:[Hh]ours?|[Oo]pen|[Aa]vailable)\s*[:\-]\s*([^\n]{5,100})',
    'fees': r'(?:[Ff]ees?|[Cc]ost|[Pp]rice|[Rr]ate)\s*[:\-]\s*([^\n]{5,100})',
    'coverage_area': r'(?:[Cc]overage|[Aa]rea|[Ss]erves?)\s*[:\-]\s*([^\n]{10,200})',
    'years_in_business': r'(?:[Yy]ears?|[Ee]xperience)\s*[:\-]\s*([^\n]{5,50})',
    'background_check': r'[Bb]ackground\s*[Cc]heck\s*[:\-]\s*([^\n]{5,100})',
    'bonding_company': r'[Bb]onding\s*[Cc]ompany\s*[:\-]\s*([^\n]{5,100})',
    'errors_omissions_insurance': r'(?:[Ee]rrors?|[Oo]missions?|[E&O])\s*[:\-]\s*([^\n]{5,100})',
    'notary_type': r'[Nn]otary\s*[Tt]ype\s*[:\-]\s*([^\n]{5,100})',
}
```

### Phase 9: Build Complete Record

```python
detail = {
    # IDENTIFIERS
    'listing_name': name or listing_summary.get('listing_name', ''),
    'listing_url': listing_url,
    'city_searched': city,
    'state_searched': state,
    'page_number': page_num,
    'position_on_page': position,
    
    # CONTACT
    'name': name,
    'email': email or '',
    'phone': phone or '',
    'mobile': '',           # Separate if distinguishable
    'fax': '',
    'website': '',
    
    # ADDRESS
    'street_address': street_address or '',
    'city': city_addr or '',
    'state': state_addr or '',
    'zip_code': zip_code or '',
    'county': '',
    
    # CREDENTIALS
    'commission_number': commission_number or '',
    'commission_expiration': commission_expiration or '',
    'bonding_company': bonding_company or '',
    'errors_omissions_insurance': e_o_insurance or '',
    'notary_type': notary_type or '',
    
    # SERVICES & BUSINESS
    'services_offered': services_offered or '',
    'languages_spoken': languages_spoken or '',
    'business_hours': business_hours or '',
    'fees': fees or '',
    'coverage_area': coverage_area or '',
    'years_in_business': years_in_business or '',
    'background_check': background_check or '',
    
    # METADATA
    'description': '',
    'scraped_at': datetime.now().isoformat(),
    'scraped_from_url': page.url,
    'source_page_number': page_num,
    'position_on_page': position,
    'raw_html': raw_html[:50000],  # Truncated for storage
}

self.all_listings.append(detail)
self.total_listings_scraped += 1
```

### Phase 10: Incremental Save (Every 10 Listings)

```python
if self.total_listings_scraped % self.batch_size == 0:
    self.save_listings_incremental()
    self.save_progress()

# SAVE_LISTINGS_INCREMENTAL():
df = pd.DataFrame(self.all_listings)
with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
    # SHEET 1: All_Listings (raw)
    pd.DataFrame(self.all_listings).to_excel(writer, sheet_name='All_Listings', index=False)
    
    # SHEET 2: Summary (clean)
    summary_data = []
    for l in self.all_listings:
        summary_data.append({
            'Name': l.get('name', ''),
            'City_Searched': l.get('city_searched', ''),
            'State_Searched': l.get('state_searched', ''),
            'Listing_Name': l.get('listing_name', ''),
            'Email': l.get('email', ''),
            'Phone': l.get('phone', ''),
            'Mobile': l.get('mobile', ''),
            'Address': f"{l.get('street_address', '')}, {l.get('city', '')}, {l.get('state', '')} {l.get('zip_code', '')}",
            'Commission_Number': l.get('commission_number', ''),
            'Commission_Expiration': l.get('commission_expiration', ''),
            'Services': l.get('services_offered', ''),
            'Scraped_At': l.get('scraped_at', '')
        })
    pd.DataFrame(summary_data).to_excel(writer, sheet_name='Summary', index=False)

# CSV export
pd.DataFrame(self.all_listings).to_csv(csv_path, index=False, encoding='utf-8')
```

---

## 📊 Excel Output Structure

### File Naming
```
notarycafe_listings_YYYYMMDD_HHMMSS.xlsx
notarycafe_listings_YYYYMMDD_HHMMSS.csv
scraping_progress_YYYYMMDD_HHMMSS.json
error_log_YYYYMMDD_HHMMSS.jsonl
```

### Sheet 1: `All_Listings` (Complete Raw Data)
| Column | Type | Description |
|--------|------|-------------|
| listing_name | string | Name from listing link |
| listing_url | string | Full URL to detail page |
| city_searched | string | City used in search |
| state_searched | string | State used in search |
| page_number | int | Search results page |
| position_on_page | int | Row index on page |
| name | string | Notary name (from H1) |
| email | string | Email address |
| phone | string | Primary phone |
| mobile | string | Mobile phone (if separate) |
| fax | string | Fax number |
| website | string | Website URL |
| street_address | string | Street address |
| city | string | City from detail page |
| state | string | State from detail page |
| zip_code | string | ZIP code |
| county | string | County |
| commission_number | string | Commission/license number |
| commission_expiration | string | Expiration date (MM/DD/YYYY) |
| bonding_company | string | Bonding company name |
| errors_omissions_insurance | string | E&O insurance info |
| notary_type | string | Type (e.g., "Mobile", "RON", "Signing Agent") |
| services_offered | string | Services description |
| languages_spoken | string | Languages |
| business_hours | string | Hours of operation |
| fees | string | Fee schedule |
| coverage_area | string | Geographic coverage |
| years_in_business | string | Years experience |
| background_check | string | Background check status |
| description | string | Full description / error msg |
| scraped_at | ISO8601 | Timestamp of scrape |
| scraped_from_url | string | Actual URL scraped (may redirect) |
| source_page_number | int | Original search page |
| position_on_page | int | Original position |
| raw_html | string | Truncated HTML (50k chars) |

### Sheet 2: `Summary` (Clean View)
| Column | Source Field |
|--------|-------------|
| Name | `name` |
| City_Searched | `city_searched` |
| State_Searched | `state_searched` |
| Listing_Name | `listing_name` |
| Email | `email` |
| Phone | `phone` |
| Mobile | `mobile` |
| Address | Concatenated: street, city, state zip |
| Commission_Number | `commission_number` |
| Commission_Expiration | `commission_expiration` |
| Services | `services_offered` |
| Scraped_At | `scraped_at` |

### CSV: `notarycafe_listings_*.csv`
- Identical to `All_Listings` sheet
- UTF-8 encoding
- No index column

---

## 📈 Progress Tracking (JSON)

### `scraping_progress_YYYYMMDD_HHMMSS.json`
```json
{
  "session_id": "20261114_143022",
  "completed_cities": ["Indiana_Indianapolis", "Indiana_Bloomington", ...],
  "total_listings_scraped": 1247,
  "last_update": "2026-11-14T14:45:33.123456",
  "city_progress": {
    "Indiana_Indianapolis": {
      "state": "Indiana",
      "city": "Indianapolis",
      "current_page": 3,
      "total_pages": 12,
      "listings_scraped": 45,
      "total_listings": 240,
      "last_listing_index": 44,
      "completed": false,
      "error_count": 0,
      "last_error": "",
      "last_update": "2026-11-14T14:44:12.000000",
      "has_data": true
    }
  }
}
```

### `error_log_YYYYMMDD_HHMMSS.jsonl` (JSON Lines)
```json
{"timestamp": "2026-11-14T14:30:15.123456", "error_type": "search_failed", "message": "Search input not found", "context": {"city": "Indianapolis", "state": "Indiana"}}
{"timestamp": "2026-11-14T14:35:22.123456", "error_type": "detail_extraction", "message": "Timeout on detail page", "context": {"listing_url": "https://www.notarycafe.com/listing/12345", "city": "Indianapolis", "state": "Indiana"}}
```

---

## 🔑 AntiGravity Integration Points

### 1. Data Ingestion Format
```python
# AntiGravity expects list of dicts with these keys:
REQUIRED_FIELDS = [
    'name', 'email', 'phone', 'street_address', 'city', 'state', 'zip_code',
    'commission_number', 'commission_expiration', 'services_offered',
    'listing_url', 'scraped_at'
]

# Optional but recommended:
OPTIONAL_FIELDS = [
    'mobile', 'fax', 'website', 'county', 'bonding_company',
    'errors_omissions_insurance', 'notary_type', 'languages_spoken',
    'business_hours', 'fees', 'coverage_area', 'years_in_business',
    'background_check', 'city_searched', 'state_searched'
]
```

### 2. Direct CSV Import
```python
import pandas as pd

# Load directly into AntiGravity
df = pd.read_csv('notarycafe_listings_20261114_143022.csv')

# Filter valid records
valid = df[df['email'].notna() & (df['email'] != '')]
valid = valid[valid['name'].notna() & (valid['name'] != '')]

# AntiGravity pipeline
for _, row in valid.iterrows():
    antigravity.ingest({
        'source': 'notarycafe',
        'source_id': row['listing_url'].split('/')[-1],
        'data': row.to_dict()
    })
```

### 3. Incremental Sync (Resume Support)
```python
# Read progress to know what's already synced
import json
with open('scraping_progress_20261114_143022.json') as f:
    progress = json.load(f)

last_synced = progress['total_listings_scraped']  # Or track separately
# Only process listings after last_synced
new_listings = all_listings[last_synced:]
```

### 4. Deduplication Key
```python
# Unique identifier for AntiGravity dedup:
dedup_key = f"{row['email']}|{row['name']}|{row['city_searched']}|{row['state_searched']}"
# Or use listing_url if stable:
dedup_key = row['listing_url']
```

---

## ⚙️ Configuration Parameters (Tunable)

| Parameter | Default | Description | AntiGravity Impact |
|-----------|---------|-------------|-------------------|
| `delay_range` | (4, 8) | Seconds between detail pages | Higher = slower but safer |
| `page_delay_range` | (3, 6) | Seconds between search pages | Affects total runtime |
| `batch_size` | 10 | Save every N listings | Lower = more frequent checkpoints |
| `max_retries` | 3 | Retries on failure | Higher = more complete data |
| `chrome_user_data` | Profile 3 | Chrome profile path | Must have NotaryCafe login |
| `headless` | False | Visible browser | Required for auth persistence |

---

## 🚨 Known Limitations & Edge Cases

| Issue | Handling | AntiGravity Note |
|-------|----------|------------------|
| Duplicate listings across cities | Same notary appears in multiple city searches | Dedup on `email + name` |
| Missing email/phone | Field left empty string | Filter before ingest |
| Pagination detection fails | Falls back to page number clicking | May miss last page |
| Detail page structure changes | Regex + JSON-LD fallback | Monitor `error_log` |
| CAPTCHA / Rate limit | Random delays + human scroll | Increase `delay_range` |
| Chrome profile locked | Fails to launch | Close all Chrome windows first |
| Listing URL redirects | `scraped_from_url` captures final | Use for dedup |

---

## 📦 Complete File Manifest for AntiGravity

```
notarycafe-scraper/
├── notarycafe_scraper.py          # Main scraper (copy as-is)
├── cities-basic-v2025_structured.xlsx  # 20,000+ US cities
├── requirements.txt               # Python deps
├── README.md                      # This documentation
├── notarycafe_output/
│   ├── .gitkeep                   # Keeps folder in git
│   ├── notarycafe_listings_*.xlsx # Excel output (2 sheets)
│   ├── notarycafe_listings_*.csv  # CSV for direct import
│   ├── scraping_progress_*.json   # Resume state
│   └── error_log_*.jsonl          # Debug/audit trail
└── .gitignore                     # Excludes output data, cache
```

---

## 🔄 Running as AntiGravity Source

```bash
# 1. One-time setup
git clone https://github.com/plumberqt/notarycafe-scraper.git
cd notarycafe-scraper
python -m venv venv && venv\Scripts\activate
pip install -r requirements.txt
playwright install chromium

# 2. Configure paths in notarycafe_scraper.py (lines 93-95, 117-119)
#    Update chrome_user_data to YOUR machine's Chrome profile

# 3. Run scraper (can interrupt/resume anytime)
python notarycafe_scraper.py

# 4. Feed to AntiGravity (after each batch or at end)
python -c "
import pandas as pd
from antigravity import ingest

df = pd.read_csv('notarycafe_output/notarycafe_listings_latest.csv')
for _, row in df.iterrows():
    if row['email'] and row['name']:
        ingest({
            'source': 'notarycafe',
            'entity_type': 'notary',
            'data': row.to_dict(),
            'dedup_key': row['listing_url']
        })
"
```

---

## ✅ Validation Checklist for AntiGravity

Before ingesting, verify:

- [ ] `email` column has valid format (regex: `^[\w\.-]+@[\w\.-]+\.\w+$`)
- [ ] `phone` column has 10 digits (strip formatting)
- [ ] `commission_expiration` parses as date (MM/DD/YYYY)
- [ ] `zip_code` is 5 or 9 digits
- [ ] `state` is 2-letter uppercase abbreviation
- [ ] `listing_url` starts with `https://www.notarycafe.com/`
- [ ] No duplicate `listing_url` in batch
- [ ] `scraped_at` is recent (within last run)

---

## 📞 Support / Debugging

| Log Location | Purpose |
|--------------|---------|
| `notarycafe_output/error_log_*.jsonl` | Failed extractions, timeouts, CAPTCHAs |
| `notarycafe_output/scraping_progress_*.json` | Resume state, city completion |
| Console output | Real-time progress: `📍 Scraping City, ST...`, `📄 Page 3...`, `📋 [5/20] Name...` |

**Common fixes:**
- "Search input not found" → Site layout changed; update `SEARCH_SELECTORS`
- "No listings on page" → Increase `delay_range` or check for CAPTCHA
- "Detail page timeout" → Increase `timeout` in `page.goto()`
- Chrome won't launch → `taskkill /f /im chrome.exe` then retry

---

*Generated for AntiGravity integration. This document matches the exact implementation in `notarycafe_scraper.py` as of 2026-11-14.*