# NotaryCafe Directory Scraper

A robust, adaptive web scraper for extracting notary public listings from [NotaryCafe.com](https://www.notarycafe.com) using Playwright with Chrome persistent context.

## Features

- **Adaptive scraping**: Handles variable listings per page, unknown page counts, and empty results
- **Resumable progress**: Saves progress per city, can resume from interruption
- **Incremental saves**: Writes data every 10 listings (configurable)
- **Anti-detection**: Uses real Chrome profile with persistent context
- **Comprehensive data extraction**: Gets all available fields from detail pages
- **Error handling**: Detailed error logging with context
- **Multi-format output**: Excel (.xlsx) with multiple sheets + CSV

## Data Extracted

### Summary Fields (from search results)
- Listing name & URL
- City/State searched
- Page number & position

### Detail Fields (from individual listing pages)
- **Contact**: Name, Email, Phone, Mobile, Fax, Website
- **Address**: Street, City, State, ZIP, County
- **Commission**: Number, Expiration date
- **Credentials**: Bonding company, E&O insurance, Notary type
- **Services**: Services offered, Languages spoken
- **Business**: Hours, Fees, Coverage area
- **Background**: Years in business, Background check status
- **Metadata**: Scraped timestamp, source URL, raw HTML

## Setup

### Prerequisites
- Python 3.10+
- Google Chrome installed
- Existing Chrome profile with NotaryCafe login (if required)

### Installation

```bash
# Clone the repo
git clone <your-repo-url>
cd notarycafe-scraper

# Create virtual environment
python -m venv venv
venv\Scripts\activate  # Windows
# source venv/bin/activate  # Mac/Linux

# Install dependencies
pip install -r requirements.txt

# Install Playwright browsers
playwright install chromium
```

### Configuration

Edit the `config` dictionary in `notarycafe_scraper.py` or create a config file:

```python
config = {
    'cities_file': r'D:\Karan\cities-basic-v2025_structured.xlsx',
    'chrome_user_data': r'C:\Users\YOUR_USER\AppData\Local\Google\Chrome\User Data',
    'profile_dir': 'Profile 3',  # Your Chrome profile with NotaryCafe login
    'output_dir': r'D:\Karan\notarycafe_output',
    'delay_range': (4, 8),        # Delay between listings (seconds)
    'page_delay_range': (3, 6),   # Delay between pages (seconds)
    'max_retries': 3,
    'batch_size': 10              # Save every N listings
}
```

### Cities File Format

The scraper expects an Excel file with a sheet named `All_Cities_Raw` containing:
- `State` - Full state name
- `State_Abbr` - State abbreviation (optional)
- `City` - City name
- `GeoID` - Geographic ID (optional)

## Usage

```bash
# Run the scraper
python notarycafe_scraper.py
```

The scraper will:
1. Load cities from the Excel file
2. Resume from previous progress (if exists)
3. Open Chrome with your profile
4. Search each city on NotaryCafe
5. Navigate through all pages
6. Click each listing and extract full details
7. Save incrementally to Excel + CSV

## Output Files

```
notarycafe_output/
├── notarycafe_listings_YYYYMMDD_HHMMSS.xlsx   # Main data (2 sheets: All_Listings, Summary)
├── notarycafe_listings_YYYYMMDD_HHMMSS.csv    # CSV export
├── scraping_progress_YYYYMMDD_HHMMSS.json     # Resume progress
└── error_log_YYYYMMDD_HHMMSS.jsonl            # Error details
```

## Key Implementation Details

### Adaptive Element Detection
The scraper tries multiple CSS selector strategies to find listings:
- Table rows (`table tbody tr`)
- Listing/result/card divs
- Direct links to `/listing/` or `/profile/`

### Pagination Handling
- Detects pagination containers automatically
- Finds "Next" buttons via multiple strategies
- Falls back to page number clicking
- Handles "no results" pages gracefully

### Detail Page Extraction
Uses multiple strategies:
1. Text pattern matching (regex for emails, phones, addresses)
2. JSON-LD structured data parsing
3. HTML attribute scanning (mailto:, tel:)
4. Semantic label matching (Services:, Hours:, etc.)

### Progress Persistence
- Tracks completed cities
- Saves current page per city
- Records last listing index processed
- Counts errors per city
- Auto-saves every batch + on completion

## Running from Home

Since the repo includes:
- `notarycafe_scraper.py` - Main scraper
- `cities-basic-v2025_structured.xlsx` - Cities data
- `requirements.txt` - Dependencies
- `notarycafe_output/` - Output folder structure

You can clone on any machine and run after:
1. Installing Python + dependencies
2. Updating Chrome user data path in config
3. Having a Chrome profile with NotaryCafe access

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Chrome won't launch | Close all Chrome windows first; check profile path |
| No listings found | Check if logged into NotaryCafe in that profile |
| Stuck on page | Increase `delay_range` in config |
| Missing data | Site structure changed; check selectors in `find_listing_elements()` |
| Permission errors | Run terminal as admin; check folder permissions |

## License

MIT License - Feel free to use and modify.

---

**Note**: This scraper is for educational/research purposes. Respect NotaryCafe's Terms of Service and robots.txt. Add appropriate delays and don't overload their servers.