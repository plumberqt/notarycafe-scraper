# NotaryCafe Complete Automated Directory Scraper

A production-grade, highly resilient web scraper for extracting notary public listings and detailed profiles from [NotaryCafe.com](https://notarycafe.com) using Playwright with anti-bot stealth mechanisms.

---

## 🚀 Key Features

- **Triple Excel Outputs**:
  - `notarycafe_city_stats.xlsx`: **City Summary & Statistics Report** — Tracks every searched city with `City`, `State`, `Total_Listings_Found` (0 if no notaries found, exact count if found), `Total_Pages`, `Status` (`Completed`, `No Results`, `Search Failed`), and timestamp. Ideal for high-level analytics, verification, and reporting.
  - `notarycafe_listings_summary.xlsx`: Quick search results table (10 listings per page) with direct contact info and social media links.
  - `notarycafe_detailed_profiles.xlsx`: Deep profile crawl with all 18 structured sections (Bio, Billing/Shipping, Pricing items, State Licenses, Insurances, Background Checks, Capabilities, Availability, Languages, Social Links).
- **Full Multi-Page Traversal**: For cities with high density (>10 listings across Page 1, Page 2, Page 3, etc.), the scraper automatically paginates through all pages until every single listing is scraped.
- **Social Media Extraction**: Automatically parses profile links for Facebook, LinkedIn, Twitter/X, Instagram, YouTube, TikTok, etc.
- **Smart City Search**: Searches directly by City Name to ensure complete listing coverage.
- **Auto-Resume & State Tracking**: Real-time checkpoint tracking in `progress.json` ensures zero duplicated work across restarts.
- **Resilient & Anti-Detection**:
  - Launches real Chrome with automation flags disabled (`--disable-blink-features=AutomationControlled`).
  - Automatic retry on network/VPN lag without crashing.
  - Graceful handling of Excel file locks if files are open while scraping.

---

## 📁 Repository Structure

```
├── scrape_notarycafe_complete.py         # Main automated scraper script
├── start_notary_scraper.bat              # One-click Windows batch runner
├── cities-basic-v2025_structured.xlsx    # Master US cities dataset (All_Cities_Raw)
├── NOTARYCAFE_SCRAPER_DETAILED_PROMPT.md # Complete technical specification
├── requirements.txt                      # Python dependencies
├── README.md                             # Documentation & Setup guide
└── notarycafe_output/                    # Auto-generated output directory
    ├── notarycafe_city_stats.xlsx        # City-level reporting & listing counts
    ├── notarycafe_listings_summary.xlsx  # Page-level summary cards + social links
    ├── notarycafe_detailed_profiles.xlsx # Full deep profiles (18 sections)
    └── progress.json                     # Real-time resume checkpoint
```

---

## 🛠️ Quick Setup & Installation

### 1. Prerequisites
- Python 3.10+
- Google Chrome installed on your machine
- VPN connected to a **USA location** (NotaryCafe is a US-focused service)

### 2. Install Dependencies
```bash
git clone https://github.com/plumberqt/notarycafe-scraper.git
cd notarycafe-scraper

pip install -r requirements.txt
playwright install chromium
```

---

## 🏃 Running the Scraper

### Method 1: One-Click Launcher (Windows)
Double-click:
```
start_notary_scraper.bat
```
Jab aap launcher open karenge, terminal screen par aapse puchhega:
```
======================================================================
             NOTARYCAFE SCRAPER - EXECUTION MENU
======================================================================
  [1] Fresh Start (Shuru se start karein - Nayi Excel files + Alphabet A to Z)
      -> Purani files ka safe backup ban jayega (notarycafe_output/backup_...)
      -> Teeno brand new empty Excel files banengi
      -> Alphabet 'A' se 'Z' tak complete start hoga

  [2] Resume (Jaha pichli baar chhoda tha wahi se continue karein)
      -> Existing Excel files aur progress.json read karega
      -> Completed cities ko skip karke next pending city se scrape karega
======================================================================
Aapka chunav [1 ya 2, Default: 2]: 
```

### Method 2: Command Line / AntiGravity Terminal
```powershell
# Interactive mode (puchhega 1: Fresh Start ya 2: Resume):
python scrape_notarycafe_complete.py

# Direct Fresh Start (Alphabet A to Z se shuru karein aur purana data backup karein):
python scrape_notarycafe_complete.py --fresh

# Direct Resume (Bina puchhe jaha chhoda tha wahi se continue karein):
python scrape_notarycafe_complete.py --resume

# Test a single city (e.g. Bloomington, IN):
python scrape_notarycafe_complete.py --city "Bloomington" --state "IN" --single-city
```

# Test with a maximum page limit per city (e.g. 1 page):
python scrape_notarycafe_complete.py --city "Bloomington" --state "IN" --single-city --max-pages 1

# Resume from a specific city:
python scrape_notarycafe_complete.py --city "Florala" --state "AL"
```

---

## 📊 Extracted Data Columns & Files

### 1. City Statistics File (`notarycafe_city_stats.xlsx`)
| Column | Description |
|---|---|
| `City` | Name of city searched |
| `State` | State abbreviation (e.g. `AL`, `IN`) |
| `Total_Listings_Found` | Exact count of listings found (0 if no results) |
| `Total_Pages` | Total pages scraped (0 if no results) |
| `Status` | `Completed` (listings scraped), `No Results` (0 found), or `Search Failed` |
| `Scraped_At` | Timestamp of completion |

### 2. Summary File (`notarycafe_listings_summary.xlsx`)
- `City_Searched`, `State_Searched`, `Page_Number`, `Position_on_Page`
- `Name`, `Company_Name`, `City`, `State`, `Zip_Code`
- `Phone`, `Email`, `Profile_URL`
- `Facebook_URL`, `LinkedIn_URL`, `Twitter_X_URL`, `Instagram_URL`, `Other_Social_Links`
- `Scraped_At`

### 3. Detailed Profiles File (`notarycafe_detailed_profiles.xlsx`)
- **Member Info**: `Name`, `Profile_URL`, `Member_ID`, `Pro_Member_Until`, `Membership_Status`, `Zip_Codes_Covered`
- **Intro & Bio**: `Introduction_Text`, `Types_of_Closings`, `Service_Area`, `Background`, `Work_History`, `Interests`, `Personal_Bio_Full`
- **Company**: `Company_Name`, `In_Business_Since`
- **Billing & Shipping**: Preferred, Addresses, City, State, Zip
- **Contact**: `Phone_Work`, `Phone_Mobile`, `Phone_Home`, `Phone_Text_Enabled`, `Email_Work`, `Email_Personal`
- **Pricing**: Individual standard pricing categories, `Pricing_All_Items`, `Other_Pricing_Notes`
- **Credentials**: `Custom_Fields`, `State_Licenses`, `Insurances`, `Background_Provider`, `Background_Conducted`, `Background_Expiration`, `Background_Ref_Number`
- **Capabilities & Hours**: `Capabilities`, `Other_Capability_Info`, `Available_Days`, `Available_Times`, `Other_Availability_Notes`, `Spoken_Languages`
- **Social Media & Links**: `Website_URL`, `Facebook_URL`, `LinkedIn_URL`, `Twitter_X_URL`, `Instagram_URL`, `Other_Social_Links`, `Document_Links`
- **Audit Metadata**: `City_Searched`, `State_Searched`, `Page_Number`, `Position_on_Page`, `Scraped_At`
