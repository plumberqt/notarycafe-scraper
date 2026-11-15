"""
NotaryCafe Complete Automated Scraper
====================================
Scrapes notary listings and detailed profiles from https://notarycafe.com/
Generates three structured Excel files:
  1. notarycafe_city_stats.xlsx         (City summary: City, State, Total_Listings_Found, Total_Pages, Status)
  2. notarycafe_listings_summary.xlsx   (10 listings per page summary + social links)
  3. notarycafe_detailed_profiles.xlsx  (Full deep profile data for each notary across all sections)

Supports:
  - Interactive Mode: Prompts user whether to Start Fresh (New Excel files + Alphabet A to Z) or Resume from checkpoint
  - Automatic Safe Backup: Backs up previous Excel files and progress into timestamped backup folder on Fresh start
  - Alphabetical Sorting: Scrapes all US cities alphabetically from A to Z (Abbeville -> Boston -> ... -> Zwolle)
  - Automatic multi-page pagination through ALL pages per city (Page 1, 2, 3...) until listings end
  - Tracking 0 results for cities with no listings and exact listing/page counts for cities with listings
  - Opening each profile in a new tab, extracting all fields, closing tab, returning to results
  - Extracting social media links (Facebook, LinkedIn, Twitter/X, Instagram, YouTube, TikTok, etc.)
  - Auto-resuming from progress.json
  - Visible Chrome browser with anti-detection flags for VPN / Cloudflare compatibility
"""

import asyncio
import json
import os
import re
import sys
import shutil
import random
from pathlib import Path
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright

# ----------------- CONFIGURATION -----------------
BASE_URL = "https://notarycafe.com"
SEARCH_URL = "https://notarycafe.com/find-a-notary"
BASE_DIR = Path(__file__).resolve().parent
CITIES_EXCEL = BASE_DIR / "cities-basic-v2025_structured.xlsx"
OUTPUT_DIR = BASE_DIR / "notarycafe_output"
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

CITY_STATS_EXCEL = OUTPUT_DIR / "notarycafe_city_stats.xlsx"
SUMMARY_EXCEL = OUTPUT_DIR / "notarycafe_listings_summary.xlsx"
DETAILED_EXCEL = OUTPUT_DIR / "notarycafe_detailed_profiles.xlsx"
PROGRESS_FILE = OUTPUT_DIR / "progress.json"

# Social media detection patterns
SOCIAL_PATTERNS = {
    'facebook': [r'https?://(?:www\.)?facebook\.com/[A-Za-z0-9\.\-_]+', r'https?://(?:www\.)?fb\.com/[A-Za-z0-9\.\-_]+'],
    'linkedin': [r'https?://(?:www\.)?linkedin\.com/in/[A-Za-z0-9\-\_]+', r'https?://(?:www\.)?linkedin\.com/company/[A-Za-z0-9\-\_]+'],
    'twitter': [r'https?://(?:www\.)?twitter\.com/[A-Za-z0-9\_]+', r'https?://(?:www\.)?x\.com/[A-Za-z0-9\_]+'],
    'instagram': [r'https?://(?:www\.)?instagram\.com/[A-Za-z0-9\.\-_]+'],
    'youtube': [r'https?://(?:www\.)?youtube\.com/(?:c/|channel/|user/|@)[A-Za-z0-9\-\_]+'],
    'tiktok': [r'https?://(?:www\.)?tiktok\.com/@[A-Za-z0-9\.\-_]+'],
}

def backup_existing_files():
    """Safely moves existing output Excel files and progress.json into a timestamped backup folder."""
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    backup_dir = OUTPUT_DIR / f"backup_{timestamp}"
    
    files_to_backup = [CITY_STATS_EXCEL, SUMMARY_EXCEL, DETAILED_EXCEL, PROGRESS_FILE]
    existing = [f for f in files_to_backup if f.exists()]
    
    if not existing:
        print("[INFO] No previous output files found. Ready for a clean fresh start.")
        return
        
    backup_dir.mkdir(parents=True, exist_ok=True)
    
    while True:
        try:
            for f in existing:
                shutil.copy2(f, backup_dir / f.name)
            for f in existing:
                if f.exists():
                    f.unlink()
            print(f"\n[BACKUP SUCCESS] Purani files ka safe backup ban gaya: notarycafe_output/{backup_dir.name}/")
            print("[FRESH START] Active output folder reset ho gaya. Nayi fresh Excel files banengi.\n")
            break
        except PermissionError as pe:
            print(f"\n[FILE LOCKED] Ek ya zyada Excel files open hain: {pe}")
            print("Kripya open Excel files (notarycafe_*.xlsx) ko CLOSE karein aur Enter dabayein...")
            try:
                input()
            except Exception:
                break


def prompt_user_mode():
    """Prompts the user interactively in console whether to start fresh or resume."""
    print("\n" + "=" * 70)
    print("             NOTARYCAFE SCRAPER - EXECUTION MENU")
    print("=" * 70)
    print("  [1] Fresh Start (Shuru se start karein - Nayi Excel files + Alphabet A to Z)")
    print("      -> Purani files ka safe backup ban jayega (notarycafe_output/backup_...)")
    print("      -> Teeno brand new empty Excel files banengi")
    print("      -> Alphabet 'A' se 'Z' tak complete start hoga")
    print()
    print("  [2] Resume (Jaha pichli baar chhoda tha wahi se continue karein)")
    print("      -> Existing Excel files aur progress.json read karega")
    print("      -> Completed cities ko skip karke next pending city se scrape karega")
    print("=" * 70)
    
    choice = ""
    try:
        choice = input("Aapka chunav [1 ya 2, Default: 2]: ").strip()
    except (EOFError, Exception):
        choice = "2"
        
    if choice in ["1", "fresh", "f", "new", "shuru"]:
        print("\nCity sorting order kaisa chahiye?")
        print("  [1] City Name Alphabet A to Z (Abbeville -> Boston -> ... -> Zwolle) [Recommended]")
        print("  [2] State-wise A to Z (Alabama A-Z -> Alaska A-Z -> ... -> Wyoming A-Z)")
        try:
            sort_choice = input("Order chunav [1 ya 2, Default: 1]: ").strip()
        except (EOFError, Exception):
            sort_choice = "1"
            
        sort_by = "state" if sort_choice == "2" else "city"
        return "fresh", sort_by
    else:
        return "resume", "city"


def load_cities_from_excel(target_city=None, target_state=None, sort_by="city"):
    """Load cities from Excel file.
    If sort_by == 'city', sorts alphabetically by city name (A to Z).
    If sort_by == 'state', sorts state-wise (AL to WY).
    If target_city specified, moves that city to the front."""
    print(f"Loading cities from: {CITIES_EXCEL}")
    df = pd.read_excel(CITIES_EXCEL, sheet_name="All_Cities_Raw")
    
    if sort_by == "city":
        df = df.sort_values(
            by=["City", "State"],
            key=lambda col: col.str.lower() if col.name == "City" else col
        )
    elif sort_by == "state":
        df = df.sort_values(
            by=["State", "City"],
            key=lambda col: col.str.lower() if col.name == "City" else col
        )
        
    cities = []
    seen = set()
    for _, row in df.iterrows():
        c_name = str(row["City"]).strip()
        s_name = str(row["State"]).strip()
        key = (s_name, c_name)
        if key not in seen and c_name and s_name:
            seen.add(key)
            cities.append({
                "city": c_name,
                "state": s_name,
                "population": row.get("Population_2025", 0)
            })
            
    print(f"Loaded {len(cities)} unique cities from file (Order: {sort_by.upper()}).")
    
    if target_city:
        target_idx = None
        for i, c in enumerate(cities):
            if c["city"].lower() == target_city.lower() and (not target_state or c["state"].upper() == target_state.upper()):
                target_idx = i
                break
        if target_idx is not None:
            match = cities.pop(target_idx)
            cities.insert(0, match)
            print(f"Set starting city to: {match['city']}, {match['state']}")
        else:
            cities.insert(0, {"city": target_city, "state": target_state or "", "population": 0})
            print(f"Added starting city: {target_city}, {target_state or ''}")
            
    return cities


class NotaryCafeScraper:
    def __init__(self, cities, start_city=None, start_state=None, max_pages=None, single_city=False):
        self.cities = cities
        self.start_city = start_city
        self.start_state = start_state
        self.max_pages = max_pages
        self.single_city = single_city
        self.completed_cities = set()
        self.completed_profiles = set()
        self.city_stats_records = []
        self.summary_records = []
        self.detailed_records = []
        self.load_existing_data()

    def load_existing_data(self):
        """Loads existing Excel files and progress.json if present to prevent duplicates and maintain state."""
        # 1. Load progress.json
        if PROGRESS_FILE.exists():
            try:
                with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                    p = json.load(f)
                    self.completed_cities = set(p.get("completed_cities", []))
                    self.completed_profiles = set(p.get("completed_profiles", []))
                    print(f"Resumed from progress.json: {len(self.completed_cities)} completed cities, {len(self.completed_profiles)} profiles.")
            except Exception as e:
                print(f"Warning reading progress.json: {e}")

        # 2. Load Listings Summary Excel
        if SUMMARY_EXCEL.exists():
            try:
                df_sum = pd.read_excel(SUMMARY_EXCEL)
                self.summary_records = df_sum.to_dict("records")
                print(f"Loaded {len(self.summary_records)} existing summary records.")
            except Exception as e:
                print(f"Warning reading existing summary excel: {e}")

        # 3. Load Detailed Profiles Excel
        if DETAILED_EXCEL.exists():
            try:
                df_det = pd.read_excel(DETAILED_EXCEL)
                self.detailed_records = df_det.to_dict("records")
                for r in self.detailed_records:
                    if r.get("Profile_URL"):
                        self.completed_profiles.add(r["Profile_URL"])
                print(f"Loaded {len(self.detailed_records)} existing detailed records.")
            except Exception as e:
                print(f"Warning reading existing detailed excel: {e}")

        # 4. Load City Statistics Excel
        if CITY_STATS_EXCEL.exists():
            try:
                df_stats = pd.read_excel(CITY_STATS_EXCEL)
                self.city_stats_records = df_stats.to_dict("records")
                print(f"Loaded {len(self.city_stats_records)} existing city stats records.")
            except Exception as e:
                print(f"Warning reading existing city stats excel: {e}")

        # 5. Backfill previously completed cities into city_stats_records if missing
        existing_stats_keys = {(str(r.get("City")).strip(), str(r.get("State")).strip()) for r in self.city_stats_records}
        summary_city_counts = {}
        summary_city_max_pages = {}
        for r in self.summary_records:
            c = str(r.get("City_Searched", "")).strip()
            s = str(r.get("State_Searched", "")).strip()
            pg = int(r.get("Page_Number", 1)) if pd.notna(r.get("Page_Number")) else 1
            if c and s:
                summary_city_counts[(c, s)] = summary_city_counts.get((c, s), 0) + 1
                summary_city_max_pages[(c, s)] = max(summary_city_max_pages.get((c, s), 1), pg)

        backfilled = 0
        for c_key in self.completed_cities:
            parts = c_key.split("_")
            c_name = parts[0].strip()
            s_name = parts[1].strip() if len(parts) > 1 else ""
            if (c_name, s_name) not in existing_stats_keys:
                cnt = summary_city_counts.get((c_name, s_name), 0)
                max_pg = summary_city_max_pages.get((c_name, s_name), 1) if cnt > 0 else 0
                self.city_stats_records.append({
                    "City": c_name,
                    "State": s_name,
                    "Total_Listings_Found": cnt,
                    "Total_Pages": max_pg,
                    "Status": "Completed" if cnt > 0 else "No Results",
                    "Scraped_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                })
                existing_stats_keys.add((c_name, s_name))
                backfilled += 1

        if backfilled > 0:
            print(f"Backfilled {backfilled} existing completed cities into city statistics report.")
            self.save_data()

    def save_data(self):
        """Saves current state safely to all three Excel files and progress.json."""
        # 1. Save progress.json first (lightweight, crash-safe)
        try:
            progress = {
                "completed_cities": list(self.completed_cities),
                "completed_profiles": list(self.completed_profiles),
                "city_stats_count": len(self.city_stats_records),
                "summary_count": len(self.summary_records),
                "detailed_count": len(self.detailed_records),
                "last_updated": datetime.now().isoformat()
            }
            with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                json.dump(progress, f, indent=2)
        except Exception as pe:
            print(f"Warning saving progress.json: {pe}")

        # 2. Save City Statistics Excel
        try:
            if self.city_stats_records:
                df_stats = pd.DataFrame(self.city_stats_records)
                df_stats.drop_duplicates(subset=["City", "State"], keep="last", inplace=True)
                df_stats["Total_Listings_Found"] = pd.to_numeric(df_stats["Total_Listings_Found"], errors="coerce").fillna(0).astype(int)
                df_stats["Total_Pages"] = pd.to_numeric(df_stats["Total_Pages"], errors="coerce").fillna(0).astype(int)
                df_stats.to_excel(CITY_STATS_EXCEL, index=False)
        except PermissionError:
            print(f"Notice: {CITY_STATS_EXCEL.name} is currently open in Excel. Skipping write this turn (data safe in memory).")
        except Exception as ste:
            print(f"Warning saving city stats excel: {ste}")

        # 3. Save Summary Excel (handle file lock if user has file open)
        try:
            if self.summary_records:
                df_s = pd.DataFrame(self.summary_records)
                df_s.drop_duplicates(subset=["Profile_URL", "City_Searched"], keep="last", inplace=True)
                df_s.to_excel(SUMMARY_EXCEL, index=False)
        except PermissionError:
            print(f"Notice: {SUMMARY_EXCEL.name} is currently open in Excel. Skipping write this turn (data safe in memory).")
        except Exception as se:
            print(f"Warning saving summary excel: {se}")

        # 4. Save Detailed Profiles Excel (cast object types to avoid LossySetitemError)
        try:
            if self.detailed_records:
                df_d = pd.DataFrame(self.detailed_records)
                for col in df_d.columns:
                    df_d[col] = df_d[col].astype(object)
                df_d.drop_duplicates(subset=["Profile_URL"], keep="last", inplace=True)
                df_d.to_excel(DETAILED_EXCEL, index=False)
        except PermissionError:
            print(f"Notice: {DETAILED_EXCEL.name} is currently open in Excel. Skipping write this turn (data safe in memory).")
        except Exception as de:
            print(f"Warning saving detailed excel: {de}")

    async def human_delay(self, min_s=1.5, max_s=3.5):
        await asyncio.sleep(random.uniform(min_s, max_s))

    async def search_city(self, page, city, state, retries=3):
        """Searches for a city on NotaryCafe (only typing city name, with auto-retry)."""
        query = city.strip()
        print(f"\n========================================================")
        print(f"Searching for city: {query} (Target State: {state})")
        print(f"========================================================")
        
        for attempt in range(1, retries + 1):
            try:
                await page.goto(SEARCH_URL, wait_until="domcontentloaded", timeout=45000)
                await self.human_delay(1.5, 3.0)
                
                search_selectors = [
                    'input[placeholder*="zip, city, state" i]',
                    'input[placeholder*="Zip Code or City" i]',
                    'input[name="search"]',
                    'input[name="q"]',
                    'input[type="text"]'
                ]
                
                search_input = None
                for sel in search_selectors:
                    el = await page.query_selector(sel)
                    if el and await el.is_visible():
                        search_input = el
                        break
                        
                if not search_input:
                    print(f"Attempt {attempt}: Search input not visible, retrying...")
                    await asyncio.sleep(3)
                    continue
                    
                await search_input.click()
                await search_input.fill("")
                await self.human_delay(0.2, 0.5)
                
                for char in query:
                    await search_input.type(char, delay=random.randint(30, 80))
                await self.human_delay(0.4, 0.8)
                
                search_button = await page.query_selector('button:has-text("Search"), input[type="submit"][value*="Search"]')
                if search_button and await search_button.is_visible():
                    await search_button.click()
                else:
                    await search_input.press("Enter")
                    
                print("Search submitted. Waiting for results...")
                try:
                    await page.wait_for_load_state("domcontentloaded", timeout=30000)
                except Exception:
                    pass
                await self.human_delay(2.0, 3.5)
                return True
                
            except Exception as e:
                print(f"Attempt {attempt}/{retries} failed searching for {city}: {e}")
                if attempt < retries:
                    await asyncio.sleep(4)
                else:
                    return False
        return False

    async def parse_search_listings(self, page, city, state, page_num):
        """Parses the listing cards visible on the search results page."""
        listings = []
        
        # Check for no results indicators
        body_text = await page.inner_text("body")
        no_res_indicators = [
            "no results", "0 results", "no notaries found", 
            "no matches found", "0 listings", "nothing found", 
            "no notaries in"
        ]
        if any(ind in body_text.lower() for ind in no_res_indicators):
            print(f"No notaries found for {city}, {state}.")
            return []
            
        # Find all 'View Profile' links
        view_profile_links = await page.query_selector_all('a:has-text("View Profile")')
        if not view_profile_links:
            links = await page.query_selector_all('a[href]')
            view_profile_links = []
            for l in links:
                txt = (await l.inner_text()).strip()
                if "view profile" in txt.lower():
                    view_profile_links.append(l)
                    
        if not view_profile_links:
            return []

        print(f"Found {len(view_profile_links)} listings on Page {page_num}.")
        
        for idx, link_el in enumerate(view_profile_links, start=1):
            try:
                href = await link_el.get_attribute("href")
                if not href or href.strip() in ["#", "", "/"] or href.endswith("#"):
                    continue
                clean_href = href.strip().lower()
                if any(clean_href == bad or clean_href.startswith(bad) for bad in ["/register", "/plan", "/about", "/contact", "/terms", "/privacy", "/forum", "http://facebook.com", "http://twitter.com"]):
                    continue
                profile_url = href if href.startswith("http") else f"{BASE_URL}{href}"
                
                # Find parent card/container of this link
                card_text = ""
                card_el = await link_el.evaluate_handle('el => el.closest("div, li, tr") ? el.closest("div, li, tr").innerText : ""')
                card_text = str(card_el) if card_el else ""
                
                card_container = await link_el.evaluate_handle('''el => {
                    let p = el.parentElement;
                    while (p && p.innerText && p.innerText.length < 50 && p.parentElement) {
                        p = p.parentElement;
                    }
                    return p ? p.innerText : "";
                }''')
                card_full_text = str(card_container) if card_container else card_text
                
                name = ""
                company = ""
                c_city = ""
                c_state = ""
                c_zip = ""
                c_phone = ""
                c_email = ""
                
                lines = [l.strip() for l in card_full_text.split("\n") if l.strip() and "miles" not in l.lower() and "profile completion" not in l.lower() and "view profile" not in l.lower()]
                if lines:
                    name = lines[0]
                    if len(lines) > 1 and not any(ch.isdigit() for ch in lines[1]):
                        company = lines[1]
                        
                phone_match = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', card_full_text)
                if phone_match:
                    c_phone = phone_match.group(0)
                    
                email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', card_full_text)
                if email_match:
                    c_email = email_match.group(0)
                    
                loc_match = re.search(r'([A-Za-z\s]+)\s+([A-Z]{2})\s+(\d{5})', card_full_text)
                if loc_match:
                    c_city = loc_match.group(1).strip()
                    c_state = loc_match.group(2).strip()
                    c_zip = loc_match.group(3).strip()
                else:
                    c_city = city
                    c_state = state

                card_social = self.extract_social_links(card_full_text)
                
                item = {
                    "City_Searched": city,
                    "State_Searched": state,
                    "Page_Number": page_num,
                    "Position_on_Page": idx,
                    "Name": name,
                    "Company_Name": company,
                    "City": c_city,
                    "State": c_state,
                    "Zip_Code": c_zip,
                    "Phone": c_phone,
                    "Email": c_email,
                    "Profile_URL": profile_url,
                    "Facebook_URL": card_social.get("facebook", ""),
                    "LinkedIn_URL": card_social.get("linkedin", ""),
                    "Twitter_X_URL": card_social.get("twitter", ""),
                    "Instagram_URL": card_social.get("instagram", ""),
                    "Other_Social_Links": card_social.get("other", ""),
                    "Scraped_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                }
                listings.append(item)
            except Exception as e:
                print(f"Error parsing card {idx}: {e}")
                
        return listings

    def extract_social_links(self, text_or_html):
        """Extracts social media profile links from text or HTML."""
        found = {}
        for platform, patterns in SOCIAL_PATTERNS.items():
            for pat in patterns:
                m = re.search(pat, text_or_html, re.IGNORECASE)
                if m:
                    url = m.group(0)
                    if "notarycafe" not in url.lower() and "sharer" not in url.lower() and "intent" not in url.lower():
                        found[platform] = url
                        break
        return found

    async def scrape_detail_page(self, context, profile_url, summary_info):
        """Opens the profile page in a new tab, extracts all structured fields, and closes the tab."""
        print(f"    -> Opening profile: {profile_url}")
        detail_page = await context.new_page()
        
        try:
            await detail_page.goto(profile_url, wait_until="domcontentloaded", timeout=45000)
            await self.human_delay(1.5, 3.0)
            
            # Human-like scroll to trigger dynamic content
            await detail_page.evaluate("""
                () => {
                    window.scrollTo(0, document.body.scrollHeight / 2);
                }
            """)
            await asyncio.sleep(0.5)
            await detail_page.evaluate("""
                () => {
                    window.scrollTo(0, document.body.scrollHeight);
                }
            """)
            await asyncio.sleep(0.5)
            await detail_page.evaluate("() => window.scrollTo(0, 0)")
            
            page_text = await detail_page.inner_text("body")
            page_html = await detail_page.content()
            
            # 1. Name & Member Info
            h1 = await detail_page.query_selector("h1")
            name = (await h1.inner_text()).strip() if h1 else summary_info.get("Name", "")
            
            member_id = ""
            m_id = re.search(r'Id:\s*([A-Za-z0-9]+)', page_text, re.IGNORECASE)
            if m_id: member_id = m_id.group(1).strip()
            
            pro_member_until = ""
            m_pro = re.search(r'Pro Member Until:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_pro: pro_member_until = m_pro.group(1).strip()
            
            status = ""
            m_stat = re.search(r'Status:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_stat: status = m_stat.group(1).strip()
            
            zip_codes = ""
            m_zips = re.search(r'Zip Codes:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_zips: zip_codes = m_zips.group(1).strip()
            
            # 2. Professional Introduction & Experience
            intro_text = ""
            closings_types = ""
            service_area = ""
            
            m_intro = re.search(r'Professional Introduction & Experience\s*\n+([^\n]+(?:\n+[^\n]+){0,5})', page_text)
            if m_intro: intro_text = m_intro.group(1).strip()
            
            m_clos = re.search(r'(?:Types of Closings|Closings):\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_clos: closings_types = m_clos.group(1).strip()
            
            m_serv = re.search(r'(?:Service Area|Radius):\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_serv: service_area = m_serv.group(1).strip()
            
            # 3. Personal Bio
            personal_bio = ""
            background = ""
            work_history = ""
            interests = ""
            
            m_bio = re.search(r'Personal Bio\s*\n+([^\n]+(?:\n+[^\n]+){0,4})', page_text)
            if m_bio: personal_bio = m_bio.group(1).strip()
            
            m_bg = re.search(r'Background:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bg: background = m_bg.group(1).strip()
            
            m_wh = re.search(r'Work History:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_wh: work_history = m_wh.group(1).strip()
            
            m_int = re.search(r'Interests:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_int: interests = m_int.group(1).strip()
            
            # 4. Name & Company
            company_name = summary_info.get("Company_Name", "")
            m_comp = re.search(r'Company Name:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_comp and not company_name: company_name = m_comp.group(1).strip()
            
            in_business_since = ""
            m_bus = re.search(r'In Business Since:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bus: in_business_since = m_bus.group(1).strip()
            
            # 5. Billing Address
            b_preferred = "Yes" if "Preferred" in page_text else ""
            b_addr1 = ""
            b_addr2 = ""
            b_city = summary_info.get("City", "")
            b_state = summary_info.get("State", "")
            b_zip = summary_info.get("Zip_Code", "")
            
            m_b_addr = re.search(r'Billing\s*\n+(?:Preferred:?\s*\w+\s*\n+)?([0-9]+[^\n]+)\n+([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5})', page_text)
            if m_b_addr:
                b_addr1 = m_b_addr.group(1).strip()
                b_city = m_b_addr.group(2).strip()
                b_state = m_b_addr.group(3).strip()
                b_zip = m_b_addr.group(4).strip()
                
            # 6. Shipping Address
            s_same = "Yes" if "Same as Billing" in page_text else ""
            
            # 7. Phone Numbers
            p_work = summary_info.get("Phone", "")
            p_mobile = ""
            p_home = ""
            p_text_enabled = "Yes" if "text enabled" in page_text.lower() else "No"
            
            m_work_ph = re.search(r'Work:\s*([0-9\(\)\-\.\s]+)', page_text)
            if m_work_ph: p_work = m_work_ph.group(1).strip()
            
            m_mob_ph = re.search(r'Mobile:\s*([0-9\(\)\-\.\s]+)', page_text)
            if m_mob_ph: p_mobile = m_mob_ph.group(1).strip()
            
            m_hm_ph = re.search(r'Home:\s*([0-9\(\)\-\.\s]+)', page_text)
            if m_hm_ph: p_home = m_hm_ph.group(1).strip()
            
            # 8. Email Addresses
            e_work = summary_info.get("Email", "")
            e_personal = ""
            
            m_em_all = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', page_text)
            clean_emails = [e for e in m_em_all if "notarycafe" not in e.lower()]
            if clean_emails:
                e_work = clean_emails[0]
                if len(clean_emails) > 1:
                    e_personal = clean_emails[1]
                    
            # 9. Pricing Information
            price_purchase = ""
            price_seller = ""
            price_modification = ""
            price_heloc = ""
            price_reverse = ""
            pricing_notes = ""
            
            m_p1 = re.search(r'Purchase or refi[^\$]*(\$[\d\.]+)', page_text, re.IGNORECASE)
            if m_p1: price_purchase = m_p1.group(1).strip()
            
            m_p2 = re.search(r'Seller pkg[^\$]*(\$[\d\.]+)', page_text, re.IGNORECASE)
            if m_p2: price_seller = m_p2.group(1).strip()
            
            m_p3 = re.search(r'Modification[^\$]*(\$[\d\.]+)', page_text, re.IGNORECASE)
            if m_p3: price_modification = m_p3.group(1).strip()
            
            m_p4 = re.search(r'Home equity lines[^\$]*(\$[\d\.]+)', page_text, re.IGNORECASE)
            if m_p4: price_heloc = m_p4.group(1).strip()
            
            m_p5 = re.search(r'Reverse mortgage[^\$]*(\$[\d\.]+)', page_text, re.IGNORECASE)
            if m_p5: price_reverse = m_p5.group(1).strip()
            
            pricing_matches = re.findall(r'(\d+\.\s*[^\$\n]+:\s*\$[\d\.]+)', page_text)
            pricing_all_str = "; ".join(pricing_matches) if pricing_matches else ""
            
            # 10. Custom Fields
            custom_matches = re.findall(r'Custom Field \d+:\s*([^\n]+)', page_text, re.IGNORECASE)
            custom_fields_str = "; ".join(custom_matches) if custom_matches else ""
            
            # 11. State Licenses
            licenses = re.findall(r'([A-Z]{2}\s*License\s*#?\d*:[^\n]+)', page_text, re.IGNORECASE)
            licenses_str = "; ".join(licenses) if licenses else ""
            
            # 12. Insurances
            insurances = re.findall(r'(?:Carrier and Amount|Insurance|Bonding):\s*([^\n]+)', page_text, re.IGNORECASE)
            insurances_str = "; ".join(insurances) if insurances else ""
            
            # 13. Background Checks
            bg_provider = ""
            bg_conducted = ""
            bg_expiration = ""
            bg_ref = ""
            
            m_bg_prov = re.search(r'Provider:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bg_prov: bg_provider = m_bg_prov.group(1).strip()
            
            m_bg_cond = re.search(r'Conducted:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bg_cond: bg_conducted = m_bg_cond.group(1).strip()
            
            m_bg_exp = re.search(r'Expiration:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bg_exp: bg_expiration = m_bg_exp.group(1).strip()
            
            m_bg_ref = re.search(r'Ref\.?\s*#:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_bg_ref: bg_ref = m_bg_ref.group(1).strip()
            
            # 14. Capabilities
            capabilities_list = [
                "CSA", "Attorney", "Fax", "Email", "Internet", "Laser Printer", 
                "Abstractor", "Mobile Hotspot", "E-Sign", "24 Hour Service", 
                "Fingerprinting", "Weddings", "Hospital Signing", "Jail Signings", 
                "EscrowTab", "IPEN", "RON Capable"
            ]
            found_caps = [c for c in capabilities_list if c.lower() in page_text.lower()]
            capabilities_str = ", ".join(found_caps)
            
            other_caps = ""
            m_caps = re.search(r'Capabilities\s*\n+([^\n]+(?:\n+[^\n]+){0,2})', page_text)
            if m_caps: other_caps = m_caps.group(1).strip()
            
            # 15. Availability
            days_list = ["Mon", "Tue", "Wed", "Thu", "Fri", "Sat", "Sun"]
            active_days = [d for d in days_list if d in page_text]
            days_str = ", ".join(active_days) if active_days else "Mon-Fri"
            
            times_str = ""
            if "AM" in page_text and "PM" in page_text: times_str = "AM, PM"
            elif "AM" in page_text: times_str = "AM"
            elif "PM" in page_text: times_str = "PM"
            
            avail_notes = ""
            m_avnotes = re.search(r'Other Availability Notes\s*\n+([^\n]+)', page_text, re.IGNORECASE)
            if m_avnotes: avail_notes = m_avnotes.group(1).strip()
            
            # 16. Spoken Languages
            languages = "English"
            m_lang = re.search(r'Spoken Languages\s*\n+([A-Za-z,\s]+)', page_text)
            if m_lang: languages = m_lang.group(1).strip()
            
            # 17. Websites & Social Media
            website_url = ""
            m_web = re.search(r'Websites\s*\n+(?:URL\s*\n+)?([a-zA-Z0-9\-\.]+\.[a-zA-Z]{2,}(?:/[^\s]*)?)', page_text)
            if m_web: website_url = m_web.group(1).strip()
            
            socials = self.extract_social_links(page_html)
            
            # 18. Document Links
            doc_links = []
            for doc in ["Bond", "Commission", "E&O", "Insurance", "State License", "W9"]:
                if doc in page_text:
                    doc_links.append(doc)
            docs_str = ", ".join(doc_links)
            
            detail_record = {
                # Name & Member Info
                "Name": name,
                "Profile_URL": profile_url,
                "Member_ID": member_id,
                "Pro_Member_Until": pro_member_until,
                "Membership_Status": status,
                "Zip_Codes_Covered": zip_codes,
                
                # Professional Introduction & Experience
                "Introduction_Text": intro_text,
                "Types_of_Closings": closings_types,
                "Service_Area": service_area,
                
                # Personal Bio
                "Background": background,
                "Work_History": work_history,
                "Interests": interests,
                "Personal_Bio_Full": personal_bio,
                
                # Name & Company
                "Company_Name": company_name,
                "In_Business_Since": in_business_since,
                
                # Billing Address
                "Billing_Preferred": b_preferred,
                "Billing_Address_1": b_addr1,
                "Billing_Address_2": b_addr2,
                "Billing_City": b_city,
                "Billing_State": b_state,
                "Billing_Zip": b_zip,
                
                # Shipping Address
                "Shipping_Same_as_Billing": s_same,
                "Shipping_Address_1": b_addr1 if s_same == "Yes" else "",
                "Shipping_City": b_city if s_same == "Yes" else "",
                "Shipping_State": b_state if s_same == "Yes" else "",
                "Shipping_Zip": b_zip if s_same == "Yes" else "",
                
                # Phone Numbers
                "Phone_Work": p_work,
                "Phone_Mobile": p_mobile,
                "Phone_Home": p_home,
                "Phone_Text_Enabled": p_text_enabled,
                
                # Email Addresses
                "Email_Work": e_work,
                "Email_Personal": e_personal,
                
                # Pricing Information
                "Pricing_Purchase_or_Refi": price_purchase,
                "Pricing_Seller_Pkg": price_seller,
                "Pricing_Modification": price_modification,
                "Pricing_HELOC": price_heloc,
                "Pricing_Reverse_Mortgage": price_reverse,
                "Pricing_All_Items": pricing_all_str,
                "Other_Pricing_Notes": pricing_notes,
                
                # Custom Fields
                "Custom_Fields": custom_fields_str,
                
                # State Licenses & Insurances
                "State_Licenses": licenses_str,
                "Insurances": insurances_str,
                
                # Background Checks
                "Background_Provider": bg_provider,
                "Background_Conducted": bg_conducted,
                "Background_Expiration": bg_expiration,
                "Background_Ref_Number": bg_ref,
                
                # Capabilities & Availability
                "Capabilities": capabilities_str,
                "Other_Capability_Info": other_caps,
                "Available_Days": days_str,
                "Available_Times": times_str,
                "Other_Availability_Notes": avail_notes,
                "Spoken_Languages": languages,
                
                # Websites & Social Media Links
                "Website_URL": website_url,
                "Facebook_URL": socials.get("facebook", ""),
                "LinkedIn_URL": socials.get("linkedin", ""),
                "Twitter_X_URL": socials.get("twitter", ""),
                "Instagram_URL": socials.get("instagram", ""),
                "Other_Social_Links": socials.get("other", ""),
                
                # Document Links
                "Document_Links": docs_str,
                
                # Metadata
                "City_Searched": summary_info.get("City_Searched", ""),
                "State_Searched": summary_info.get("State_Searched", ""),
                "Page_Number": summary_info.get("Page_Number", 1),
                "Position_on_Page": summary_info.get("Position_on_Page", 1),
                "Scraped_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            return detail_record
            
        except Exception as e:
            print(f"Error extracting profile {profile_url}: {e}")
            return None
        finally:
            await detail_page.close()

    async def run(self):
        print("\n" + "=" * 65)
        print("Starting NotaryCafe Scraper...")
        print(f"Target Excel 1 (City Stats):     {CITY_STATS_EXCEL}")
        print(f"Target Excel 2 (Summary Cards):  {SUMMARY_EXCEL}")
        print(f"Target Excel 3 (Full Profiles):  {DETAILED_EXCEL}")
        print("=" * 65)
        
        async with async_playwright() as p:
            print("Launching visible Chrome with anti-detection flags...")
            try:
                browser = await p.chromium.launch(
                    headless=False,
                    channel="chrome",
                    args=[
                        "--disable-blink-features=AutomationControlled",
                        "--start-maximized",
                        "--no-default-browser-check",
                        "--no-first-run"
                    ]
                )
            except Exception as e:
                print(f"Standard Chrome channel launch failed ({e}), falling back to default chromium...")
                browser = await p.chromium.launch(
                    headless=False,
                    args=[
                        "--disable-blink-features=AutomationControlled",
                        "--start-maximized"
                    ]
                )
                
            context = await browser.new_context(
                viewport=None,
                locale="en-US",
                user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36"
            )
            
            # Anti-detection stealth script
            await context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                window.chrome = {runtime: {}};
            """)
            
            search_page = await context.new_page()
            
            try:
                for c_idx, c_info in enumerate(self.cities, start=1):
                    city = c_info["city"]
                    state = c_info["state"]
                    city_key = f"{city}_{state}"
                    
                    if city_key in self.completed_cities:
                        print(f"[{c_idx}/{len(self.cities)}] City already completed: {city}, {state}. Skipping.")
                        continue
                        
                    print(f"\n>>> [{c_idx}/{len(self.cities)}] PROCESSING CITY: {city}, {state}")
                    
                    searched = await self.search_city(search_page, city, state)
                    if not searched:
                        print(f"Failed to submit search for {city}, {state}. Recording status...")
                        self.city_stats_records.append({
                            "City": city,
                            "State": state,
                            "Total_Listings_Found": 0,
                            "Total_Pages": 0,
                            "Status": "Search Failed",
                            "Scraped_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                        self.save_data()
                        continue
                        
                    current_page = 1
                    city_listings_count = 0
                    
                    # Multi-page pagination loop (scrapes ALL pages until no more pages)
                    while True:
                        print(f"\n--- [City: {city}, {state}] Processing Page {current_page} ---")
                        
                        listings = await self.parse_search_listings(search_page, city, state, current_page)
                        
                        if not listings:
                            if current_page == 1:
                                print(f"Zero listings found for {city}, {state}. Recording in City Stats...")
                                self.city_stats_records.append({
                                    "City": city,
                                    "State": state,
                                    "Total_Listings_Found": 0,
                                    "Total_Pages": 0,
                                    "Status": "No Results",
                                    "Scraped_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                                })
                                self.completed_cities.add(city_key)
                                self.save_data()
                            else:
                                print(f"No further listings found on page {current_page}. Pagination complete for {city}, {state}.")
                            break
                            
                        # Process listings on current page
                        print(f"Scraping {len(listings)} listings on Page {current_page}...")
                        for l_idx, summary_item in enumerate(listings, start=1):
                            prof_url = summary_item["Profile_URL"]
                            print(f"  [{l_idx}/{len(listings)}] {summary_item['Name']} ({prof_url})")
                            
                            self.summary_records.append(summary_item)
                            
                            if prof_url not in self.completed_profiles:
                                detail_record = await self.scrape_detail_page(context, prof_url, summary_item)
                                if detail_record:
                                    self.detailed_records.append(detail_record)
                                    self.completed_profiles.add(prof_url)
                            else:
                                print("    (Profile already in database, skipping deep crawl)")
                                
                            self.save_data()
                            await self.human_delay(1.0, 2.5)
                            
                        city_listings_count += len(listings)
                        
                        # Check user page limit if specified
                        if self.max_pages and current_page >= self.max_pages:
                            print(f"Reached max_pages limit ({self.max_pages}) for {city}, {state}.")
                            break
                            
                        # Detect and click NEXT page button
                        next_selectors = [
                            'a:has-text("Next")',
                            'button:has-text("Next")',
                            'a[rel="next"]',
                            '.pagination .next a',
                            '.pager .next a',
                            'li.next:not(.disabled) a',
                            'a[aria-label*="Next" i]',
                            f'.pagination a:has-text("{current_page + 1}")',
                            f'nav a:has-text("{current_page + 1}")',
                            f'a:has-text("{current_page + 1}")'
                        ]
                        
                        next_btn = None
                        for sel in next_selectors:
                            try:
                                candidate = await search_page.query_selector(sel)
                                if candidate and await candidate.is_visible():
                                    is_disabled = await candidate.evaluate(
                                        'el => el.classList.contains("disabled") || el.getAttribute("aria-disabled") === "true" || el.hasAttribute("disabled") || el.closest(".disabled") !== null'
                                    )
                                    if not is_disabled:
                                        next_btn = candidate
                                        break
                            except Exception:
                                pass
                                
                        if not next_btn:
                            print(f"No Next page button or next page link found. Finished all pages for {city}, {state}.")
                            break
                            
                        print(f"Navigating to Page {current_page + 1} for {city}, {state}...")
                        try:
                            prev_first_url = listings[0].get("Profile_URL") if listings else None
                            await next_btn.click()
                            try:
                                await search_page.wait_for_load_state("domcontentloaded", timeout=30000)
                            except Exception:
                                pass
                            await self.human_delay(3.0, 5.0)
                            
                            # Verify if new page loaded
                            new_cards = await search_page.query_selector_all('a:has-text("View Profile")')
                            if new_cards:
                                new_url = await new_cards[0].get_attribute("href")
                                if new_url and prev_first_url and new_url in prev_first_url:
                                    print("Waiting additional time for page change...")
                                    await self.human_delay(2.0, 4.0)
                                    
                            current_page += 1
                        except Exception as n_err:
                            print(f"Error navigating to next page: {n_err}")
                            break
                            
                    # City complete with results!
                    if city_listings_count > 0:
                        self.city_stats_records.append({
                            "City": city,
                            "State": state,
                            "Total_Listings_Found": city_listings_count,
                            "Total_Pages": current_page,
                            "Status": "Completed",
                            "Scraped_At": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        })
                        self.completed_cities.add(city_key)
                        self.save_data()
                        print(f"Successfully finished all listings for city: {city}, {state}! Total: {city_listings_count} listings across {current_page} pages.")
                        
                    if self.single_city:
                        print(f"Single city mode requested. Finished {city}, {state}.")
                        break
                        
                    await self.human_delay(3.0, 6.0)
                    
            finally:
                self.save_data()
                await browser.close()
                print("\nBrowser closed. Scraping session finished.")
                print(f"Total City Stats Records:      {len(self.city_stats_records)}")
                print(f"Total Summary Listings Saved:  {len(self.summary_records)}")
                print(f"Total Detailed Profiles Saved: {len(self.detailed_records)}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="NotaryCafe Automated Directory Scraper")
    parser.add_argument("--city", default=None, help="Starting city name (e.g. Bloomington).")
    parser.add_argument("--state", default=None, help="Starting state abbreviation (e.g. IN).")
    parser.add_argument("--fresh", action="store_true", help="Start completely fresh from Alphabet A to Z (backups old data).")
    parser.add_argument("--resume", action="store_true", help="Resume from last saved checkpoint without asking.")
    parser.add_argument("--sort-by", choices=["city", "state"], default="city", help="Sort order: 'city' (A to Z by city) or 'state' (State-wise A to Z).")
    parser.add_argument("--max-pages", type=int, default=None, help="Maximum pages to scrape per city.")
    parser.add_argument("--single-city", action="store_true", help="Scrape only the specified city then stop.")
    args = parser.parse_args()

    # Determine execution mode: command line flag OR interactive menu
    if args.fresh:
        mode = "fresh"
        sort_by = args.sort_by
    elif args.resume:
        mode = "resume"
        sort_by = args.sort_by
    elif sys.stdin.isatty():
        mode, sort_by = prompt_user_mode()
    else:
        mode = "resume"
        sort_by = args.sort_by

    if mode == "fresh":
        print("\n--- INITIATING FRESH START (ALPHABET A TO Z) ---")
        backup_existing_files()
    else:
        print("\n--- INITIATING RESUME MODE (CHECKING EXISTING FILES) ---")

    print(f"Scraper target: {args.city or 'Sequential Order'}, {args.state or 'Auto'} | Mode: {mode.upper()} | Sort: {sort_by.upper()} | Single city: {args.single_city} | Max pages: {args.max_pages}")
    cities_list = load_cities_from_excel(target_city=args.city, target_state=args.state, sort_by=sort_by)
    scraper = NotaryCafeScraper(
        cities_list, 
        start_city=args.city, 
        start_state=args.state, 
        max_pages=args.max_pages, 
        single_city=args.single_city
    )
    asyncio.run(scraper.run())
