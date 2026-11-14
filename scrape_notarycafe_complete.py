"""
NotaryCafe Complete Automated Scraper
====================================
Scrapes notary listings and detailed profiles from https://notarycafe.com/
Generates two structured Excel files:
  1. notarycafe_listings_summary.xlsx  (10 listings per page summary + social links)
  2. notarycafe_detailed_profiles.xlsx (Full deep profile data for each notary)

Supports:
  - Starting from a specific city (e.g. Bloomington, IN) or reading the full list from cities-basic-v2025_structured.xlsx
  - Automatic pagination through all 10-result pages per city
  - Opening each profile in a new tab, extracting all fields, closing tab, returning to results
  - Extracting social media links (Facebook, LinkedIn, Twitter/X, Instagram, etc.)
  - Auto-resuming from progress.json
  - Visible Chrome browser with anti-detection flags for VPN / Cloudflare compatibility
"""

import asyncio
import json
import os
import re
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

def load_cities_from_excel(target_city=None, target_state=None):
    """Load cities from Excel file. If target specified, starts from that city."""
    print(f"Loading cities from: {CITIES_EXCEL}")
    df = pd.read_excel(CITIES_EXCEL, sheet_name="All_Cities_Raw")
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
            
    print(f"Loaded {len(cities)} unique cities from file.")
    
    if target_city and target_state:
        # Move target city to the front
        target_idx = None
        for i, c in enumerate(cities):
            if c["city"].lower() == target_city.lower() and c["state"].upper() == target_state.upper():
                target_idx = i
                break
        if target_idx is not None:
            match = cities.pop(target_idx)
            cities.insert(0, match)
            print(f"Set starting city to: {match['city']}, {match['state']}")
        else:
            cities.insert(0, {"city": target_city, "state": target_state, "population": 0})
            print(f"Added starting city: {target_city}, {target_state}")
            
    return cities


class NotaryCafeScraper:
    def __init__(self, cities, start_city="Bloomington", start_state="IN", max_pages=None, single_city=False):
        self.cities = cities
        self.start_city = start_city
        self.start_state = start_state
        self.max_pages = max_pages
        self.single_city = single_city
        self.completed_cities = set()
        self.completed_profiles = set()
        self.summary_records = []
        self.detailed_records = []
        self.load_existing_data()

    def load_existing_data(self):
        """Loads existing Excel files and progress.json if present to prevent duplicates."""
        if PROGRESS_FILE.exists():
            try:
                with open(PROGRESS_FILE, "r", encoding="utf-8") as f:
                    p = json.load(f)
                    self.completed_cities = set(p.get("completed_cities", []))
                    self.completed_profiles = set(p.get("completed_profiles", []))
                    print(f"Resumed from progress.json: {len(self.completed_cities)} completed cities, {len(self.completed_profiles)} profiles.")
            except Exception as e:
                print(f"Warning reading progress.json: {e}")

        if SUMMARY_EXCEL.exists():
            try:
                df_sum = pd.read_excel(SUMMARY_EXCEL)
                self.summary_records = df_sum.to_dict("records")
                print(f"Loaded {len(self.summary_records)} existing summary records.")
            except Exception as e:
                print(f"Warning reading existing summary excel: {e}")

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

    def save_data(self):
        """Saves current state safely to both Excel files and progress.json."""
        # 1. Save progress.json first (lightweight, crash-safe)
        try:
            progress = {
                "completed_cities": list(self.completed_cities),
                "completed_profiles": list(self.completed_profiles),
                "summary_count": len(self.summary_records),
                "detailed_count": len(self.detailed_records),
                "last_updated": datetime.now().isoformat()
            }
            with open(PROGRESS_FILE, "w", encoding="utf-8") as f:
                json.dump(progress, f, indent=2)
        except Exception as pe:
            print(f"Warning saving progress.json: {pe}")

        # 2. Save Summary Excel (handle file lock if user has file open)
        try:
            if self.summary_records:
                df_s = pd.DataFrame(self.summary_records)
                df_s.drop_duplicates(subset=["Profile_URL", "City_Searched"], keep="last", inplace=True)
                df_s.to_excel(SUMMARY_EXCEL, index=False)
        except PermissionError:
            print(f"Notice: {SUMMARY_EXCEL.name} is currently open in Excel. Skipping write this turn (data safe in memory).")
        except Exception as se:
            print(f"Warning saving summary excel: {se}")

        # 3. Save Detailed Profiles Excel (cast object types to avoid LossySetitemError)
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
        """Parses the 10 listing cards visible on the search results page."""
        listings = []
        
        # Check for no results
        body_text = await page.inner_text("body")
        no_res_indicators = ["no results", "0 results", "no notaries found", "no matches found"]
        if any(ind in body_text.lower() for ind in no_res_indicators):
            print(f"No notaries found for {city}, {state}.")
            return []
            
        # Find all 'View Profile' links
        view_profile_links = await page.query_selector_all('a:has-text("View Profile")')
        if not view_profile_links:
            # Fallback check for any links starting with / that lead to member profiles
            links = await page.query_selector_all('a[href]')
            view_profile_links = []
            for l in links:
                txt = (await l.inner_text()).strip()
                if "view profile" in txt.lower():
                    view_profile_links.append(l)
                    
        print(f"Found {len(view_profile_links)} listings on Page {page_num}.")
        
        for idx, link_el in enumerate(view_profile_links, start=1):
            try:
                href = await link_el.get_attribute("href")
                if not href or href.strip() in ["#", "", "/"] or href.endswith("#"):
                    continue
                # Ignore system routes
                clean_href = href.strip().lower()
                if any(clean_href == bad or clean_href.startswith(bad) for bad in ["/register", "/plan", "/about", "/contact", "/terms", "/privacy", "/forum", "http://facebook.com", "http://twitter.com"]):
                    continue
                profile_url = href if href.startswith("http") else f"{BASE_URL}{href}"
                
                # Find parent card/container of this link
                card_text = ""
                card_el = await link_el.evaluate_handle('el => el.closest("div, li, tr") ? el.closest("div, li, tr").innerText : ""')
                card_text = str(card_el) if card_el else ""
                
                # In case evaluation didn't get full container, walk up 2-3 levels
                card_container = await link_el.evaluate_handle('''el => {
                    let p = el.parentElement;
                    while (p && p.innerText && p.innerText.length < 50 && p.parentElement) {
                        p = p.parentElement;
                    }
                    return p ? p.innerText : "";
                }''')
                card_full_text = str(card_container) if card_container else card_text
                
                # Extract fields from card
                name = ""
                company = ""
                c_city = ""
                c_state = ""
                c_zip = ""
                c_phone = ""
                c_email = ""
                
                # Name & Company from lines
                lines = [l.strip() for l in card_full_text.split("\n") if l.strip() and "miles" not in l.lower() and "profile completion" not in l.lower() and "view profile" not in l.lower()]
                if lines:
                    name = lines[0]
                    if len(lines) > 1 and not any(ch.isdigit() for ch in lines[1]):
                        company = lines[1]
                        
                # Phone regex
                phone_match = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', card_full_text)
                if phone_match:
                    c_phone = phone_match.group(0)
                    
                # Email regex
                email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', card_full_text)
                if email_match:
                    c_email = email_match.group(0)
                    
                # Location regex (City State Zip)
                loc_match = re.search(r'([A-Za-z\s]+)\s+([A-Z]{2})\s+(\d{5})', card_full_text)
                if loc_match:
                    c_city = loc_match.group(1).strip()
                    c_state = loc_match.group(2).strip()
                    c_zip = loc_match.group(3).strip()
                else:
                    c_city = city
                    c_state = state

                # Extract social links from card
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
                    # Ignore generic share links
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
            
            # Get text and html
            page_text = await detail_page.inner_text("body")
            page_html = await detail_page.content()
            
            # --- EXTRACT ALL REQUIRED FIELDS ---
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
            if m_comp: company_name = m_comp.group(1).strip()
            
            in_business_since = ""
            m_since = re.search(r'In Business Since:\s*([^\n]+)', page_text, re.IGNORECASE)
            if m_since: in_business_since = m_since.group(1).strip()
            
            # 5. Billing Address
            b_preferred = "Yes" if "Billing" in page_text and "Preferred" in page_text else ""
            b_addr1 = ""
            b_addr2 = ""
            b_city = ""
            b_state = ""
            b_zip = ""
            
            m_b_addr1 = re.search(r'Address 1:\s*([^\n]+)', page_text)
            if m_b_addr1: b_addr1 = m_b_addr1.group(1).strip()
            
            m_b_addr2 = re.search(r'Address 2:\s*([^\n]+)', page_text)
            if m_b_addr2: b_addr2 = m_b_addr2.group(1).strip()
            
            m_b_city = re.search(r'City:\s*([^\n]+)', page_text)
            if m_b_city: b_city = m_b_city.group(1).strip()
            
            m_b_state = re.search(r'State:\s*([A-Za-z]{2})', page_text)
            if m_b_state: b_state = m_b_state.group(1).strip()
            
            m_b_zip = re.search(r'Zip:\s*(\d{5}(?:-\d{4})?)', page_text)
            if m_b_zip: b_zip = m_b_zip.group(1).strip()
            
            # 6. Shipping Address
            s_same = "Yes" if "Same as Billing" in page_text else "No"
            
            # 7. Phone Numbers (Work, Mobile, Home)
            p_work = summary_info.get("Phone", "")
            p_mobile = ""
            p_home = ""
            p_pref = ""
            p_text_enabled = "Yes" if "text enabled" in page_text.lower() or "text" in page_text.lower() else ""
            
            m_p_work = re.search(r'Work:\s*(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', page_text, re.IGNORECASE)
            if m_p_work: p_work = m_p_work.group(1).strip()
            
            m_p_mob = re.search(r'Mobile:\s*(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', page_text, re.IGNORECASE)
            if m_p_mob: p_mobile = m_p_mob.group(1).strip()
            
            m_p_home = re.search(r'Home:\s*(\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4})', page_text, re.IGNORECASE)
            if m_p_home: p_home = m_p_home.group(1).strip()
            
            if not p_work and not p_mobile:
                all_phones = re.findall(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', page_text)
                if all_phones: p_work = all_phones[0]
                
            # 8. Email Addresses (Work, Personal)
            e_work = summary_info.get("Email", "")
            e_personal = ""
            
            m_e_work = re.search(r'Work\s+([\w\.-]+@[\w\.-]+\.\w+)', page_text)
            if m_e_work: e_work = m_e_work.group(1).strip()
            
            m_e_pers = re.search(r'Personal\s+([\w\.-]+@[\w\.-]+\.\w+)', page_text)
            if m_e_pers: e_personal = m_e_pers.group(1).strip()
            
            if not e_work:
                all_emails = re.findall(r'[\w\.-]+@[\w\.-]+\.\w+', page_text)
                # Filter out system emails
                clean_emails = [e for e in all_emails if "notarycafe" not in e.lower() and "bixagency" not in e.lower()]
                if clean_emails: e_work = clean_emails[0]
                
            # 9. Pricing Information
            pricing_items = []
            pricing_regex = re.findall(r'(\d+\.\s*[^\n\$]+)\s*(\$\d+(?:\.\d{2})?)', page_text)
            for desc, amt in pricing_regex:
                pricing_items.append(f"{desc.strip()}: {amt.strip()}")
            pricing_all_str = " | ".join(pricing_items)
            
            pricing_notes = ""
            m_pnotes = re.search(r'Other Pricing Notes\s*\n+([^\n]+)', page_text, re.IGNORECASE)
            if m_pnotes: pricing_notes = m_pnotes.group(1).strip()
            
            # Specific common pricing slots
            price_purchase = ""
            price_seller = ""
            price_modification = ""
            price_heloc = ""
            price_reverse = ""
            for item in pricing_items:
                i_low = item.lower()
                if "purchase" in i_low or "refi" in i_low: price_purchase = item
                elif "seller" in i_low: price_seller = item
                elif "modification" in i_low: price_modification = item
                elif "heloc" in i_low or "equity" in i_low: price_heloc = item
                elif "reverse" in i_low: price_reverse = item

            # 10. Custom Fields
            custom_fields = []
            m_cfields = re.findall(r'(?:Custom Field \d+:|Custom:\s*)([^\n]+)', page_text, re.IGNORECASE)
            for cf in m_cfields:
                custom_fields.append(cf.strip())
            custom_fields_str = " | ".join(custom_fields)
            
            # 11. State Licenses
            licenses = []
            lic_matches = re.findall(r'([A-Z]{2})\s+([A-Za-z0-9#\-]+)\s+(\d{1,2}/\d{1,2}/\d{4})', page_text)
            for st, comm, exp in lic_matches:
                licenses.append(f"{st} #{comm} (Exp: {exp})")
            licenses_str = " | ".join(licenses) if licenses else ""
            if not licenses_str:
                m_lic_text = re.findall(r'(?:License|Commission)\s*#?([A-Za-z0-9\-]+)', page_text, re.IGNORECASE)
                if m_lic_text: licenses_str = ", ".join(m_lic_text)
                
            # 12. Insurances
            insurances = []
            m_ins = re.findall(r'([A-Za-z0-9\.\s]+(?:USD|\$|Bonding|Insurance)[^\n]*)', page_text)
            for ins in m_ins:
                if len(ins.strip()) < 80 and any(w in ins.lower() for w in ["usd", "$", "merchants", "biberk", "insurance"]):
                    insurances.append(ins.strip())
            insurances_str = " | ".join(insurances)
            
            # 13. Background Checks
            bg_provider = ""
            bg_conducted = ""
            bg_expiration = ""
            bg_ref = ""
            
            m_bg_p = re.search(r'Provider\s*\n+([A-Za-z\s]+)', page_text)
            if m_bg_p: bg_provider = m_bg_p.group(1).strip()
            
            m_bg_c = re.search(r'Conducted\s*\n+(\d{1,2}/\d{1,2}/\d{4})', page_text)
            if m_bg_c: bg_conducted = m_bg_c.group(1).strip()
            
            m_bg_e = re.search(r'Expiration\s*\n+(\d{1,2}/\d{1,2}/\d{4})', page_text)
            if m_bg_e: bg_expiration = m_bg_e.group(1).strip()
            
            m_bg_r = re.search(r'Ref\.\s*#\s*\n+([A-Za-z0-9\-]+)', page_text)
            if m_bg_r: bg_ref = m_bg_r.group(1).strip()
            
            # 14. Capabilities
            capabilities_list = [
                "CSA", "Attorney", "Fax", "Email", "Internet", "Laser Printer", 
                "Abstractor", "Mobile Hotspot", "E-Sign", "24 Hour Service", 
                "Fingerprinting", "Weddings", "Hospital Signing", "Jail Signings", 
                "EscrowTab (IPEN)", "RON Capable"
            ]
            active_caps = [c for c in capabilities_list if c.lower() in page_text.lower()]
            capabilities_str = ", ".join(active_caps)
            
            other_caps = ""
            m_ocaps = re.search(r'Other Capability Information\s*\n+([^\n]+)', page_text, re.IGNORECASE)
            if m_ocaps: other_caps = m_ocaps.group(1).strip()
            
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
            
            # Build final full profile record
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
        print("\n" + "=" * 60)
        print("Starting NotaryCafe Scraper...")
        print(f"Target Excel 1: {SUMMARY_EXCEL}")
        print(f"Target Excel 2: {DETAILED_EXCEL}")
        print("=" * 60)
        
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
                    
                    # 1. Search for city
                    searched = await self.search_city(search_page, city, state)
                    if not searched:
                        print(f"Failed to submit search for {city}, {state}. Continuing.")
                        continue
                        
                    current_page = 1
                    while True:
                        print(f"\n--- [City: {city}, {state}] Processing Page {current_page} ---")
                        
                        # 2. Extract listings on this page (10 listings)
                        listings = await self.parse_search_listings(search_page, city, state, current_page)
                        if not listings:
                            print(f"No more listings on page {current_page} for {city}, {state}.")
                            break
                            
                        # 3. For each listing, extract summary & full profile
                        for l_idx, summary_item in enumerate(listings, start=1):
                            prof_url = summary_item["Profile_URL"]
                            print(f"  [{l_idx}/{len(listings)}] {summary_item['Name']} ({prof_url})")
                            
                            # Add to summary records
                            self.summary_records.append(summary_item)
                            
                            # Scrape full detail profile in new tab
                            if prof_url not in self.completed_profiles:
                                detail_record = await self.scrape_detail_page(context, prof_url, summary_item)
                                if detail_record:
                                    self.detailed_records.append(detail_record)
                                    self.completed_profiles.add(prof_url)
                            else:
                                print(f"    (Profile already in database, skipping deep crawl)")
                                
                            # Save after each listing
                            self.save_data()
                            await self.human_delay(1.0, 2.5)
                            
                        # 4. Check for Next Page button
                        # Detect total pages indicator: e.g. "Page 1 of 7 • 70 results"
                        body_txt = await search_page.inner_text("body")
                        page_match = re.search(r'Page\s+(\d+)\s+of\s+(\d+)', body_txt, re.IGNORECASE)
                        total_pages = int(page_match.group(2)) if page_match else 1
                        
                        print(f"City {city}: Page {current_page} of {total_pages} complete.")
                        
                        if current_page >= total_pages or (self.max_pages and current_page >= self.max_pages):
                            print(f"Reached page limit ({current_page}/{total_pages}) for {city}, {state}.")
                            break
                            
                        # Click Next button
                        next_btn = await search_page.query_selector('a:has-text("Next"), button:has-text("Next"), .pagination .next a')
                        if not next_btn:
                            # Try generic search for Next
                            all_a = await search_page.query_selector_all('a')
                            for a in all_a:
                                txt = (await a.inner_text()).strip().lower()
                                if txt == "next":
                                    next_btn = a
                                    break
                                    
                        if next_btn and await next_btn.is_visible():
                            print("Clicking 'Next' page button...")
                            await next_btn.click()
                            await search_page.wait_for_load_state("domcontentloaded", timeout=45000)
                            await self.human_delay(3.0, 5.0)
                            current_page += 1
                        else:
                            print("Could not find visible 'Next' button. Moving to next city.")
                            break
                            
                    # City finished!
                    self.completed_cities.add(city_key)
                    self.save_data()
                    print(f"Successfully finished all listings for city: {city}, {state}!")
                    
                    if self.single_city:
                        print(f"Single city mode requested. Finished {city}, {state}.")
                        break
                        
                    await self.human_delay(3.0, 6.0)
                    
            finally:
                self.save_data()
                await browser.close()
                print("\nBrowser closed. Scraping session finished.")
                print(f"Total Summary Listings Saved: {len(self.summary_records)}")
                print(f"Total Detailed Profiles Saved: {len(self.detailed_records)}")


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser(description="NotaryCafe Scraper")
    parser.add_argument("--city", default="Bloomington", help="Starting city name (default: Bloomington)")
    parser.add_argument("--state", default="IN", help="Starting state abbreviation (default: IN)")
    parser.add_argument("--max-pages", type=int, default=None, help="Maximum pages to scrape per city (e.g. 1 for quick test)")
    parser.add_argument("--single-city", action="store_true", help="Scrape only the specified city then stop")
    args = parser.parse_args()

    print(f"Scraper target: {args.city}, {args.state} | Single city: {args.single_city} | Max pages: {args.max_pages}")
    cities_list = load_cities_from_excel(target_city=args.city, target_state=args.state)
    scraper = NotaryCafeScraper(
        cities_list, 
        start_city=args.city, 
        start_state=args.state, 
        max_pages=args.max_pages, 
        single_city=args.single_city
    )
    asyncio.run(scraper.run())
