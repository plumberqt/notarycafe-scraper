"""
NotaryCafe Directory Scraper - Adaptive Version
Handles variable listings per page, unknown page counts, empty results, adaptive scraping.
"""

import asyncio
import json
import random
import re
import traceback
from pathlib import Path
from dataclasses import dataclass, asdict, field
from typing import List, Dict, Optional, Set, Any
from playwright.async_api import async_playwright
import pandas as pd
from datetime import datetime


@dataclass
class ListingDetail:
    listing_name: str = ""
    listing_url: str = ""
    city_searched: str = ""
    state_searched: str = ""
    page_number: int = 0
    position_on_page: int = 0
    
    name: str = ""
    email: str = ""
    phone: str = ""
    mobile: str = ""
    fax: str = ""
    website: str = ""
    
    street_address: str = ""
    city: str = ""
    state: str = ""
    zip_code: str = ""
    county: str = ""
    
    commission_number: str = ""
    commission_expiration: str = ""
    bonding_company: str = ""
    errors_omissions_insurance: str = ""
    notary_type: str = ""
    services_offered: str = ""
    languages_spoken: str = ""
    business_hours: str = ""
    fees: str = ""
    coverage_area: str = ""
    years_in_business: str = ""
    background_check: str = ""
    
    description: str = ""
    fees: str = ""
    coverage_area: str = ""
    years_in_business: str = ""
    background_check: str = ""
    
    scraped_at: str = ""
    scraped_from_url: str = ""
    source_page_number: int = 0
    position_on_page: int = 0
    
    raw_html: str = ""
    
    def to_dict(self) -> dict:
        return asdict(self)


@dataclass
class CityProgress:
    state: str
    city: str
    current_page: int = 1
    total_pages: int = 0
    listings_scraped: int = 0
    total_listings: int = 0
    last_listing_index: int = 0
    completed: bool = False
    error_count: int = 0
    last_error: str = ""
    last_update: str = ""
    has_data: bool = False


class NotaryCafeScraper:
    def __init__(self, config: Dict):
        self.config = config
        self.base_url = "https://www.notarycafe.com"
        self.search_url = "https://www.notarycafe.com/search"
        
        self.cities_file = config.get('cities_file', r'D:\Karan\cities-basic-v2025_structured.xlsx')
        self.chrome_user_data = config.get('chrome_user_data', r'C:\Users\nishu\AppData\Local\Google\Chrome\User Data')
        self.profile_dir = config.get('profile_dir', 'Profile 3')
        
        self.output_dir = Path(config.get('output_dir', r'D:\Karan\notarycafe_output'))
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_file = self.output_dir / f'notarycafe_listings_{self.session_id}.xlsx'
        self.progress_file = self.output_dir / f'scraping_progress_{self.session_id}.json'
        self.error_log = self.output_dir / f'error_log_{self.session_id}.jsonl'
        
        self.all_listings: List[Dict] = []
        self.city_progress: Dict[str, Dict] = {}
        self.completed_cities: Set[str] = set()
        self.total_listings_scraped = 0
        self.start_time = datetime.now()
        
        self.delay_range = config.get('delay_range', (4, 8))
        self.page_delay_range = config.get('page_delay_range', (3, 6))
        self.max_retries = config.get('max_retries', 3)
        self.batch_size = config.get('batch_size', 10)
        
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        self.cities_file = config.get('cities_file', r'D:\Karan\cities-basic-v2025_structured.xlsx')
        self.chrome_user_data = config.get('chrome_user_data', r'C:\Users\nishu\AppData\Local\Google\Chrome\User Data')
        self.profile_dir = config.get('profile_dir', 'Profile 3')
        
        self.output_dir = Path(config.get('output_dir', r'D:\Karan\notarycafe_output'))
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        self.output_file = self.output_dir / f'notarycafe_listings_{self.session_id}.xlsx'
        self.progress_file = self.output_dir / f'scraping_progress_{self.session_id}.json'
        self.error_log = self.output_dir / f'error_log_{self.session_id}.jsonl'
        
        self.all_listings: List[Dict] = []
        self.city_progress: Dict[str, Dict] = {}
        self.completed_cities: Set[str] = set()
        self.total_listings_scraped = 0
        self.start_time = datetime.now()
        
        self.delay_range = config.get('delay_range', (4, 8))
        self.page_delay_range = config.get('page_delay_range', (3, 6))
        self.max_retries = config.get('max_retries', 3)
        self.batch_size = config.get('batch_size', 10)
        
        self.session_id = datetime.now().strftime("%Y%m%d_%H%M%S")
        
    def load_cities(self) -> List[Dict]:
        print("Loading cities from structured file...")
        df = pd.read_excel(self.cities_file, sheet_name='All_Cities_Raw')
        
        cities = []
        for _, row in df.iterrows():
            cities.append({
                'state': row['State'],
                'state_abbr': row['State_Abbr'] if 'State_Abbr' in row else '',
                'city': row['City'],
                'geo_id': row['GeoID'] if 'GeoID' in row else None
            })
        
        seen = set()
        unique_cities = []
        for c in cities:
            key = (c['state'], c['city'])
            if key not in seen:
                seen.add(key)
                unique_cities.append(c)
        
        print(f"Loaded {len(unique_cities)} unique cities")
        return unique_cities
    
    def load_progress(self):
        if self.progress_file.exists():
            try:
                with open(self.progress_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                
                self.completed_cities = set(data.get('completed_cities', []))
                self.total_listings_scraped = data.get('total_listings_scraped', 0)
                
                for key, val in data.get('city_progress', {}).items():
                    self.city_progress[key] = val
                
                print(f"Resumed: {len(self.completed_cities)} completed, {self.total_listings_scraped} listings")
            except Exception as e:
                print(f"Could not load progress: {e}")
    
    def save_progress(self):
        try:
            data = {
                'session_id': self.session_id,
                'completed_cities': list(self.completed_cities),
                'total_listings_scraped': self.total_listings_scraped,
                'last_update': datetime.now().isoformat(),
                'city_progress': self.city_progress
            }
            with open(self.progress_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Warning: Could not save progress: {e}")
    
    def save_listings_incremental(self):
        if not self.all_listings:
            return
        
        try:
            df = pd.DataFrame(self.all_listings)
            output_path = self.output_dir / f'notarycafe_listings_{self.session_id}.xlsx'
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                pd.DataFrame(self.all_listings).to_excel(writer, sheet_name='All_Listings', index=False)
                
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
            
            csv_path = self.output_dir / f'notarycafe_listings_{self.session_id}.csv'
            pd.DataFrame(self.all_listings).to_csv(csv_path, index=False, encoding='utf-8')
            
            print(f"  ✅ Saved {len(self.all_listings)} listings")
        except Exception as e:
            print(f"  ❌ Error saving: {e}")
    
    def log_error(self, error_type: str, message: str, context: Dict = None):
        try:
            error_entry = {
                'timestamp': datetime.now().isoformat(),
                'error_type': error_type,
                'message': message,
                'context': context or {}
            }
            with open(self.error_log, 'a', encoding='utf-8') as f:
                f.write(json.dumps(error_entry, ensure_ascii=False) + '\n')
        except:
            pass
    
    async def wait_random(self, delay_range: tuple = None):
        dr = delay_range or self.delay_range
        delay = random.uniform(*dr)
        await asyncio.sleep(delay)
    
    async def human_like_scroll(self, page):
        try:
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
        except:
            pass
    
    async def find_listing_elements(self, page) -> List:
        """Adaptively find listing elements on the page"""
        strategies = [
            ('table tbody tr', 'table rows'),
            ('div[class*="listing"]', 'listing divs'),
            ('div[class*="result"]', 'result divs'),
            ('div[class*="card"]', 'card divs'),
            ('a[href*="/listing/"]', 'listing links'),
            ('a[href*="/profile/"]', 'profile links'),
            ('tr[class*="row"]', 'row classes'),
            ('li[class*="item"]', 'list items'),
        ]
        
        for selector, description in strategies:
            try:
                elements = await page.query_selector_all(selector)
                if elements and len(elements) > 0:
                    # Filter out header rows and non-listing elements
                    valid_elements = []
                    for el in elements:
                        try:
                            text = await el.inner_text()
                            if len(text.strip()) > 10:  # Has meaningful content
                                valid_elements.append(el)
                        except:
                            continue
                    
                    if valid_elements:
                        return valid_elements
            except:
                continue
        
        return []
    
    async def extract_listing_summary(self, element, city: str, state: str, page_num: int, idx: int) -> Optional[Dict]:
        try:
            data = {
                'city_searched': city,
                'state_searched': state,
                'page_number': page_num,
                'position_on_page': idx,
                'listing_name': '',
                'listing_url': '',
                'name': '',
                'email': '',
                'phone': '',
                'city': '',
                'state': '',
                'zip': '',
                'address': ''
            }
            
            # Get link
            link = await element.query_selector('a[href]')
            if link:
                href = await link.get_attribute('href')
                if href:
                    data['listing_url'] = href if href.startswith('http') else f"https://www.notarycafe.com{href}"
                    text = await link.inner_text()
                    data['listing_name'] = text.strip()
            
            # Extract text
            try:
                full_text = await element.inner_text()
                data['raw_text'] = full_text[:500]
                
                lines = [l.strip() for l in full_text.split('\n') if l.strip()]
                if lines:
                    data['listing_name'] = lines[0]
                
                # Phone
                phone_match = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', full_text)
                if phone_match:
                    data['phone'] = phone_match.group()
                
                # Email
                email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', full_text)
                if email_match:
                    data['email'] = email_match.group()
                
                # Address
                addr_match = re.search(r'([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5})', full_text)
                if addr_match:
                    data['city'] = addr_match.group(1).strip()
                    data['state'] = addr_match.group(2)
                    data['zip'] = addr_match.group(3)
            except:
                pass
            
            return data if (data.get('listing_name') or data.get('listing_url')) else None
        except Exception as e:
            return None
    
    async def extract_listings_from_page(self, page, city: str, state: str, page_num: int) -> List[Dict]:
        listings = []
        try:
            elements = await self.find_listing_elements(page)
            
            if not elements:
                return []
            
            for idx, element in enumerate(elements):
                try:
                    data = await self.extract_listing_summary(element, city, state, page_num, idx)
                    if data:
                        data['city_searched'] = city
                        data['state_searched'] = state
                        data['page_number'] = page_num
                        data['position_on_page'] = idx
                        listings.append(data)
                except Exception as e:
                    continue
        except Exception as e:
            print(f"  Error extracting: {e}")
        return listings
    
    async def detect_pagination(self, page) -> Dict:
        """Detect pagination info adaptively"""
        result = {
            'has_pagination': False,
            'total_pages': 1,
            'current_page': 1,
            'next_page_selector': None
        }
        
        try:
            # Look for pagination indicators
            pagination_selectors = [
                'nav.pagination',
                '.pagination',
                '.pager',
                '[class*="pagination"]',
                '[class*="pager"]',
                'nav[aria-label="pagination"]'
            ]
            
            for selector in pagination_selectors:
                container = await page.query_selector(selector)
                if container:
                    result['has_pagination'] = True
                    break
            
            # Find page numbers
            page_numbers = []
            page_links = await page.query_selector_all('a[href*="page"], a[href*="p="], a[href*="pg="]')
            for link in page_links:
                text = await link.inner_text()
                if text.isdigit():
                    page_numbers.append(int(text))
            
            if page_numbers:
                result['total_pages'] = max(page_numbers)
            
            # Find next page selector
            next_selectors = [
                'a:has-text("Next")',
                'a:has-text("next")',
                'a[rel="next"]',
                '.pagination .next a',
                '.pager .next a',
                'a[aria-label="Next"]'
            ]
            
            for selector in next_selectors:
                btn = await page.query_selector(selector)
                if btn and await btn.is_visible():
                    result['next_page_selector'] = selector
                    break
            
            # Current page
            try:
                active = await page.query_selector('.pagination .active, .pager .active, .active[aria-current="page"]')
                if active:
                    text = await active.inner_text()
                    if text.isdigit():
                        result['current_page'] = int(text)
            except:
                pass
                
        except Exception as e:
            pass
        
        return result
    
    async def click_next_page(self, page, pagination_info: Dict) -> bool:
        if not pagination_info.get('has_pagination'):
            return False
        
        next_selector = pagination_info.get('next_page_selector')
        if next_selector:
            try:
                btn = await page.query_selector(next_selector)
                if btn and await btn.is_visible() and await btn.is_enabled():
                    await btn.click()
                    await page.wait_for_load_state('domcontentloaded', timeout=30000)
                    await asyncio.sleep(random.uniform(2, 4))
                    return True
            except:
                pass
        
        # Try clicking page number
        try:
            next_page = pagination_info.get('current_page', 1) + 1
            page_link = await page.query_selector(f'a:has-text("{next_page}")')
            if page_link:
                await page_link.click()
                await page.wait_for_load_state('domcontentloaded', timeout=30000)
                await asyncio.sleep(random.uniform(2, 4))
                return True
        except:
            pass
        
        return False
    
    async def wait_for_listings(self, page, timeout=30000) -> bool:
        """Wait for listings to appear"""
        try:
            await page.wait_for_load_state('domcontentloaded', timeout=timeout)
            await asyncio.sleep(1)
            
            # Check for various listing indicators
            indicators = [
                'table tbody tr',
                'div[class*="listing"]',
                'div[class*="result"]',
                'a[href*="/listing/"]',
                'a[href*="/profile/"]',
                'div[class*="card"]'
            ]
            
            for selector in indicators:
                try:
                    await page.wait_for_selector(selector, timeout=5000)
                    elements = await page.query_selector_all(selector)
                    if elements and len(elements) > 0:
                        return True
                except:
                    continue
            
            return False
        except:
            return False
    
    async def search_city(self, page, city: str, state: str) -> bool:
        try:
            await page.goto(self.search_url, wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(random.uniform(2, 4))
            
            # Find search input
            search_selectors = [
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
            
            search_input = None
            for selector in search_selectors:
                try:
                    inp = await page.query_selector(selector)
                    if inp and await inp.is_visible():
                        search_input = inp
                        break
                except:
                    continue
            
            if not search_input:
                inputs = await page.query_selector_all('input[type="text"], input[type="search"]')
                for inp in inputs:
                    if await inp.is_visible():
                        search_input = inp
                        break
            
            if not search_input:
                return False
            
            query = f"{city}, {state}"
            await search_input.click()
            await search_input.fill("")
            await asyncio.sleep(0.3)
            await search_input.type(query, delay=random.randint(50, 150))
            await asyncio.sleep(0.5)
            await search_input.press('Enter')
            await page.wait_for_load_state('domcontentloaded', timeout=60000)
            await asyncio.sleep(random.uniform(2, 4))
            
            return await self.wait_for_listings(page)
            
        except Exception as e:
            return False
    
    async def detect_no_results(self, page) -> bool:
        """Check if page indicates no results"""
        try:
            text = await page.inner_text('body')
            no_result_indicators = [
                'no results', 'no listings', 'no matches',
                '0 results', '0 listings', 'nothing found',
                'no notaries found', 'no notaries in'
            ]
            text_lower = text.lower()
            for indicator in no_result_indicators:
                if indicator in text_lower:
                    return True
        except:
            pass
        return False
    
    async def extract_detail_fields(self, page, detail: Dict):
        try:
            full_text = await page.inner_text('body')
            
            # Name from h1
            try:
                h1 = await page.query_selector('h1')
                if h1:
                    listing_name = (await h1.inner_text()).strip()
                    if listing_name:
                        listing_name = listing_name.strip()
            except:
                pass
            
            full_text = await page.inner_text('body')
            
            # Email
            email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', full_text)
            if email_match:
                detail['email'] = email_match.group()
            
            # Phone
            phone_patterns = [
                r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}',
                r'\d{3}-\d{3}-\d{4}',
                r'\(\d{3}\)\s*\d{3}-\d{4}',
            ]
            for pattern in phone_patterns:
                matches = re.findall(pattern, full_text)
                if matches:
                    phone = matches[0]
                    if not phone:
                        continue
                    break
            
            # Address
            addr_patterns = [
                r'(\d+\s+[A-Za-z0-9\s]+(?:Street|St|Avenue|Ave|Road|Rd|Drive|Dr|Lane|Ln|Boulevard|Blvd|Court|Ct|Circle|Cir|Way|Place|Pl)\.?),\s*([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5}(?:-\d{4})?)',
                r'(\d+\s+[A-Za-z0-9\s]+),\s*([A-Za-z\s]+),\s*([A-Z]{2})\s*(\d{5})',
            ]
            
            for pattern in addr_patterns:
                match = re.search(pattern, full_text)
                if match:
                    break
            
            # Commission
            commission_patterns = [
                r'[Cc]ommission\s*(?:number|#|no\.?)\s*[:\-]?\s*([A-Z0-9\-]+)',
                r'[Cc]ommission\s*[Ee]xpir(?:ation|y)\s*[:\-]?\s*(\d{1,2}/\d{1,2}/\d{4})',
            ]
            for pattern in commission_patterns:
                match = re.search(pattern, full_text, re.IGNORECASE)
                if match:
                    break
            
            # Services
            services_match = re.search(r'(?:[Ss]ervices|[Oo]ffered?)\s*[:\-]\s*([^\n]{20,200})', full_text)
            if services_match:
                detail['services_offered'] = services_match.group(1).strip()
            
            # Languages
            lang_match = re.search(r'[Ll]anguages?\s*[:\-]\s*([^\n]{5,100})', full_text)
            if lang_match:
                detail['languages_spoken'] = lang_match.group(1).strip()
            
            # Hours
            hours_match = re.search(r'(?:[Hh]ours?|[Oo]pen|[Aa]vailable)\s*[:\-]\s*([^\n]{5,100})', full_text)
            if hours_match:
                detail['business_hours'] = hours_match.group(1).strip()
            
            # Fees
            fees_match = re.search(r'(?:[Ff]ees?|[Cc]ost|[Pp]rice|[Rr]ate)\s*[:\-]\s*([^\n]{5,100})', full_text)
            if fees_match:
                detail['fees'] = fees_match.group(1).strip()
            
            # Coverage
            coverage_match = re.search(r'(?:[Cc]overage|[Aa]rea|[Ss]erves?)\s*[:\-]\s*([^\n]{10,200})', full_text)
            if coverage_match:
                detail['coverage_area'] = coverage_match.group(1).strip()
            
        except Exception as e:
            print(f"    Error extracting detail fields: {e}")
    
    async def extract_from_html(self, html: str, detail: Dict):
        try:
            email_match = re.search(r'mailto:([\w\.-]+@[\w\.-]+\.\w+)', html)
            if email_match and not detail.get('email'):
                detail['email'] = email_match.group(1)
            
            tel_match = re.search(r'tel:([\d\-\(\)\s\+]+)', html)
            if tel_match and not detail.get('phone'):
                detail['phone'] = tel_match.group(1)
            
            json_ld_matches = re.findall(r'<script type="application/ld\+json">(.*?)</script>', html, re.DOTALL)
            for match in json_ld_matches:
                try:
                    data = json.loads(match.strip())
                    if isinstance(data, dict):
                        if 'email' in data and not detail.get('email'):
                            detail['email'] = data['email']
                        if 'telephone' in data and not detail.get('phone'):
                            detail['phone'] = data['telephone']
                        if 'address' in data:
                            addr = data['address']
                            if isinstance(addr, dict):
                                if 'streetAddress' in addr: detail['street_address'] = addr['streetAddress']
                                if 'addressLocality' in addr: detail['city'] = addr['addressLocality']
                                if 'addressRegion' in addr: detail['state'] = addr['addressRegion']
                                if 'postalCode' in addr: detail['zip_code'] = addr['postalCode']
                except:
                    pass
        except:
            pass
    
    async def scrape_detail_page(self, page, listing_url: str, city: str, state: str, page_num: int, position: int) -> Dict:
        detail = {
            'listing_name': '',
            'listing_url': listing_url,
            'city_searched': city,
            'state_searched': state,
            'page_number': page_num,
            'position_on_page': position,
            'name': '',
            'email': '',
            'phone': '',
            'mobile': '',
            'fax': '',
            'website': '',
            'street_address': '',
            'city': '',
            'state': '',
            'zip_code': '',
            'county': '',
            'commission_number': '',
            'commission_expiration': '',
            'bonding_company': '',
            'errors_omissions_insurance': '',
            'notary_type': '',
            'services_offered': '',
            'languages_spoken': '',
            'business_hours': '',
            'fees': '',
            'coverage_area': '',
            'years_in_business': '',
            'background_check': '',
            'description': '',
            'fees': '',
            'coverage_area': '',
            'years_in_business': '',
            'background_check': '',
            'scraped_at': datetime.now().isoformat(),
            'scraped_from_url': '',
            'source_page_number': page_num,
            'position_on_page': position,
            'raw_html': ''
        }
        
        try:
            await page.goto(listing_url, wait_until='domcontentloaded', timeout=60000)
            await asyncio.sleep(random.uniform(2, 4))
            
            detail['scraped_from_url'] = page.url
            detail['raw_html'] = await page.content()
            
            await self.human_like_scroll(page)
            
            await self.extract_detail_fields(page, detail)
            await self.extract_from_html(await page.content(), detail)
            
        except Exception as e:
            detail['description'] = f"ERROR: {str(e)}"
        
        return detail
    
    async def scrape_city(self, page, city_info: Dict, progress: Dict) -> int:
        city = city_info['city']
        state = city_info['state']
        
        print(f"\n📍 Scraping {city}, {state}...")
        
        success = await self.search_city(page, city, state)
        if not success:
            progress['error_count'] = progress.get('error_count', 0) + 1
            progress['last_error'] = "Search failed"
            return 0
        
        # Check for no results
        if await self.detect_no_results(page):
            print(f"  ℹ️ No results for {city}, {state}")
            return 0
        
        pagination = await self.detect_pagination(page)
        total_pages = pagination.get('total_pages', 1)
        print(f"  📄 Detected pages: {total_pages}")
        
        current_page = max(1, progress.get('current_page', 1))
        
        if current_page > 1:
            print(f"  ⏭️ Resuming from page {current_page}...")
            for _ in range(current_page - 1):
                if not await self.click_next_page(page, {'current_page': 1, 'has_pagination': True}):
                    break
                await asyncio.sleep(random.uniform(2, 4))
        
        city_listings_count = 0
        current_page = max(1, progress.get('current_page', 1))
        
        while True:
            print(f"  📄 Page {current_page}...")
            
            # Check for no results on this page
            if await self.detect_no_results(page):
                print(f"  ℹ️ No results on page {current_page}")
                break
            
            listings = await self.extract_listings_from_page(page, city, state, 1)
            
            if not listings:
                print(f"  ⚠️ No listings on page {current_page}")
                # Try next page
                if not await self.click_next_page(page, {'current_page': current_page, 'has_pagination': True}):
                    break
                await asyncio.sleep(random.uniform(3, 6))
                continue
            
            print(f"  📋 Found {len(listings)} listings")
            
            for idx, listing_summary in enumerate(listings):
                listing_index = (1) * 20 + idx  # approximate
                
                if listing_index < progress.get('last_listing_index', 0):
                    continue
                
                listing_url = listing_summary.get('listing_url', '')
                if not listing_url:
                    continue
                
                print(f"    📋 [{idx+1}/{len(listings)}] {listing_summary.get('listing_name', 'Unknown')[:50]}...")
                
                detail = await self.scrape_detail_page(
                    page, listing_url, city, state, 1, idx
                )
                detail['city_searched'] = city
                detail['state_searched'] = state
                detail['page_number'] = 1
                detail['position_on_page'] = idx
                
                self.all_listings.append(detail)
                self.total_listings_scraped += 1
                city_listings_count += 1
                
                if self.total_listings_scraped % self.batch_size == 0:
                    self.save_listings_incremental()
                    self.save_progress()
                
                await asyncio.sleep(random.uniform(2, 5))
            
            # Try next page
            if not await self.click_next_page(page, {'current_page': 1, 'has_pagination': True}):
                print(f"  🛑 No more pages or next button failed")
                break
            
            await asyncio.sleep(random.uniform(3, 6))
        
        return city_listings_count
    
    async def run(self):
        print("=" * 60)
        print("NotaryCafe Directory Scraper - Adaptive")
        print("=" * 60)
        
        cities = self.load_cities()
        self.load_progress()
        
        remaining = [c for c in cities if f"{c['state']}_{c['city']}" not in self.completed_cities]
        print(f"Total: {len(cities)}, Done: {len(self.completed_cities)}, Remaining: {len(remaining)}")
        
        async with async_playwright() as p:
            context = await p.chromium.launch_persistent_context(
                user_data_dir=r'C:\Users\nishu\AppData\Local\Google\Chrome\User Data',
                channel="chrome",
                headless=False,
                args=[
                    "--disable-blink-features=AutomationControlled",
                    "--no-first-run",
                    "--no-default-browser-check",
                ],
                viewport={"width": 1920, "height": 1080},
                locale="en-US",
                timezone_id="America/New_York"
            )
            
            await context.add_init_script("""
                Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
                window.chrome = {runtime: {}};
            """)
            
            page = context.pages[0] if context.pages else await context.new_page()
            
            try:
                for i, city_info in enumerate(remaining_cities, 1):
                    key = f"{city_info['state']}_{city_info['city']}"
                    
                    if key not in self.city_progress:
                        self.city_progress[key] = {
                            'state': city_info['state'],
                            'city': city_info['city'],
                            'current_page': 1,
                            'total_pages': 0,
                            'listings_scraped': 0,
                            'total_listings': 0,
                            'last_listing_index': 0,
                            'completed': False,
                            'error_count': 0,
                            'last_error': '',
                            'last_update': '',
                            'has_data': False
                        }
                    
                    progress = self.city_progress[key]
                    
                    print(f"\n{'='*60}")
                    print(f"[{i}/{len(remaining)}] {city_info['city']}, {city_info['state']}")
                    
                    try:
                        count = await self.scrape_city(page, city_info, progress)
                        progress['completed'] = True
                        progress['has_data'] = True
                        self.completed_cities.add(f"{city_info['state']}_{city_info['city']}")
                        self.save_progress()
                        self.save_listings_incremental()
                        print(f"  ✅ Done: {count} listings")
                    except Exception as e:
                        print(f"  ❌ Error: {e}")
                        progress['error_count'] = progress.get('error_count', 0) + 1
                        progress['last_error'] = str(e)
                        self.save_progress()
                    
                    if i < len(remaining_cities):
                        await asyncio.sleep(random.uniform(5, 15))
                
                self.save_listings_incremental()
                self.save_progress()
                print(f"\n✅ Complete! Total: {self.total_listings_scraped}")
                
            finally:
                await context.close()
                self.save_listings_incremental()
                self.save_progress()


async def main():
    config = {
        'cities_file': r'D:\Karan\cities-basic-v2025_structured.xlsx',
        'chrome_user_data': r'C:\Users\nishu\AppData\Local\Google\Chrome\User Data',
        'profile_dir': 'Profile 3',
        'output_dir': r'D:\Karan\notarycafe_output',
        'delay_range': (4, 8),
        'page_delay_range': (3, 6),
        'max_retries': 3,
        'batch_size': 10
    }
    
    scraper = NotaryCafeScraper(config)
    await scraper.run()


if __name__ == "__main__":
    asyncio.run(main())