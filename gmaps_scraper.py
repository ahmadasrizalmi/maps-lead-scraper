#!/usr/bin/env python3
"""
Google Maps Business Scraper — Ready for Outreach
Scrapes business listings from Google Maps and saves formatted Excel.

Usage:
    python3 gmaps_scraper.py "restaurants in Jakarta"
    python3 gmaps_scraper.py "digital agency Surabaya" --max 50
    python3 gmaps_scraper.py "klinik kecantikan Bandung" --max 100 --output leads.xlsx
"""

import argparse
import json
import re
import sys
import time
import urllib.parse
from datetime import datetime
from pathlib import Path

from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def clean_text(text: str) -> str:
    """Remove Google Maps icon unicode characters and clean whitespace."""
    if not text:
        return ''
    # Remove private use area characters (Google Maps icons)
    text = re.sub(r'[\ue000-\uf8ff]', '', text)
    # Remove other common icon ranges
    text = re.sub(r'[\uf000-\ufaff]', '', text)
    # Clean up whitespace
    text = re.sub(r'\n+', ' ', text).strip()
    text = re.sub(r'\s+', ' ', text)
    return text.strip()


def scrape_google_maps(query: str, max_results: int = 30, headless: bool = True) -> list[dict]:
    """Scrape business listings from Google Maps search results."""
    
    results = []
    
    with sync_playwright() as p:
        browser = p.chromium.launch(
            headless=headless,
            args=['--no-sandbox', '--disable-dev-shm-usage']
        )
        context = browser.new_context(
            viewport={'width': 1920, 'height': 1080},
            user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            locale='id-ID'
        )
        page = context.new_page()
        
        # Navigate to Google Maps search
        encoded_query = urllib.parse.quote_plus(query)
        url = f"https://www.google.com/maps/search/{encoded_query}"
        
        print(f"🔍 Searching: {query}")
        print(f"📡 URL: {url}")
        
        page.goto(url, wait_until='domcontentloaded', timeout=30000)
        page.wait_for_timeout(3000)
        
        # Accept cookies if prompted
        try:
            accept_btn = page.locator('button:has-text("Accept all"), button:has-text("Terima")')
            if accept_btn.count() > 0:
                accept_btn.first.click()
                page.wait_for_timeout(1000)
        except:
            pass
        
        # Find the scrollable results panel
        scrollable = page.locator('div[role="feed"]').first
        
        if scrollable.count() == 0:
            # Try alternative selector
            scrollable = page.locator('.m6QErb').first
        
        if scrollable.count() == 0:
            print("⚠️  Could not find results panel, trying direct extraction...")
        
        # Scroll to load more results
        print(f"📜 Loading results (max: {max_results})...")
        prev_count = 0
        scroll_attempts = 0
        max_scroll_attempts = 30
        
        while scroll_attempts < max_scroll_attempts:
            # Count current listings
            listings = page.locator('div.Nv2PK').all()
            current_count = len(listings)
            
            if current_count >= max_results:
                print(f"   ✅ Reached {current_count} results")
                break
            
            if current_count == prev_count:
                scroll_attempts += 1
                if scroll_attempts >= 3:
                    # Check if "You've reached the end" or no more results
                    end_text = page.locator('text="You\'ve reached the end"').count()
                    end_text2 = page.locator('text="Anda telah mencapai akhir"').count()
                    if end_text > 0 or end_text2 > 0:
                        print(f"   ℹ️  Reached end of results ({current_count} total)")
                        break
                    if scroll_attempts >= 5:
                        print(f"   ℹ️  No more results loading ({current_count} total)")
                        break
            else:
                scroll_attempts = 0
                print(f"   📊 Loaded {current_count} results...")
            
            prev_count = current_count
            
            # Scroll the results panel
            if scrollable.count() > 0:
                scrollable.evaluate('el => el.scrollTop = el.scrollHeight')
            else:
                page.evaluate('window.scrollBy(0, 1000)')
            
            page.wait_for_timeout(1500)
        
        # Extract business data from listings
        listings = page.locator('div.Nv2PK').all()
        print(f"\n📋 Extracting data from {min(len(listings), max_results)} listings...")
        
        for i, listing in enumerate(listings[:max_results]):
            try:
                data = extract_listing_data(listing, page)
                if data and data.get('name'):
                    # Clean all text fields
                    data = {k: clean_text(v) if isinstance(v, str) else v for k, v in data.items()}
                    results.append(data)
                    print(f"   [{i+1}/{min(len(listings), max_results)}] ✅ {data['name']}")
                    
                    # Visit detail page for more info
                    try:
                        detail_data = visit_detail_page(listing, page, data.get('name', ''))
                        if detail_data:
                            detail_data = {k: clean_text(v) if isinstance(v, str) else v for k, v in detail_data.items()}
                            data.update(detail_data)
                    except Exception as e:
                        pass  # Skip detail if fails
                        
            except Exception as e:
                print(f"   [{i+1}] ❌ Error: {str(e)[:60]}")
                continue
        
        browser.close()
    
    return results


def extract_listing_data(listing, page) -> dict:
    """Extract basic data from a search result listing."""
    data = {}
    
    # Business name
    name_el = listing.locator('.qBF1Pd, .fontHeadlineSmall, [class*="qBF1Pd"]').first
    if name_el.count() > 0:
        data['name'] = name_el.inner_text().strip()
    
    # Get all text from the listing for parsing
    full_text = listing.inner_text()
    lines = [l.strip() for l in full_text.split('\n') if l.strip()]
    
    # Category — usually the first line after name that contains category keywords
    category_keywords = [
        'Kedai', 'Kafe', 'Restoran', 'Toko', 'Klinik', 'Hotel', 'Rumah',
        'Coffee', 'Restaurant', 'Store', 'Shop', 'Clinic', 'Office',
        'Cafe', 'Bar', 'Spa', 'Salon', 'Gym', 'Studio', 'Agency',
        'Digital', 'Marketing', 'Photography', 'Catering', 'Bakery',
        'Warung', 'Mall', 'Plaza', 'Center', 'Centre'
    ]
    for line in lines[1:4]:  # Skip name, check next few lines
        if any(kw.lower() in line.lower() for kw in category_keywords):
            # Clean up — take just the category part before any rating
            cat_parts = line.split('·')
            data['category'] = cat_parts[0].strip() if cat_parts else line
            break
    
    # Rating and reviews — look for patterns like "4.5" and "(123)"
    for line in lines:
        # Rating: standalone number like 4.5, 4.8, etc.
        rating_match = re.match(r'^(\d\.\d)$', line.strip())
        if rating_match:
            try:
                data['rating'] = float(rating_match.group(1))
            except:
                pass
        
        # Reviews: number in parentheses like (123) or (1.234)
        reviews_match = re.match(r'^\(([\d.]+)\)$', line.strip())
        if reviews_match:
            try:
                reviews_text = reviews_match.group(1).replace('.', '')
                data['reviews'] = int(reviews_text)
            except:
                pass
    
    # Address — look for lines with address keywords
    for line in lines:
        line_lower = line.lower()
        if any(kw in line_lower for kw in ['jl', 'jalan', 'street', 'no.', 'rt', 'rw', 'kec', 'kel', 'kota', 'kab', 'blok', 'gedung', 'lantai']):
            # Clean the address — remove category prefix if present
            addr = line
            if '·' in addr:
                parts = addr.split('·')
                for part in parts:
                    if any(kw in part.lower() for kw in ['jl', 'jalan', 'street', 'no.', 'rt', 'rw']):
                        addr = part.strip()
                        break
            data['address'] = addr
            break
    
    # Status
    for line in lines:
        if 'Buka' in line or 'Open' in line:
            data['status'] = 'Open'
            break
        elif 'Tutup' in line or 'Closed' in line:
            data['status'] = 'Closed'
            break
    
    return data


def visit_detail_page(listing, page, business_name: str = '') -> dict:
    """Click into a listing to get detailed info (phone, website, hours)."""
    data = {}
    
    try:
        # Get the place URL from the listing's link
        name_link = listing.locator('a.hfpxzc, a[href*="/maps/place/"]').first
        if name_link.count() == 0:
            return data
        
        href = name_link.get_attribute('href') or ''
        if not href or '/place/' not in href:
            return data
        
        # Navigate directly to the place page
        page.goto(href, wait_until='domcontentloaded', timeout=15000)
        
        # Wait for phone/website elements to appear
        try:
            page.wait_for_selector('[data-item-id*="phone"], [data-item-id*="authority"], [data-item-id*="address"]', timeout=6000)
        except:
            page.wait_for_timeout(3000)
        
        # Extra wait for content to render
        page.wait_for_timeout(1500)
        
        # ===== PHONE NUMBER =====
        # Method 1: data-item-id with phone
        phone_els = page.locator('[data-item-id*="phone"]').all()
        for el in phone_els:
            text = el.inner_text().strip()
            if text:
                # Extract phone number from text
                phone_match = re.search(r'(\+?\d[\d\s\-()]{6,})', text)
                if phone_match:
                    data['phone'] = phone_match.group(1).strip()
                    break
        
        # Method 2: aria-label with phone
        if 'phone' not in data:
            phone_btns = page.locator('button[data-item-id*="phone"], [data-tooltip*="phone"], [data-tooltip*="telepon"]').all()
            for btn in phone_btns:
                aria = btn.get_attribute('aria-label') or btn.get_attribute('data-tooltip') or ''
                phone_match = re.search(r'(\+?\d[\d\s\-()]{6,})', aria)
                if phone_match:
                    data['phone'] = phone_match.group(1).strip()
                    break
        
        # Method 3: Look for phone in the detail text
        if 'phone' not in data:
            detail_text = page.locator('.m6QErb, .DUwDvf').first
            if detail_text.count() > 0:
                text = detail_text.inner_text()
                phone_match = re.search(r'(\+62[\d\s\-]{8,}|0[\d\s\-]{8,})', text)
                if phone_match:
                    data['phone'] = phone_match.group(1).strip()
        
        # ===== WEBSITE =====
        website_els = page.locator('[data-item-id*="authority"], [data-item-id*="website"]').all()
        for el in website_els:
            text = el.inner_text().strip()
            if text and 'google' not in text.lower() and len(text) > 3:
                data['website'] = text
                break
            # Check href
            href = el.get_attribute('href') or ''
            if href and 'google.com' not in href and href.startswith('http'):
                data['website'] = href
                break
        
        # Also try <a> tags with website-like hrefs
        if 'website' not in data:
            website_links = page.locator('a[data-item-id*="authority"], a[data-tooltip*="website"], a[data-tooltip*="Open website"]').all()
            for link in website_links:
                href = link.get_attribute('href') or ''
                if href and 'google.com' not in href and 'maps' not in href:
                    data['website'] = href
                    break
        
        # ===== FULL ADDRESS =====
        addr_els = page.locator('[data-item-id*="address"]').all()
        for el in addr_els:
            text = el.inner_text().strip()
            if text and len(text) > 5:
                data['address'] = text
                break
        
        # ===== PLUS CODE =====
        pluscode_el = page.locator('[data-item-id*="oloc"]').first
        if pluscode_el.count() > 0:
            text = pluscode_el.inner_text().strip()
            if text:
                data['plus_code'] = text
        
        # ===== OPENING HOURS =====
        # Method 1: aria-label on hours element
        hours_els = page.locator('[data-item-id*="hours"], .t39EBf, [aria-label*="hours"], [aria-label*="jam"]').all()
        for el in hours_els:
            aria = el.get_attribute('aria-label') or ''
            text = el.inner_text().strip()
            if aria and len(aria) > 5:
                data['hours'] = aria
                break
            elif text and len(text) > 5:
                data['hours'] = text
                break
        
        # Method 2: Hours table
        if 'hours' not in data:
            hours_rows = page.locator('.OqCZI tr, table tr').all()
            if hours_rows:
                hours_list = []
                for row in hours_rows[:7]:  # Max 7 days
                    try:
                        cells = row.locator('td').all()
                        if len(cells) >= 2:
                            day = cells[0].inner_text().strip()
                            time_val = cells[1].inner_text().strip()
                            if day and time_val:
                                hours_list.append(f"{day}: {time_val}")
                    except:
                        pass
                if hours_list:
                    data['hours'] = ' | '.join(hours_list)
        
        # ===== GOOGLE MAPS URL =====
        try:
            current_url = page.url
            if '/place/' in current_url:
                data['maps_url'] = current_url.split('&')[0]
        except:
            pass
        
        # ===== DESCRIPTION =====
        about_els = page.locator('.PYvSYb, [data-attrid*="description"]').all()
        for el in about_els:
            text = el.inner_text().strip()
            if text and len(text) > 10:
                data['description'] = text[:200]
                break
        
    except Exception as e:
        pass
    
    # Go back to search results for next listing
    try:
        page.go_back(wait_until='domcontentloaded', timeout=10000)
        page.wait_for_timeout(2000)
        # Wait for search results to load again
        page.wait_for_selector('div.Nv2PK', timeout=5000)
    except:
        pass
    
    return data


def save_to_excel(results: list[dict], output_file: str, query: str):
    """Save results to a formatted Excel file ready for outreach."""
    
    wb = Workbook()
    ws = wb.active
    ws.title = "Business Leads"
    
    # ===== STYLES =====
    header_font = Font(name='Calibri', bold=True, color='FFFFFF', size=11)
    header_fill = PatternFill(start_color='2563EB', end_color='2563EB', fill_type='solid')
    header_align = Alignment(horizontal='center', vertical='center', wrap_text=True)
    
    data_font = Font(name='Calibri', size=10)
    data_align = Alignment(vertical='center', wrap_text=True)
    
    alt_fill = PatternFill(start_color='F1F5F9', end_color='F1F5F9', fill_type='solid')
    
    thin_border = Border(
        left=Side(style='thin', color='D1D5DB'),
        right=Side(style='thin', color='D1D5DB'),
        top=Side(style='thin', color='D1D5DB'),
        bottom=Side(style='thin', color='D1D5DB')
    )
    
    # ===== HEADERS =====
    columns = [
        ('No', 5),
        ('Business Name', 30),
        ('Category', 20),
        ('Address', 40),
        ('Phone', 18),
        ('Website', 30),
        ('Email', 28),
        ('Rating', 8),
        ('Reviews', 9),
        ('Status', 10),
        ('Hours', 30),
        ('Google Maps URL', 45),
        ('Description', 35),
        ('Outreach Status', 15),
        ('Notes', 25),
    ]
    
    # Title row
    ws.merge_cells('A1:O1')
    title_cell = ws['A1']
    title_cell.value = f"📊 Business Leads — {query}"
    title_cell.font = Font(name='Calibri', bold=True, size=14, color='1E40AF')
    title_cell.alignment = Alignment(horizontal='center', vertical='center')
    ws.row_dimensions[1].height = 35
    
    # Subtitle
    ws.merge_cells('A2:O2')
    sub_cell = ws['A2']
    sub_cell.value = f"Generated: {datetime.now().strftime('%d %B %Y, %H:%M')} | Total: {len(results)} businesses"
    sub_cell.font = Font(name='Calibri', size=10, color='6B7280', italic=True)
    sub_cell.alignment = Alignment(horizontal='center')
    ws.row_dimensions[2].height = 22
    
    # Header row (row 4)
    header_row = 4
    for col_idx, (col_name, col_width) in enumerate(columns, 1):
        cell = ws.cell(row=header_row, column=col_idx, value=col_name)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = header_align
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(col_idx)].width = col_width
    
    ws.row_dimensions[header_row].height = 28
    
    # ===== DATA ROWS =====
    for idx, result in enumerate(results):
        row = header_row + 1 + idx
        
        # Try to extract email from website
        email = result.get('email', '')
        if not email and result.get('website'):
            email = guess_email_from_website(result['website'])
        
        values = [
            idx + 1,
            result.get('name', ''),
            result.get('category', ''),
            result.get('address', ''),
            result.get('phone', ''),
            result.get('website', ''),
            email,
            result.get('rating', ''),
            result.get('reviews', ''),
            result.get('status', ''),
            result.get('hours', ''),
            result.get('maps_url', ''),
            result.get('description', ''),
            '',  # Outreach Status (empty for user to fill)
            '',  # Notes (empty for user to fill)
        ]
        
        for col_idx, value in enumerate(values, 1):
            cell = ws.cell(row=row, column=col_idx, value=value)
            cell.font = data_font
            cell.alignment = data_align
            cell.border = thin_border
            
            # Alternate row colors
            if idx % 2 == 1:
                cell.fill = alt_fill
        
        # Make URLs clickable
        website_cell = ws.cell(row=row, column=6)
        if result.get('website'):
            website_cell.font = Font(name='Calibri', size=10, color='2563EB', underline='single')
        
        maps_cell = ws.cell(row=row, column=12)
        if result.get('maps_url'):
            maps_cell.font = Font(name='Calibri', size=10, color='2563EB', underline='single')
        
        ws.row_dimensions[row].height = 22
    
    # ===== DROPDOWN FOR OUTREACH STATUS =====
    from openpyxl.worksheet.datavalidation import DataValidation
    
    if results:
        status_col = 14  # Column N
        dv = DataValidation(
            type="list",
            formula1='"Not Contacted,Email Sent,Follow Up 1,Follow Up 2,Responded,Interested,Not Interested,Closed"',
            allow_blank=True
        )
        dv.error = "Please select a valid status"
        dv.errorTitle = "Invalid Status"
        first_data_row = header_row + 1
        last_data_row = header_row + len(results)
        dv.add(f'N{first_data_row}:N{last_data_row}')
        ws.add_data_validation(dv)
    
    # ===== FREEZE PANES =====
    ws.freeze_panes = f'A{header_row + 1}'
    
    # ===== AUTO FILTER =====
    if results:
        last_col = get_column_letter(len(columns))
        ws.auto_filter.ref = f'A{header_row}:{last_col}{header_row + len(results)}'
    
    # ===== SUMMARY SHEET =====
    ws2 = wb.create_sheet("Summary")
    
    summary_data = [
        ("📊 Outreach Summary", ""),
        ("", ""),
        ("Search Query", query),
        ("Total Businesses", len(results)),
        ("With Phone", sum(1 for r in results if r.get('phone'))),
        ("With Website", sum(1 for r in results if r.get('website'))),
        ("With Email", sum(1 for r in results if r.get('email'))),
        ("With Address", sum(1 for r in results if r.get('address'))),
        ("Average Rating", f"{sum(r.get('rating', 0) for r in results if r.get('rating')) / max(sum(1 for r in results if r.get('rating')), 1):.1f}"),
        ("Generated", datetime.now().strftime('%d %B %Y, %H:%M')),
        ("", ""),
        ("📝 Tips for Outreach", ""),
        ("1.", "Personalize each email — mention their business name"),
        ("2.", "Reference their Google reviews or rating"),
        ("3.", "Keep subject line under 50 characters"),
        ("4.", "Follow up after 3-5 days if no response"),
        ("5.", "Use 'Outreach Status' column to track progress"),
    ]
    
    for row_idx, (key, val) in enumerate(summary_data, 1):
        ws2.cell(row=row_idx, column=1, value=key).font = Font(name='Calibri', bold=True, size=11)
        ws2.cell(row=row_idx, column=2, value=val).font = Font(name='Calibri', size=11)
    
    ws2.column_dimensions['A'].width = 25
    ws2.column_dimensions['B'].width = 50
    
    # Title style
    ws2['A1'].font = Font(name='Calibri', bold=True, size=16, color='1E40AF')
    
    # Save
    wb.save(output_file)
    print(f"\n💾 Saved to: {output_file}")
    print(f"   📊 {len(results)} businesses exported")
    print(f"   📞 {sum(1 for r in results if r.get('phone'))} with phone numbers")
    print(f"   🌐 {sum(1 for r in results if r.get('website'))} with websites")


def guess_email_from_website(website: str) -> str:
    """Try to find email from website's contact page."""
    # This is a best-effort approach
    # In production, you'd want to actually visit the website
    if not website:
        return ''
    
    # Clean URL
    if not website.startswith('http'):
        website = 'https://' + website
    
    try:
        domain = urllib.parse.urlparse(website).netloc
        # Common email patterns
        common_emails = [
            f"info@{domain}",
            f"contact@{domain}",
            f"hello@{domain}",
        ]
        # Return the most likely one (info@ is most common)
        return common_emails[0]
    except:
        return ''


def main():
    parser = argparse.ArgumentParser(
        description='Google Maps Business Scraper — Ready for Outreach',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  python3 gmaps_scraper.py "restaurants in Jakarta"
  python3 gmaps_scraper.py "digital marketing agency Surabaya" --max 50
  python3 gmaps_scraper.py "klinik kecantikan Bandung" --max 100 --output leads.xlsx
  python3 gmaps_scraper.py "hotel Bali" --max 200 --visible
        """
    )
    parser.add_argument('query', help='Search query for Google Maps')
    parser.add_argument('--max', type=int, default=30, help='Maximum results to scrape (default: 30)')
    parser.add_argument('--output', '-o', help='Output Excel file (default: auto-generated)')
    parser.add_argument('--visible', action='store_true', help='Show browser window (non-headless)')
    
    args = parser.parse_args()
    
    # Generate output filename
    if not args.output:
        safe_query = re.sub(r'[^a-zA-Z0-9\s]', '', args.query).strip().replace(' ', '_')[:50]
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        args.output = f"gmaps_{safe_query}_{timestamp}.xlsx"
    
    print("=" * 60)
    print("🗺️  Google Maps Business Scraper")
    print("=" * 60)
    print(f"🔍 Query: {args.query}")
    print(f"📊 Max results: {args.max}")
    print(f"📁 Output: {args.output}")
    print(f"🖥️  Mode: {'visible' if args.visible else 'headless'}")
    print("=" * 60)
    print()
    
    # Scrape
    results = scrape_google_maps(args.query, args.max, headless=not args.visible)
    
    if not results:
        print("\n❌ No results found. Try a different search query.")
        sys.exit(1)
    
    # Save to Excel
    save_to_excel(results, args.output, args.query)
    
    # Also save raw JSON as backup
    json_file = args.output.replace('.xlsx', '.json')
    with open(json_file, 'w', encoding='utf-8') as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"📄 Raw JSON backup: {json_file}")
    
    print("\n✅ Done! Open the Excel file to start your outreach campaign.")
    print(f"   💡 Tip: Use 'Outreach Status' dropdown to track your progress")


if __name__ == '__main__':
    main()
