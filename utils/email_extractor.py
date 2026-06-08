"""
Email & Social Media Extractor
Visits business websites to extract emails and social media links.
"""

import re
import urllib.parse
from typing import Optional


def extract_emails_from_html(html: str, domain: str = '') -> list[str]:
    """Extract email addresses from HTML content."""
    emails = set()
    
    # Pattern 1: Standard email regex
    email_pattern = re.compile(r'[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}')
    found = email_pattern.findall(html)
    for email in found:
        email = email.lower().strip('.')
        # Filter out common false positives
        if not any(x in email for x in [
            'example.com', 'test.com', 'domain.com', 'email.com',
            'sentry.io', 'w3.org', 'schema.org', 'googleapis.com',
            'google.com', 'facebook.com', 'twitter.com', 'wordpress.org',
            'gravatar.com', 'wp.com', 'cloudflare.com', 'jquery.com',
            'bootstrap.com', 'fontawesome', '.png', '.jpg', '.gif',
            '.svg', '.webp', '.css', '.js',
        ]):
            emails.add(email)
    
    # Pattern 2: Obfuscated emails (at) and [dot]
    obfuscated = re.compile(r'[a-zA-Z0-9._%+\-]+\s*[\[\(]\s*at\s*[\]\)]\s*[a-zA-Z0-9.\-]+\s*[\[\(]\s*dot\s*[\]\)]\s*[a-zA-Z]{2,}', re.IGNORECASE)
    for match in obfuscated.findall(html):
        cleaned = match.replace('(at)', '@').replace('[at]', '@').replace('(dot)', '.').replace('[dot]', '.')
        cleaned = re.sub(r'\s+', '', cleaned)
        emails.add(cleaned.lower())
    
    # Pattern 3: mailto: links
    mailto_pattern = re.compile(r'mailto:([a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,})')
    for match in mailto_pattern.findall(html):
        emails.add(match.lower())
    
    # Prioritize emails with the same domain
    email_list = list(emails)
    if domain:
        domain_lower = domain.lower().replace('www.', '')
        # Sort: same domain first, then generic (info@, contact@), then others
        def priority(e):
            e_domain = e.split('@')[1] if '@' in e else ''
            if e_domain == domain_lower:
                return 0
            if e.startswith(('info@', 'contact@', 'hello@', 'support@', 'sales@')):
                return 1
            return 2
        email_list.sort(key=priority)
    
    return email_list


def extract_social_links(html: str, base_url: str = '') -> dict:
    """Extract social media links from HTML content."""
    socials = {}
    
    patterns = {
        'instagram': [
            re.compile(r'https?://(?:www\.)?instagram\.com/[a-zA-Z0-9_.]+/?'),
            re.compile(r'https?://(?:www\.)?instagr\.am/[a-zA-Z0-9_.]+/?'),
        ],
        'facebook': [
            re.compile(r'https?://(?:www\.)?facebook\.com/[a-zA-Z0-9_.]+/?'),
            re.compile(r'https?://(?:www\.)?fb\.com/[a-zA-Z0-9_.]+/?'),
            re.compile(r'https?://(?:www\.)?facebook\.com/pages/[a-zA-Z0-9_.\-]+/\d+/?'),
        ],
        'tiktok': [
            re.compile(r'https?://(?:www\.)?tiktok\.com/@[a-zA-Z0-9_.]+/?'),
        ],
        'twitter': [
            re.compile(r'https?://(?:www\.)?twitter\.com/[a-zA-Z0-9_]+/?'),
            re.compile(r'https?://(?:www\.)?x\.com/[a-zA-Z0-9_]+/?'),
        ],
        'youtube': [
            re.compile(r'https?://(?:www\.)?youtube\.com/(?:c/|channel/|@)[a-zA-Z0-9_\-]+/?'),
            re.compile(r'https?://(?:www\.)?youtube\.com/user/[a-zA-Z0-9_\-]+/?'),
        ],
        'linkedin': [
            re.compile(r'https?://(?:www\.)?linkedin\.com/(?:company|in)/[a-zA-Z0-9_\-]+/?'),
        ],
        'whatsapp': [
            re.compile(r'https?://(?:wa\.me|api\.whatsapp\.com/send\?phone=)\d+'),
            re.compile(r'https?://(?:www\.)?whatsapp\.com/channel/[a-zA-Z0-9]+'),
        ],
        'telegram': [
            re.compile(r'https?://(?:t\.me|telegram\.me)/[a-zA-Z0-9_]+/?'),
        ],
    }
    
    for platform, platform_patterns in patterns.items():
        for pattern in platform_patterns:
            matches = pattern.findall(html)
            if matches:
                # Clean and deduplicate
                cleaned = list(set(m.rstrip('/') for m in matches))
                socials[platform] = cleaned[0] if len(cleaned) == 1 else cleaned
                break
    
    return socials


def extract_phones_from_html(html: str) -> list[str]:
    """Extract additional phone numbers from HTML content."""
    phones = set()
    
    # Indonesian phone patterns
    patterns = [
        re.compile(r'\+62[\s\-]?\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}'),
        re.compile(r'0\d{2,4}[\s\-]?\d{3,4}[\s\-]?\d{3,4}'),
        re.compile(r'\+62[\s\-]?\d{9,13}'),
    ]
    
    for pattern in patterns:
        for match in pattern.findall(html):
            cleaned = re.sub(r'[\s\-]', '', match)
            if len(cleaned) >= 10:
                phones.add(match.strip())
    
    return list(phones)


def get_website_info(url: str, timeout: int = 10) -> dict:
    """Visit a website and extract email, social media, and additional phones."""
    import requests
    
    if not url:
        return {}
    
    # Ensure URL has protocol
    if not url.startswith('http'):
        url = 'https://' + url
    
    info = {}
    
    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8',
            'Accept-Language': 'id-ID,id;q=0.9,en;q=0.8',
        }
        
        response = requests.get(url, headers=headers, timeout=timeout, allow_redirects=True)
        html = response.text
        
        # Get domain for email prioritization
        domain = urllib.parse.urlparse(response.url).netloc
        
        # Extract emails
        emails = extract_emails_from_html(html, domain)
        if emails:
            info['emails'] = emails
            info['email'] = emails[0]  # Best guess
        
        # Extract social media
        socials = extract_social_links(html, response.url)
        if socials:
            info['social_media'] = socials
            # Flatten for easy access
            for platform, link in socials.items():
                info[f'social_{platform}'] = link if isinstance(link, str) else link[0]
        
        # Extract additional phones
        phones = extract_phones_from_html(html)
        if phones:
            info['website_phones'] = phones
        
        # Also check /contact, /about, /hubungi pages
        for path in ['/contact', '/about', '/hubungi', '/tentang-kami', '/kontak']:
            try:
                contact_url = urllib.parse.urljoin(response.url, path)
                contact_resp = requests.get(contact_url, headers=headers, timeout=5)
                if contact_resp.status_code == 200:
                    contact_html = contact_resp.text
                    
                    # More emails from contact page
                    contact_emails = extract_emails_from_html(contact_html, domain)
                    for email in contact_emails:
                        if email not in emails:
                            emails.append(email)
                    
                    # More social links
                    contact_socials = extract_social_links(contact_html, contact_url)
                    for platform, link in contact_socials.items():
                        if platform not in socials:
                            socials[platform] = link
                            info[f'social_{platform}'] = link if isinstance(link, str) else link[0]
                    
                    # More phones
                    contact_phones = extract_phones_from_html(contact_html)
                    for phone in contact_phones:
                        if phone not in phones:
                            phones.append(phone)
                    
                    if emails:
                        info['emails'] = emails
                        info['email'] = emails[0]
                    if socials:
                        info['social_media'] = socials
                    if phones:
                        info['website_phones'] = phones
            except:
                pass
        
    except Exception as e:
        info['error'] = str(e)[:100]
    
    return info
