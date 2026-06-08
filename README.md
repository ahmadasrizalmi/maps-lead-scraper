# 🧲 Maps Lead Scraper

Scrape business leads from Google Maps → formatted Excel for outreach campaigns.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-1.40+-2EAD33?logo=playwright&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.28+-FF4B4B?logo=streamlit&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

## ✨ Features

### Core Scraping
- 🔍 **Auto-scrape** from Google Maps (no API key needed)
- 📞 **Phone numbers**, 🌐 **websites**, 📍 **full addresses**
- ⭐ **Ratings & reviews**, 🕐 **opening hours**
- 📝 **Business descriptions**

### Email & Social Media Extraction
- 📧 **Auto-extract emails** from business websites
- 📱 **Instagram**, **Facebook**, **TikTok**, **Twitter/X** links
- 💬 **WhatsApp**, **Telegram** links
- 🔗 **LinkedIn** company pages
- 🔍 Scrapes `/contact`, `/about`, `/hubungi` pages too

### Export & Integration
- 📊 **Formatted Excel** with headers, filters, dropdowns
- 📋 **Outreach Status** tracking (Not Contacted → Closed)
- 📄 **JSON backup** for data processing
- 📈 **Google Sheets** integration (coming soon)

### Proxy Support
- 🔄 **Proxy rotation** for large-scale scraping
- 📁 Load proxies from file or comma-separated string
- 🎲 Automatic rotation with failure detection

## 📸 Screenshots

### GUI Mode (Streamlit)
Modern dark theme with real-time progress tracking.

### Excel Output
Professional formatted spreadsheet ready for outreach.

## 🚀 Quick Start

### 1. Install Dependencies

```bash
pip install -r requirements.txt
playwright install chromium
```

### 2. Run GUI Mode

```bash
streamlit run gmaps_gui.py
```

Opens at `http://localhost:8501`

### 3. Run CLI Mode

```bash
# Basic usage
python3 gmaps_scraper.py "klinik kecantikan Jakarta"

# With email extraction
python3 gmaps_scraper.py "digital marketing agency Surabaya" --max 50 --extract-emails

# With proxy rotation
python3 gmaps_scraper.py "restaurant Bali" --max 100 --proxy-file proxies.txt

# With inline proxies
python3 gmaps_scraper.py "hotel Bandung" --max 50 --proxies "http://proxy1:8080,http://proxy2:8080"
```

## 📊 Excel Output Format

| Column | Description |
|--------|-------------|
| No | Row number |
| Business Name | Company/business name |
| Category | Business category |
| Address | Full address |
| Phone | Phone number |
| Website | Business website |
| **Email** | Extracted from website |
| **Instagram** | Instagram profile link |
| **Facebook** | Facebook page link |
| **TikTok** | TikTok profile link |
| Rating | Google rating (1-5) |
| Reviews | Number of reviews |
| Status | Open/Closed |
| Hours | Opening hours |
| Google Maps URL | Direct link |
| Description | Business description |
| **Outreach Status** | Dropdown: Not Contacted → Closed |
| **Notes** | Your notes |

## 📧 Email Extraction

The `--extract-emails` flag visits each business website and:

1. Scans the homepage HTML for email patterns
2. Checks `/contact`, `/about`, `/hubungi` pages
3. Extracts from `mailto:` links
4. Handles obfuscated emails `(at)` and `[dot]`
5. Prioritizes emails matching the business domain
6. Filters out false positives (sentry, w3.org, etc.)

### Social Media Detection

Automatically detects links to:
- Instagram (instagram.com, instagr.am)
- Facebook (facebook.com, fb.com)
- TikTok (tiktok.com/@username)
- Twitter/X (twitter.com, x.com)
- YouTube (youtube.com/c/, /channel/, /@)
- LinkedIn (linkedin.com/company/, /in/)
- WhatsApp (wa.me, api.whatsapp.com)
- Telegram (t.me, telegram.me)

## 🔄 Proxy Rotation

For large-scale scraping, use proxies to avoid rate limits:

```bash
# From file (one proxy per line)
python3 gmaps_scraper.py "query" --proxy-file proxies.txt

# Inline (comma-separated)
python3 gmaps_scraper.py "query" --proxies "http://proxy1:8080,socks5://proxy2:1080"
```

### Proxy File Format

```
# proxies.txt
http://proxy1.example.com:8080
http://user:pass@proxy2.example.com:3128
socks5://proxy3.example.com:1080
```

## 🎯 Use Cases

- 📧 **Email outreach** campaigns
- 📞 **Cold calling** lead lists
- 🏪 **Market research** by location
- 📊 **Competitor analysis**
- 🎯 **Sales prospecting**
- 📱 **Social media marketing** (find Instagram/FB accounts)

## 📁 Project Structure

```
maps-lead-scraper/
├── gmaps_scraper.py          # CLI version
├── gmaps_gui.py              # Streamlit GUI version
├── requirements.txt          # Python dependencies
├── README.md                 # This file
└── utils/
    ├── __init__.py
    ├── email_extractor.py    # Email & social media extraction
    ├── proxy_rotator.py      # Proxy rotation support
    └── sheets_exporter.py    # Google Sheets integration (WIP)
```

## ⚠️ Disclaimer

This tool is for educational purposes. Use responsibly and respect:
- Google Maps Terms of Service
- Privacy laws (GDPR, etc.)
- Anti-spam regulations

## 📝 License

MIT License

## 🤝 Contributing

Pull requests welcome! Ideas for improvement:
- [x] Email extraction from websites
- [x] Social media links scraping
- [x] Proxy rotation support
- [ ] Google Sheets integration (in progress)
- [ ] Multi-location batch scraping
- [ ] Export to CRM formats (HubSpot, Salesforce)
- [ ] Email validation (MX record check)
- [ ] Rate limiting & throttling

---

Built with ❤️ using Python, Playwright & Streamlit
