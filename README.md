# 🧲 Maps Lead Scraper

Scrape business leads from Google Maps → formatted Excel for outreach campaigns.

![Python](https://img.shields.io/badge/Python-3.10+-blue?logo=python&logoColor=white)
![Playwright](https://img.shields.io/badge/Playwright-1.40+-2EAD33?logo=playwright&logoColor=white)
![Streamlit](https://img.shields.io/badge/Streamlit-1.28+-FF4B4B?logo=streamlit&logoColor=white)
![License](https://img.shields.io/badge/License-MIT-yellow)

## ✨ Features

- 🔍 **Auto-scrape** from Google Maps (no API key needed)
- 📞 **Phone numbers**, 🌐 **websites**, 📍 **full addresses**
- ⭐ **Ratings & reviews**, 🕐 **opening hours**
- 📊 **Formatted Excel** with headers, filters, dropdowns
- 📋 **Outreach Status** tracking (Not Contacted → Closed)
- 🎨 **Beautiful GUI** (Streamlit web app)
- 💻 **CLI mode** for automation/scripting

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

# With options
python3 gmaps_scraper.py "digital marketing agency Surabaya" --max 50 --output leads.xlsx
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
| Email | Guessed from website |
| Rating | Google rating (1-5) |
| Reviews | Number of reviews |
| Status | Open/Closed |
| Hours | Opening hours |
| Google Maps URL | Direct link |
| Description | Business description |
| **Outreach Status** | Dropdown: Not Contacted → Closed |
| **Notes** | Your notes |

## 🎯 Use Cases

- 📧 **Email outreach** campaigns
- 📞 **Cold calling** lead lists
- 🏪 **Market research** by location
- 📊 **Competitor analysis**
- 🎯 **Sales prospecting**

## ⚠️ Disclaimer

This tool is for educational purposes. Use responsibly and respect:
- Google Maps Terms of Service
- Privacy laws (GDPR, etc.)
- Anti-spam regulations

## 📝 License

MIT License

## 🤝 Contributing

Pull requests welcome! Ideas for improvement:
- [ ] Email extraction from websites
- [ ] Social media links scraping
- [ ] Multi-location batch scraping
- [ ] Google Sheets integration
- [ ] Proxy rotation support
- [ ] Export to CRM formats

---

Built with ❤️ using Python, Playwright & Streamlit
