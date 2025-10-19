# RPA Web Scraping Automation - Movie Site

A Robotic Process Automation (RPA) solution for automated web scraping of movie data from websites. This Python-based project demonstrates efficient data extraction techniques for collecting movie information.

## 🎬 Overview

This project implements an automated web scraping system designed to extract movie-related information from various movie websites. Built with Python, it provides a robust solution for collecting, processing, and storing movie data efficiently.

## ✨ Features

- **Automated Data Extraction**: Scrapes movie information including titles, ratings, release dates, and more
- **RPA Integration**: Utilizes Robotic Process Automation principles for reliable and repeatable scraping
- **Data Processing**: Cleans and structures scraped data for easy analysis
- **Error Handling**: Robust error handling to manage website changes and connection issues
- **Export Capabilities**: Save scraped data in various formats (CSV, JSON, etc.)

## 🌿 Branches

This repository is organized into multiple branches, each focusing on different features and implementations:

### 🔹 `master`
The main branch containing the core web scraping functionality and base implementation.

### 🔹 `database-connection`
Implements database integration for storing scraped movie data. This branch includes:
- Database schema setup
- Connection handling
- Data persistence layer
- CRUD operations for movie records

### 🔹 `excel_ouput_and_smtp`
Adds advanced export and notification features:
- Excel file generation for scraped data
- SMTP email integration
- Automated email reports with attached data
- Custom formatting for Excel outputs

### 🔹 `robot_format`
Focuses on RPA framework compatibility and robot file formats:
- Robot Framework integration
- Standardized robot file structure
- Test automation capabilities
- Enhanced RPA workflow definitions

**To switch between branches:**
```bash
git checkout <branch-name>
```

## 🛠️ Technologies Used

- **Python**: Core programming language
- **Web Scraping Libraries**: BeautifulSoup4, Selenium, or Scrapy
- **RPA Framework**: UiPath, Automation Anywhere, or similar
- **Data Processing**: Pandas for data manipulation
- **HTTP Requests**: Requests library for web communication

## 📋 Prerequisites

Before running this project, ensure you have the following installed:

- Python 3.7 or higher
- pip (Python package installer)
- Web browser driver (ChromeDriver or GeckoDriver for Selenium)

## 🚀 Installation

1. Clone the repository:
```bash
git clone https://github.com/sajinamatya/RPA-web-scraping-automation-movie-site-.git
cd RPA-web-scraping-automation-movie-site-
```

2. Choose your desired branch:
```bash
# For database functionality
git checkout database-connection

# For Excel and email features
git checkout excel_ouput_and_smtp

# For Robot Framework format
git checkout robot_format
```

3. Install required dependencies:
```bash
pip install -r requirements.txt
```

4. Configure your settings (if applicable):
```bash
# Update configuration file with target URLs and scraping parameters
```

## 💻 Usage

Run the main scraping script:

```bash
python main.py
```

Or import and use specific modules:

```python
from scraper import MovieScraper

scraper = MovieScraper()
movies = scraper.scrape_movies()
scraper.save_data(movies, 'output.csv')
```

## 📁 Project Structure

```
RPA-web-scraping-automation-movie-site-/
│
├── main.py                 # Main execution script
├── scraper.py             # Core scraping logic
├── data_processor.py      # Data cleaning and processing
├── config.py              # Configuration settings
├── requirements.txt       # Python dependencies
├── output/               # Scraped data output directory
└── README.md             # Project documentation
```

## ⚙️ Configuration

Customize the scraping behavior by modifying the configuration file:

- Target URLs
- Scraping intervals
- Data fields to extract
- Output format preferences
- Browser settings (for Selenium)

## 📊 Output

The scraped data typically includes:

- Movie titles
- Release dates
- Ratings and reviews
- Cast and crew information
- Plot summaries
- Genre classifications
- Box office information

## ⚠️ Important Notes

- **Respect robots.txt**: Always check and respect the target website's robots.txt file
- **Rate Limiting**: Implement appropriate delays between requests to avoid overloading servers
- **Terms of Service**: Ensure compliance with website terms of service
- **Legal Considerations**: Web scraping may have legal implications depending on jurisdiction

## 👤 Author

**Sajin Amatya**

- GitHub: [@sajinamatya](https://github.com/sajinamatya)

---

**Disclaimer**: This tool is for educational purposes only. Always ensure you have permission to scrape websites and comply with their terms of service.
