# Configuration constants for Rotten Tomatoes Bot

# File paths
EXCEL_FILE_PATH = "movies.xlsx"
OUTPUT_EXCEL_FOLDER = "output-excel"

# Database configuration
DB_DRIVER = "ODBC Driver 17 for SQL Server"
DB_TABLE_NAME = "movies"

# Browser configuration
BROWSER_TIMEOUT = 60
SEARCH_RESULTS_LIMIT = 10

# Email configuration
DEFAULT_EMAIL_SUBJECT = "Rotten Tomatoes Movie Reviews - Extraction Complete"
DEFAULT_EMAIL_BODY = "Movie scraping completed successfully. Please find the results in the attached Excel file."

# Rotten Tomatoes URLs
ROTTEN_TOMATOES_BASE_URL = "https://www.rottentomatoes.com"
ROTTEN_TOMATOES_SEARCH_URL = "https://www.rottentomatoes.com/search?search={}"

# Scraping selectors
SELECTORS = {
    "movie_filter": "css:li[data-filter='movie']",
    "search_results": "xpath://search-page-media-row[@data-qa='data-row']",
    "title_primary": "css:a[slot='title']",
    "title_fallback": "css:a[data-qa='info-name']",
    "tomatometer": "css:rt-text[slot='criticsScore']",
    "audience_score": "css:rt-text[slot='audienceScore']",
    "storyline": "css:rt-text[slot='content']",
    "rating": "css:rt-text[slot='metadataProp']",
    "genres": "xpath://rt-link[@data-qa='item-value' and contains(@href, 'genres:')]"
}