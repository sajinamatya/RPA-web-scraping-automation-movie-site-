import os
import re
import datetime
from dotenv import load_dotenv
import pyodbc
from qrlib.QRComponent import QRComponent
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files
from RPA.Email.ImapSmtp import ImapSmtp


class DefaultComponent(QRComponent):
    
    def __init__(self):
        super().__init__()
        self.sel = Selenium(auto_close=False)
        self.browser = Selenium()
        self.excel_file_path = "movies.xlsx"
        # Load environment variables
        load_dotenv(os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env'))
        self.movie_results = []

    def test(self):
        try:
            self.logger.info("Test task - Movie scraper component ready")
        except Exception as e:
            self.logger.error("Test task failed")
            raise RuntimeError(f"Test task failed: {e}")
    
    # Private helper methods for movie scraping
    
    def _create_database_table(self):
        """Create movies table in database"""
        server = os.environ.get("SERVER")
        database = os.environ.get("DATABASE")
        
        if not server or not database:
            raise ValueError("Missing required environment variables: SERVER and/or DATABASE")
        
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection=yes;"
        )
        conn = None
        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            cursor.execute("""
                IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='movies' AND xtype='U')
                CREATE TABLE movies (
                    id INT IDENTITY(1,1) PRIMARY KEY,
                    movie_name NVARCHAR(255),
                    tomatometer_score NVARCHAR(50),
                    audience_score NVARCHAR(50),
                    storyline NVARCHAR(MAX),
                    rating NVARCHAR(50),
                    genres NVARCHAR(255),
                    review_1 NVARCHAR(MAX),
                    review_2 NVARCHAR(MAX),
                    review_3 NVARCHAR(MAX),
                    review_4 NVARCHAR(MAX),
                    review_5 NVARCHAR(MAX),
                    status NVARCHAR(50)
                )
            """)
            conn.commit()
            self.logger.info("Database table 'movies' created or already exists")
        except pyodbc.Error as e:
            raise RuntimeError(f"Database table creation failed: {e}")
        finally:
            if conn:
                conn.close()
    
    def _open_rotten_tomatoes_site(self):
        """Open Rotten Tomatoes website"""
        try:
            self.browser.open_available_browser("https://www.rottentomatoes.com", headless=False)
            self.browser.set_selenium_timeout(60)
            self.logger.info("Opened Rotten Tomatoes website")
        except Exception as e:
            self.logger.error(f"Failed to navigate to Rotten Tomatoes: {e}")
            raise RuntimeError(f"Failed to navigate to Rotten Tomatoes: {e}")
    
    def _get_movies_from_excel(self):
        """Read movie names from Excel file"""
        try:
            read_excel = Files()
            read_excel.open_workbook(self.excel_file_path)
            excel_sheet = read_excel.read_worksheet_as_table(header=True)
            read_excel.close_workbook()
            movie_names = [row["Movies"] for row in excel_sheet if "Movies" in row]
            
            if not movie_names:
                raise ValueError(f"No movies found in Excel file '{self.excel_file_path}'. Make sure the file has a 'Movies' column with data.")
                
            self.logger.info(f"Found {len(movie_names)} movies in Excel file")
            return movie_names
        except FileNotFoundError as e:
            self.logger.error(f"Excel file '{self.excel_file_path}' not found. Please ensure the file exists.")
            raise FileNotFoundError(f"Excel file '{self.excel_file_path}' not found: {e}")
        except Exception as e:
            self.logger.error(f"Error reading movie names from Excel file '{self.excel_file_path}': {e}")
            raise ValueError(f"Error reading movie names from Excel file '{self.excel_file_path}': {e}")
    
    def _search_and_extract_movie(self, movie_name):
        """Search and extract data for a single movie"""
        movie_dict = {
            'movie_name': movie_name,
            'tomatometer_score': 'N/A',
            'audience_score': 'N/A',
            'storyline': 'N/A',
            'rating': 'N/A',
            'genres': 'N/A',
            'review_1': 'N/A',
            'review_2': 'N/A',
            'review_3': 'N/A',
            'review_4': 'N/A',
            'review_5': 'N/A',
            'status': 'No exact match found'
        }
        try:
            self.logger.info(f"Processing movie: '{movie_name}'")
            self._search_movie(movie_name)
            self._select_movie_section()
            results = self.browser.find_elements("xpath://search-page-media-row[@data-qa='data-row']")[:10]
            exact_matches = []
            self.logger.info(f"Found {len(results)} results for '{movie_name}'")
            for i, result in enumerate(results):
                try:
                    title_element = None
                    try:
                        title_element = self.browser.find_element("css:a[slot='title']", parent=result)
                        title = self.browser.get_text(title_element).strip()
                    except Exception as e1:
                        try:
                            title_element = self.browser.find_element("css:a[data-qa='info-name']", parent=result)
                            title = self.browser.get_text(title_element).strip()
                        except Exception as e2:
                            self.logger.warning(f"Error finding title element for result {i}: {e2}")
                            raise
                    year = 2020
                    if title.lower() == movie_name.strip().lower():
                        self.logger.info(f"EXACT MATCH: '{title}'")
                        exact_matches.append((title, year, title_element))
                except Exception as e:
                    self.logger.warning(f"Error processing result {i}: {e}")
                    raise
            if exact_matches:
                try:
                    most_recent = max(exact_matches, key=lambda x: x[1])
                    self.logger.info(f"Selected: '{most_recent[0]}'")
                    self.browser.click_element(most_recent[2])
                    self.browser.wait_until_element_is_visible("css:rt-text[slot='criticsScore']", timeout=15)
                    tomatometer = self._extract_tomatometer_score()
                    audience = self._extract_audience_score()
                    movie_dict['tomatometer_score'] = tomatometer
                    movie_dict['audience_score'] = audience
                    movie_dict['status'] = 'Success'
                    self.logger.info(f"SUCCESS: Extracted data for '{movie_name}' - Tomatometer: {tomatometer}, Audience: {audience}")
                except Exception as e:
                    self.logger.error(f"Error extracting details for '{movie_name}': {e}")
                    try:
                        self.browser.go_back()
                    except Exception as e2:
                        self.logger.error(f"Error going back after extraction failure: {e2}")
                    raise
            else:
                self.logger.warning(f"No exact match found for '{movie_name}'")
                raise RuntimeError(f"No exact match found for '{movie_name}'")
        except Exception as e:
            self.logger.error(f"Error processing '{movie_name}': {e}")
            movie_dict['status'] = f'Error: {str(e)}'
            raise RuntimeError(f"Error processing '{movie_name}': {e}")
        return movie_dict
    
    def _search_movie(self, movie_name):
        """Search for a movie on Rotten Tomatoes"""
        search_url = f"https://www.rottentomatoes.com/search?search={movie_name.replace(' ', '%20')}"
        self.browser.go_to(search_url)
    
    def _select_movie_section(self):
        """Click on Movies filter"""
        try:
            self.browser.wait_until_element_is_visible("css:li[data-filter='movie']", timeout=10)
            self.browser.click_element_when_visible("css:li[data-filter='movie']")
        except Exception as e:
            self.logger.error(f"Could not click 'Movies' filter: {e}")
            raise RuntimeError(f"Could not click 'Movies' filter: {e}")
    
    def _extract_tomatometer_score(self):
        """Extract Tomatometer score"""
        try:
            tomatometer_element = self.browser.find_element("css:rt-text[slot='criticsScore']")
            return self.browser.get_text(tomatometer_element).strip()
        except:
            return "N/A"
    
    def _extract_audience_score(self):
        """Extract Audience score"""
        try:
            audience_element = self.browser.find_element("css:rt-text[slot='audienceScore']")
            return self.browser.get_text(audience_element).strip()
        except:
            return "N/A"
    
    def _insert_movie_to_db(self, movie_data):
        """Insert movie data to database"""
        server = os.environ.get("SERVER")
        database = os.environ.get("DATABASE")
        if not movie_data:
            return
        conn_str = (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={server};"
            f"DATABASE={database};"
            f"Trusted_Connection=yes;"
        )
        conn = None
        try:
            conn = pyodbc.connect(conn_str)
            cursor = conn.cursor()
            insert_sql = """
                INSERT INTO movies (movie_name, tomatometer_score, audience_score, storyline, 
                                  rating, genres, review_1, review_2, review_3, review_4, review_5, status)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """
            for movie in movie_data:
                cursor.execute(insert_sql, (
                    movie.get("movie_name", ""),
                    movie.get("tomatometer_score", ""),
                    movie.get("audience_score", ""),
                    movie.get("storyline", ""),
                    movie.get("rating", ""),
                    movie.get("genres", ""),
                    movie.get("review_1", ""),
                    movie.get("review_2", ""),
                    movie.get("review_3", ""),
                    movie.get("review_4", ""),
                    movie.get("review_5", ""),
                    movie.get("status", "")
                ))
            conn.commit()
            self.logger.info(f"Successfully inserted {len(movie_data)} movie(s) to database")
        except Exception as e:
            self.logger.error(f"Database insertion failed: {e}")
            raise RuntimeError(f"Database insertion failed: {e}")
        finally:
            if conn:
                conn.close()
    
    def _save_to_excel(self, movie_data):
        """Save movie data to Excel file"""
        if not movie_data:
            raise ValueError("No movie data provided for Excel export")
        output_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'output-excel')
        os.makedirs(output_folder, exist_ok=True)
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        excel_path = os.path.join(output_folder, f"movie_reviews_{timestamp}.xlsx")
        try:
            excel = Files()
            rows = [
                [
                    movie.get('movie_name', ''),
                    movie.get('tomatometer_score', ''),
                    movie.get('audience_score', ''),
                    movie.get('storyline', ''),
                    movie.get('rating', ''),
                    movie.get('genres', ''),
                    movie.get('review_1', ''),
                    movie.get('review_2', ''),
                    movie.get('review_3', ''),
                    movie.get('review_4', ''),
                    movie.get('review_5', ''),
                    movie.get('status', '')
                ]
                for movie in movie_data
            ]
            headers = [
                "Movie Name", "Tomatometer Score", "Audience Score", "Storyline", "Rating", "Genres",
                "Review 1", "Review 2", "Review 3", "Review 4", "Review 5", "Status"
            ]
            excel.create_workbook(excel_path)
            excel.create_worksheet("Movie Data")
            excel.append_rows_to_worksheet([headers] + rows)
            excel.save_workbook(excel_path)
            excel.close_workbook()
            self.logger.info(f"Successfully saved {len(movie_data)} movies to Excel file: {excel_path}")
            return excel_path
        except Exception as e:
            self.logger.error(f"Failed to save data to Excel file '{excel_path}': {e}")
            raise ValueError(f"Failed to save data to Excel file '{excel_path}': {e}")
    
    def _send_email_notification(self, excel_path):
        """Send email notification with Excel attachment"""
        try:
            smtp_server = os.environ.get("SMTP_SERVER")
            smtp_port = os.environ.get("SMTP_PORT", 587)
            smtp_user = os.environ.get("SMTP_USER")
            smtp_password = os.environ.get("SMTP_PASSWORD")
            if not all([smtp_server, smtp_user, smtp_password]):
                self.logger.warning("SMTP configuration missing - skipping email notification")
                raise RuntimeError("SMTP configuration missing")
            recipients = ["sajinamatya88@gmail.com"]
            email = ImapSmtp(smtp_server=smtp_server, smtp_port=int(smtp_port))
            email.authorize(account=smtp_user, password=smtp_password)
            subject = "Rotten Tomatoes Movie Reviews - Extraction Complete"
            body = "Movie scraping completed successfully. Please find the results in the attached Excel file."
            email.send_message(
                sender=smtp_user,
                recipients=recipients,
                subject=subject,
                body=body,
                attachments=[excel_path]
            )
            self.logger.info(f"Email sent successfully to: {', '.join(recipients)}")
        except Exception as e:
            self.logger.error(f"Failed to send email notification: {e}")
            raise RuntimeError(f"Failed to send email notification: {e}")
