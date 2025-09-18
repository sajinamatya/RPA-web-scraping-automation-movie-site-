from qrlib.QRComponent import QRComponent
from RPA.Browser.Selenium import Selenium
from RPA.Excel.Files import Files


import logging

class TomatoComponent(QRComponent):
    def __init__(self):
        super().__init__()
        self.browser = Selenium()
        self.excel_file_path = "movies.xlsx"  # Default Excel file path
        if not hasattr(self, 'logger') or self.logger is None:
            self.logger = logging.getLogger(__name__)
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
                    movie_dict['status'] = f'Error: {str(e)}'
            else:
                self.logger.warning(f"No exact match found for '{movie_name}'")
                movie_dict['status'] = 'No exact match found'
        except Exception as e:
            self.logger.error(f"Error processing '{movie_name}': {e}")
            movie_dict['status'] = f'Error: {str(e)}'
        return movie_dict
    

    def _search_movie(self, movie_name):
        """Search for a movie on Rotten Tomatoes"""
        search_url = f"https://www.rottentomatoes.com/search?search={movie_name.replace(' ', '%20')}"
        self.browser.go_to(search_url)
    

    def _select_movie_section(self):
        """Click on Movies filter"""
        try:
            # Try multiple selectors for the Movies filter
            selectors = [
                "css:li[data-filter='movie']",
                "css:button[data-filter='movie']", 
                "css:[data-filter='movie']",
                "xpath://li[contains(@class, 'movie')]",
                "xpath://button[contains(text(), 'Movies')]"
            ]
            
            element_found = False
            for selector in selectors:
                try:
                    self.browser.wait_until_element_is_visible(selector, timeout=5)
                    self.browser.click_element_when_visible(selector)
                    self.logger.info(f"Successfully clicked Movies filter using selector: {selector}")
                    element_found = True
                    break
                except Exception:
                    continue
            
            if not element_found:
                self.logger.warning("Movies filter not found, continuing without filtering")
                # Don't raise error, just continue - some search pages might not have filters
                
        except Exception as e:
            self.logger.warning(f"Could not click 'Movies' filter, continuing anyway: {e}")
            # Don't raise error to allow processing to continue
    
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
    
    def logout(self):
        """Close browser and clean up resources"""
        try:
            if hasattr(self, 'browser') and self.browser:
                self.browser.close_all_browsers()
                self.logger.info("Browser closed successfully")
        except Exception as e:
            self.logger.error(f"Error closing browser: {e}")
            # Don't raise exception during cleanup