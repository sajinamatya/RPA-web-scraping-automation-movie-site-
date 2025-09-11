from robocorp.tasks import task
from RPA.Browser.Selenium import Selenium
from RPA.Database import Database
from RPA.Excel.Files import Files
import re
import os
from dotenv import load_dotenv
import pyodbc

EXCEL_FILE_PATH = "movies.xlsx"
browser = Selenium()

load_dotenv(os.path.join(os.path.dirname(__file__), '.env'))


@task
def robot_rottentomatoes():
    """ Main task function to run the Rotten Tomatoes movie data extraction and save to DB and send mail.
    """
    open_rotten_tomatoes_sites()
    movies_name = get_movie_name_from_excel()
    results = search_and_extract_movies(movies_name)
    save_results_to_db(results)
    browser.close_all_browsers()


def open_rotten_tomatoes_sites():
    """
    Opens the Rotten Tomatoes website in a web browser.
    """
    try:
        browser.open_available_browser("https://www.rottentomatoes.com", headless=False)
        browser.set_selenium_timeout(60)  # Set timeout to 60 seconds
    except Exception as e:
        print(f"Error navigating to Rotten Tomatoes: {e}")


def Search_movie_on_rotten_tomatoes(movie_name):
    """
    Searches for a movie on Rotten Tomatoes.
    """
    search_url = f"https://www.rottentomatoes.com/search?search={movie_name.replace(' ', '%20')}"
    browser.go_to(search_url)

def select_movie_section():
    try:
        # Always click the 'Movies' filter to ensure only movies are shown
        browser.wait_until_element_is_visible("css:li[data-filter='movie']", timeout=10)
        browser.click_element_when_visible("css:li[data-filter='movie']")
    except Exception as e:
        print("    Could not click 'Movies' filter (may already be selected or missing):", e)


def get_movie_name_from_excel():
    """  Read the movie names from the Excel file and return them as a list. """
    read_excel = Files()
    read_excel.open_workbook(EXCEL_FILE_PATH)
    excel_sheet = read_excel.read_worksheet_as_table(header=True)
    read_excel.close_workbook()
    movie_names = [row["Movies"] for row in excel_sheet if "Movies" in row]
    print(movie_names)
    return movie_names


def search_and_extract_movies(movies_name):
    """
    Combined function that searches for movies and extracts detailed information.
    Returns a list of dictionaries containing all movie data.
    """
    total_movies = len(movies_name)
    processed_count = 0
    movie_data = []
    
    for movie_name in movies_name:
        processed_count += 1
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
            print(f"\n[{processed_count}/{total_movies}] Processing movie: '{movie_name}'")
            
            # Search for the movie
            Search_movie_on_rotten_tomatoes(movie_name)
            select_movie_section()

 
            results = browser.find_elements("xpath://search-page-media-row[@data-qa='data-row']")[:10]  # Limit to 10 results
  

            exact_matches = []
            print(f"Found {len(results)} results for '{movie_name}'")
            
            # Look for exact matches
            for i, result in enumerate(results):
                try:
                    # Try to get year from the custom element's attributes first
                    year = 0
                    try:
                        # Try multiple year-related attributes
                        year_attrs = ["releaseyear", "startyear", "endyear", "year"]
                        for attr in year_attrs:
                            year_attr = browser.get_element_attribute(result, attr)
                            if year_attr and year_attr.isdigit() and int(year_attr) > 1900:
                                year = int(year_attr)
                                print(f"    Result {i}: Year from '{attr}' attribute: {year}")
                                break
                        
                        if year == 0:
                            # Print all attributes for debugging
                            try:
                                all_attrs = browser.execute_javascript("""
                                    var element = arguments[0];
                                    var attrs = {};
                                    for (var i = 0; i < element.attributes.length; i++) {
                                        var attr = element.attributes[i];
                                        attrs[attr.name] = attr.value;
                                    }
                                    return attrs;
                                """, result)
                                print(f"    Result {i}: All attributes: {all_attrs}")
                            except:
                                pass
                    except Exception as e:
                        print(f"    Result {i}: Error getting year from attributes: {e}")
                    
                    # Get title from the custom element link - now we need to look inside the shadow DOM
                    # Since we can't access shadow DOM directly, we'll use the slotted content
                    try:
                        title_element = browser.find_element("css:a[slot='title']", parent=result)
                        title = browser.get_text(title_element).strip()
                    except:
                        # Fallback: try to find the link another way
                        try:
                            title_element = browser.find_element("css:a[data-qa='info-name']", parent=result)
                            title = browser.get_text(title_element).strip()
                        except:
                            print(f"    Result {i}: Could not find title element")
                            continue
                    
                    # If we didn't get year from attribute, try DOM (though likely won't work with shadow DOM)
                    if year == 0:
                        try:
                            year_element = None
                            year_selectors = [
                                "//span[@data-qa='info-year']",
                                ".//span[@data-qa='info-year']", 
                                ".//span[contains(@class, 'year')]",
                                ".//span[contains(text(), '(')]"
                            ]
                            
                            for selector in year_selectors:
                                try:
                                    year_element = browser.find_element(selector, parent=result)
                                    break
                                except:
                                    continue
                            
                            if year_element:
                                year_text = browser.get_text(year_element).strip()
                                match = re.search(r"\d{4}", year_text)
                                year = int(match.group()) if match else 0
                                print(f"    Result {i}: '{title}' ({year}) - Year text: '{year_text}'")
                            else:
                                print(f"    Result {i}: '{title}' (No year element found)")
                        except Exception as ye:
                            print(f"    Result {i}: '{title}' (No year found) - Error: {ye}")
                    else:
                        print(f"    Result {i}: '{title}' ({year}) - From element attribute")
                        
                    print(f"    Comparing: '{title.lower()}' vs '{movie_name.strip().lower()}'")
                    if title.lower() == movie_name.strip().lower():
                        print(f"    EXACT MATCH: '{title}' ({year})")
                        exact_matches.append((title, year, title_element))
                    else:
                        print(f"    Not a match")
                except Exception as e:
                    print(f"Error processing result {i}: {e}")
                    continue
            
            # Process exact matches
            if exact_matches:
                try:
                    # Debug: Show all exact matches found
                    print(f"Found {len(exact_matches)} exact matches:")
                    for title, year, element in exact_matches:
                        print(f"  - '{title}' ({year})")
                    
                    most_recent = max(exact_matches, key=lambda x: x[1])
                    print(f"Selected most recent: '{most_recent[0]}' ({most_recent[1]})")
                    print(most_recent[2])
                    # Click the movie link
                    browser.click_element(most_recent[2])
                    
                    # Wait for the page to load
                    browser.wait_until_element_is_visible("css:rt-text[slot='criticsScore']", timeout=15)
                    
                    # Extract scores using multiple robust strategies
                    tomatometer = "N/A"
                    audience = "N/A"
                    
                    #  Direct rt-text slot extraction
                    try:
                        tomatometer_element = browser.find_element("css:rt-text[slot='criticsScore']")
                        tomatometer = browser.get_text(tomatometer_element).strip()
                        if not tomatometer:
                            tomatometer = browser.get_element_attribute(tomatometer_element, "textContent").strip()
                    except:
                        #  Try data-qa selectors
                        try:
                            tomatometer_element = browser.find_element("css:[data-qa='tomatometer-score']")
                            tomatometer = browser.get_text(tomatometer_element).strip()
                        except:
                            #  Try score-board attributes
                            try:
                                score_board = browser.find_element("css:score-board")
                                tomatometer = browser.get_element_attribute(score_board, "tomatometerscore")
                            except:
                                pass

                    try:
                        audience_element = browser.find_element("css:rt-text[slot='audienceScore']")
                        audience = browser.get_text(audience_element).strip()
                        if not audience:
                            audience = browser.get_element_attribute(audience_element, "textContent").strip()
                    except:
                        # Try data-qa selectors
                        try:
                            audience_element = browser.find_element("css:[data-qa='audience-score']")
                            audience = browser.get_text(audience_element).strip()
                        except:
                            #  Try score-board attributes
                            try:
                                score_board = browser.find_element("css:score-board")
                                audience = browser.get_element_attribute(score_board, "audiencescore")
                            except:
                                pass
                    
                    print(f"Tomatometer: {tomatometer}")
                    print(f"Audience Score: {audience}")
                    
                    # Extract storyline with multiple strategies
                    storyline = ""
                    try:
                        # rt-text with slot="content"
                        storyline_element = browser.find_element("css:rt-text[slot='content']")
                        storyline = browser.get_text(storyline_element).strip()
                    except:
                        try:
                            #  synopsis-wrap container
                            storyline_element = browser.find_element("css:div.synopsis-wrap rt-text[data-qa='synopsis-value']")
                            storyline = browser.get_text(storyline_element).strip()
                        except:
                            try:
                                # Generic synopsis selectors
                                storyline_element = browser.find_element("css:[data-qa='synopsis'], .synopsis, .plot-synopsis")
                                storyline = browser.get_text(storyline_element).strip()
                            except:
                                pass
                    
                    # Extract rating with multiple strategies
                    rating = ""
                    try:
                        # rt-text with slot="metadataProp"
                        rating_element = browser.find_element("css:rt-text[slot='metadataProp']")
                        rating = browser.get_text(rating_element).strip()
                        if rating and rating.endswith(','):
                            rating = rating[:-1].strip()
                    except:
                        try:
                            #  category-wrap data-qa
                            rating_element = browser.find_element("css:div.category-wrap[data-qa='item'] rt-text[data-qa='item-value']")
                            rating = browser.get_text(rating_element).strip()
                        except:
                            try:
                                # Generic rating selectors
                                rating_element = browser.find_element("css:[data-qa='rating'], .rating, .mpaa-rating")
                                rating = browser.get_text(rating_element).strip()
                            except:
                                pass
                    
                    # Extract genres with direct approach
                    genres = ""
                    try:
                        genre_elements = browser.find_elements("xpath://rt-link[@data-qa='item-value' and contains(@href, 'genres:')]")
                        genres = ", ".join([browser.get_text(elem).strip() for elem in genre_elements])
                    except:
                        genres = ""
                    
                    # Extract top 5 critic reviews with multiple strategies
                    reviews = []
                    try:
                        #  media-review-card-critic elements
                        review_elements = browser.find_elements("css:media-review-card-critic rt-text[data-qa='review-text']")
                        for elem in review_elements[:5]:
                            review_text = browser.get_text(elem).strip()
                            if review_text:
                                reviews.append(review_text)
                    except:
                        pass
                    
                    #  try other selectors
                    if len(reviews) < 5:
                        try:
                            additional_selectors = [
                                "css:div[data-qa='review-quote']",
                                "css:blockquote",
                                "css:.the_review",
                                "css:p.review-quote",
                                "css:div.review_quote",
                                "css:p[data-qa='review-quote']",
                                "css:div.review-text",
                                "css:div.review__text"
                            ]
                            for selector in additional_selectors:
                                if len(reviews) >= 5:
                                    break
                                try:
                                    elements = browser.find_elements(selector)
                                    for elem in elements:
                                        if len(reviews) >= 5:
                                            break
                                        review_text = browser.get_text(elem).strip()
                                        if review_text and review_text not in reviews:
                                            reviews.append(review_text)
                                except:
                                    continue
                        except:
                            pass
                    
                    # Clean and format extracted data
                    def clean_text(text):
                        """Clean and normalize text"""
                        if not text:
                            return ""
                        # Remove extra whitespace and normalize
                        import html
                        try:
                            cleaned = html.unescape(text)
                        except:
                            cleaned = text
                        cleaned = re.sub(r"\s+", " ", cleaned).strip()
                        return cleaned
                    
                    # Apply cleaning to extracted data
                    tomatometer = clean_text(tomatometer) if tomatometer != "N/A" else "N/A"
                    audience = clean_text(audience) if audience != "N/A" else "N/A"
                    storyline = clean_text(storyline)
                    rating = clean_text(rating)
                    # Remove any trailing or extra slashes in genres (e.g., 'Drama/ ' -> 'Drama')
                    genres = clean_text(genres)
                    genres = re.sub(r"\s*/\s*", ", ", genres).strip().strip(',')
                    genres = re.sub(r",\s*$", "", genres)  # Remove trailing comma if any
                    reviews = [clean_text(review) for review in reviews]
                    
                    # Ensure we have 5 review slots
                    while len(reviews) < 5:
                        reviews.append("")
                    
                    # Update movie dictionary with extracted data
                    movie_dict['tomatometer_score'] = tomatometer if tomatometer else 'N/A'
                    movie_dict['audience_score'] = audience if audience else 'N/A'
                    movie_dict['storyline'] = storyline if storyline else 'N/A'
                    movie_dict['rating'] = rating if rating else 'N/A'
                    movie_dict['genres'] = genres if genres else 'N/A'
                    for idx, review in enumerate(reviews):
                        movie_dict[f'review_{idx+1}'] = review if review else 'N/A'
                    movie_dict['status'] = 'Success'
                    
                    print(f"SUCCESS: Extracted data for '{movie_name}'")
                    print(f"   Tomatometer: {tomatometer}, Audience: {audience}")
                    
                except Exception as e:
                    print(f"Error extracting details for '{movie_name}': {e}")
                    # Try to go back to search results for next movie
                    try:
                        browser.go_back()
                    except:
                        pass
            else:
                print(f"No exact match found for '{movie_name}'")
                
        except Exception as e:
            print(f"Error processing '{movie_name}': {e}")
            print("Continuing to next movie...")
        
        movie_data.append(movie_dict)
    
    print(f"\nCompleted processing {processed_count} movies.")
    return movie_data



import pyodbc

def save_results_to_db(movie_data):
    server = os.environ.get("SERVER")
    database = os.environ.get("DATABASE")

    conn_str = (
        f"DRIVER={{ODBC Driver 17 for SQL Server}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"Trusted_Connection=yes;"
    )

    try:
        conn = pyodbc.connect(conn_str)
        cursor = conn.cursor()

        # Create table if not exists
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

        # Insert data
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
        print("Data saved successfully")

    except Exception as e:
        print(f"Database error: {e}")
        raise
    finally:
        if 'conn' in locals():
            conn.close()
