import pyodbc
from qrlib.QRComponent import QRComponent
import os 
from dotenv import load_dotenv
# Load the .env file
load_dotenv()

# Now these will work


import logging

class DatabaseComponent(QRComponent):
    def __init__(self):
        super().__init__()
        if not hasattr(self, 'logger') or self.logger is None:
            self.logger = logging.getLogger(__name__)
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
            self.logger.info("Database table 'movies' created ")
        except pyodbc.Error as e:
            raise RuntimeError(f"Database table creation failed: {e}")
        finally:
            if conn:
                conn.close()

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