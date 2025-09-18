
import os 
import datetime
import logging
from RPA.Excel.Files import Files
from qrlib.QRComponent import QRComponent




class ExcelComponent(QRComponent):
    def __init__(self):
        super().__init__()
        if not hasattr(self, 'logger') or self.logger is None:
            self.logger = logging.getLogger(__name__)
            
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