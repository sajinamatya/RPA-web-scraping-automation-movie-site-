import time
from qrlib.QRProcess import QRProcess
from qrlib.QRDecorators import run_item
from qrlib.QRRunItem import QRRunItem

from EmailComponent import EmailComponent
from ExcelComponent import ExcelComponent
from DatabaseComponent import DatabaseComponent
from TomatoComponent import TomatoComponent
class DefaultProcess(QRProcess):

    def __init__(self):
        super().__init__()
        self.tomato_component = TomatoComponent()
        self.database_component = DatabaseComponent()
        self.excel_component = ExcelComponent()
        self.email_component = EmailComponent()
        self.data = []
        self.movie_results = []

    @run_item(is_ticket=False)
    def before_run(self, *args, **kwargs):
        run_item: QRRunItem = kwargs["run_item"]
        self.notify(run_item)
        try:
            self.database_component._create_database_table()
        except Exception as e:
            self.database_component.logger.error(f"Database creation failed: {e}")
            raise RuntimeError("Failed to create database table") from e
        try:
            self.tomato_component._open_rotten_tomatoes_site()
        except Exception as e:
            self.tomato_component.logger.error(f"Failed to open Rotten Tomatoes site: {e}")
            raise RuntimeError("Failed to open Rotten Tomatoes site") from e
        self.tomato_component.logger.info("Database creation and website opening process complete - ready to process movies")

    @run_item(is_ticket=False, post_success=False)
    def before_run_item(self, *args, **kwargs):
        run_item: QRRunItem = kwargs["run_item"]
        self.notify(run_item)
        movie_name = args[0] if args else "Unknown"
        self.tomato_component.logger.info(f"Preparing to process: {movie_name}")

    @run_item(is_ticket=True)
    def execute_run_item(self, *args, **kwargs):
        run_item: QRRunItem = kwargs["run_item"]
        self.notify(run_item)
        try:
            movie_name = args[0] if args else "Unknown"
            movie_result = self.tomato_component._search_and_extract_movie(movie_name)
            self.database_component._insert_movie_to_db([movie_result])
            self.movie_results.append(movie_result)
            run_item.report_data[f"movie_name"] = movie_name
            run_item.report_data[f"movie_status"] = movie_result.get('status', 'Unknown')
            run_item.report_data[f"tomatometer"] = movie_result.get('tomatometer_score', 'N/A')
            run_item.report_data[f"audience"] = movie_result.get('audience_score', 'N/A')
        except Exception as e:
            run_item.report_data["movie_error"] = str(e)
            self.tomato_component.logger.error(f"Failed to process movie {args[0] if args else 'Unknown'}: {e}")
            raise

    @run_item(is_ticket=False, post_success=False)
    def after_run_item(self, *args, **kwargs):
        run_item: QRRunItem = kwargs["run_item"]
        self.notify(run_item)
        movie_name = args[0] if args else "Unknown"
        self.tomato_component.logger.info(f"Completed processing: {movie_name}")

    @run_item(is_ticket=False, post_success=False)
    def after_run(self, *args, **kwargs):
        run_item: QRRunItem = kwargs["run_item"]
        self.notify(run_item)
        try:
            if self.movie_results:
                excel_path = self.excel_component._save_to_excel(self.movie_results)
                self.email_component._send_email_notification(excel_path)
                successful_movies = [movie for movie in self.movie_results if movie.get('status') == 'Success']
                failed_movies = [movie for movie in self.movie_results if movie.get('status') != 'Success']
                run_item.report_data["total_movies_processed"] = len(self.movie_results)
                run_item.report_data["successful_count"] = len(successful_movies)
                run_item.report_data["failed_count"] = len(failed_movies)
                run_item.report_data["success_rate"] = f"{(len(successful_movies)/len(self.movie_results)*100):.1f}%" if self.movie_results else "0%"
                run_item.report_data["excel_file"] = excel_path
                self.tomato_component.logger.info(f"Final Summary: {len(successful_movies)}/{len(self.movie_results)} movies processed successfully")
        except Exception as e:
            self.tomato_component.logger.error(f"Error in final processing: {e}")
            run_item.report_data["final_processing_error"] = str(e)
        self.tomato_component.logout()
 
    def execute_run(self):
        try:
            movies_list = self.tomato_component._get_movies_from_excel()
            self.tomato_component.logger.info(f"Processing {len(movies_list)} movies from Excel file")
            for movie_name in movies_list:
                self.before_run_item(movie_name)
                self.execute_run_item(movie_name)
                self.after_run_item(movie_name)
        except Exception as e:
            self.tomato_component.logger.error(f"Error in execute_run: {e}")
            raise

