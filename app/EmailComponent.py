from RPA.Email.ImapSmtp import ImapSmtp
import os 
import logging
from dotenv import load_dotenv
from qrlib.QRComponent import QRComponent

# Load environment variables
load_dotenv()

class EmailComponent(QRComponent):
    def __init__(self):
        super().__init__()
        if not hasattr(self, 'logger') or self.logger is None:
            self.logger = logging.getLogger(__name__)
            
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