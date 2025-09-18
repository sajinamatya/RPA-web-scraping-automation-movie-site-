from qrlib.QRBot import QRBot
from DefaultProcess import DefaultProcess

class Bot(QRBot):

    def __init__(self):
        super().__init__()
        self.process = DefaultProcess()

    def start(self):
        self.setup_platform_components()
        self.process.before_run()
        self.process.execute_run()
        self.process.after_run()

    def teardown(self):
        self.process.after_run()

  
