import logging
from PyQt6.QtCore import pyqtSignal, QObject


class QPlainTextEditLogger(logging.Handler, QObject):
    appendPlainText = pyqtSignal(str)

    def __init__(self, parent=None):
        logging.Handler.__init__(self)  # Khởi tạo lớp cha của logging.Handler
        QObject.__init__(self, parent)  # Khởi tạo QObject với parent
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        msg = self.format(record)
        self.appendPlainText.emit(msg)