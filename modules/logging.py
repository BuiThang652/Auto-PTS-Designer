import logging
from PyQt6.QtCore import pyqtSignal, QObject


class QPlainTextEditLogger(logging.Handler, QObject):
    appendPlainText = pyqtSignal(str)

    def __init__(self, parent=None):
        super().__init__()
        QObject.__init__(self)
        self.setFormatter(logging.Formatter('%(asctime)s - %(levelname)s - %(message)s'))

    def emit(self, record):
        msg = self.format(record)
        self.appendPlainText.emit(msg)