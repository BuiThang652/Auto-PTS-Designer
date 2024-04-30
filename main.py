from modules import open_photoshop, drag_images_from_folder_7x10, open_existing_document, drag_images_from_folder_10x10, QPlainTextEditLogger, drag_images_from_folder_6x9, drag_images_from_folder_13x18, drag_images_from_folder_9x13
from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox
from PyQt6 import uic
import sys
import os
import logging
from PyQt6.QtCore import QThread, pyqtSignal

class DesignThread(QThread):
    def __init__(self, url, logger):
        super().__init__()
        self.url = url
        self.logger = logger  # Pass the logger to the thread

    def run(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_base = os.path.join(script_dir, "base", "page_white.jpg")
            self.logger.info("Khởi động Photoshop...")
            ps_app = open_photoshop(self.logger)
            
            if ps_app is not None:
                self.logger.info("Photoshop đã mở thành công.")
                ps_base = open_existing_document(ps_app, file_path_base)
                
                if ps_base is not None:
                    self.logger.info(f"Mở file: {file_path_base}")
                    self.logger.info(f"Bắt đầu thiết kế")
                    drag_images_from_folder_10x10(self.logger, ps_app, self.url, ps_base)
                    self.logger.info(f"Thiết kế xong.")
                else:
                    self.logger.error("Không thể mở tài liệu trong Photoshop.")
            else:
                self.logger.error("Không thể mở ứng dụng photoshop Photoshop.")
        except Exception as e:
            self.logger.error(f"Gặp lỗi: {str(e)}")

class DesignThread_710(QThread):
    def __init__(self, url, logger):
        super().__init__()
        self.url = url
        self.logger = logger  # Pass the logger to the thread

    def run(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_base = os.path.join(script_dir, "base", "page_white.jpg")
            self.logger.info("Khởi động Photoshop...")
            ps_app = open_photoshop(self.logger)
            
            if ps_app is not None:
                self.logger.info("Photoshop đã mở thành công.")
                ps_base = open_existing_document(ps_app, file_path_base)
                if ps_base is not None:
                    self.logger.info(f"Mở file: {file_path_base}")
                    self.logger.info(f"Bắt đầu thiết kế")
                    drag_images_from_folder_7x10(self.logger, ps_app, self.url, ps_base)
                    self.logger.info(f"Thiết kế xong.")
                else:
                    self.logger.error("Không thể mở tài liệu trong Photoshop.")
            else:
                self.logger.error("Không thể mở ứng dụng photoshop Photoshop.")
        except Exception as e:
            self.logger.error(f"Gặp lỗi: {str(e)}")

class DesignThread_69(QThread):
    def __init__(self, url, logger):
        super().__init__()
        self.url = url
        self.logger = logger  # Pass the logger to the thread

    def run(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_base = os.path.join(script_dir, "base", "page_white.jpg")
            self.logger.info("Khởi động Photoshop...")
            ps_app = open_photoshop(self.logger)
            
            if ps_app is not None:
                self.logger.info("Photoshop đã mở thành công.")
                ps_base = open_existing_document(ps_app, file_path_base)
                if ps_base is not None:
                    self.logger.info(f"Mở file: {file_path_base}")
                    self.logger.info(f"Bắt đầu thiết kế")
                    drag_images_from_folder_6x9(self.logger, ps_app, self.url, ps_base)
                    self.logger.info(f"Thiết kế xong.")
                else:
                    self.logger.error("Không thể mở tài liệu trong Photoshop.")
            else:
                self.logger.error("Không thể mở ứng dụng photoshop Photoshop.")
        except Exception as e:
            self.logger.error(f"Gặp lỗi: {str(e)}")

class DesignThread_1318(QThread):
    def __init__(self, url, logger):
        super().__init__()
        self.url = url
        self.logger = logger  # Pass the logger to the thread

    def run(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_base = os.path.join(script_dir, "base", "page_white.jpg")
            self.logger.info("Khởi động Photoshop...")
            ps_app = open_photoshop(self.logger)
            
            if ps_app is not None:
                self.logger.info("Photoshop đã mở thành công.")
                ps_base = open_existing_document(ps_app, file_path_base)
                if ps_base is not None:
                    self.logger.info(f"Mở file: {file_path_base}")
                    self.logger.info(f"Bắt đầu thiết kế")
                    drag_images_from_folder_13x18(self.logger, ps_app, self.url, ps_base)
                    self.logger.info(f"Thiết kế xong.")
                else:
                    self.logger.error("Không thể mở tài liệu trong Photoshop.")
            else:
                self.logger.error("Không thể mở ứng dụng photoshop Photoshop.")
        except Exception as e:
            self.logger.error(f"Gặp lỗi: {str(e)}")

class DesignThread_913(QThread):
    def __init__(self, url, logger):
        super().__init__()
        self.url = url
        self.logger = logger  # Pass the logger to the thread

    def run(self):
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            file_path_base = os.path.join(script_dir, "base", "page_white.jpg")
            self.logger.info("Khởi động Photoshop...")
            ps_app = open_photoshop(self.logger)
            
            if ps_app is not None:
                self.logger.info("Photoshop đã mở thành công.")
                ps_base = open_existing_document(ps_app, file_path_base)
                if ps_base is not None:
                    self.logger.info(f"Mở file: {file_path_base}")
                    self.logger.info(f"Bắt đầu thiết kế")
                    drag_images_from_folder_9x13(self.logger, ps_app, self.url, ps_base)
                    self.logger.info(f"Thiết kế xong.")
                else:
                    self.logger.error("Không thể mở tài liệu trong Photoshop.")
            else:
                self.logger.error("Không thể mở ứng dụng photoshop Photoshop.")
        except Exception as e:
            self.logger.error(f"Gặp lỗi: {str(e)}")


class Login(QMainWindow):
    def __init__(self):
        super().__init__()
        self.ui = uic.loadUi("gui/home.ui", self)
        logTextBox = QPlainTextEditLogger(self)
        logTextBox.appendPlainText.connect(self.activity.append)
        self.logger = logging.getLogger('AppLogger')
        self.logger.addHandler(logTextBox)
        self.logger.setLevel(logging.DEBUG)
        self.logger.info("Open App Auto")
        self.ui.btn_url_1010.clicked.connect(self.select_url)
        self.ui.btn_design.clicked.connect(self.start_design_thread)
        self.ui.btn_design_1.clicked.connect(self.start_design_thread_710)
        self.ui.btn_design_2.clicked.connect(self.start_design_thread_69)
        self.ui.btn_design_3.clicked.connect(self.start_design_thread_1318)
        self.ui.btn_design_4.clicked.connect(self.start_design_thread_913)
        
    def select_url(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục", "/")
        if folder_path:
            self.url_1010 = folder_path
            self.ui.btn_url_1010.setText(folder_path)
            self.logger.info(f"Đã chọn thư mục: {folder_path}")
    
    def start_design_thread(self):
        if hasattr(self, 'url_1010'):
            self.design_thread = DesignThread(self.url_1010, self.logger)
            self.design_thread.start()
        else:
            self.logger.warning("Vui lòng chọn một thư mục trước.")
            
    def start_design_thread_710(self):
        if hasattr(self, 'url_1010') and self.url_1010:
            self.design_thread_710 = DesignThread_710(self.url_1010, self.logger)
            self.design_thread_710.start()
        else:
            self.logger.warning("Vui lòng chọn một thư mục trước.")
            
    def start_design_thread_69(self):
        if hasattr(self, 'url_1010') and self.url_1010:
            self.design_thread_69 = DesignThread_69(self.url_1010, self.logger)
            self.design_thread_69.start()
        else:
            self.logger.warning("Vui lòng chọn một thư mục trước.")

    def start_design_thread_1318(self):
        if hasattr(self, 'url_1010') and self.url_1010:
            self.design_thread_1318 = DesignThread_1318(self.url_1010, self.logger)
            self.design_thread_1318.start()
        else:
            self.logger.warning("Vui lòng chọn một thư mục trước.")
            
    def start_design_thread_913(self):
        if hasattr(self, 'url_1010') and self.url_1010:
            self.design_thread_913 = DesignThread_913(self.url_1010, self.logger)
            self.design_thread_913.start()
        else:
            self.logger.warning("Vui lòng chọn một thư mục trước.")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Login()
    window.setWindowTitle("Auto designer")
    window.show()
    sys.exit(app.exec())
