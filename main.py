from modules import drag_images_from50_folder_7x10, drag_images_from50_folder_9x13, drag_images_from50_folder_10x15, drag_images_from50_folder_13x18, resource_path, open_photoshop, drag_images_from_folder_7x10, open_existing_document, drag_images_from_folder_10x10, QPlainTextEditLogger, drag_images_from_folder_6x9, drag_images_from_folder_13x18, drag_images_from_folder_9x13
from PyQt6.QtWidgets import QMainWindow, QApplication, QFileDialog, QMessageBox, QRadioButton 
from PyQt6 import uic
from PyQt6.QtGui import QIcon
import sys
import os
import logging
from PyQt6.QtCore import QThread, pyqtSignal, Qt



class DesignThread(QThread):
    designCompleted = pyqtSignal(str)
    designBtn = pyqtSignal()
    
    def __init__(self, url, logger, selected_size, selected_paper):
        super().__init__()
        self.url = url
        self.logger = logger 
        self.selected_size = selected_size
        self.selected_paper = selected_paper

    def run(self):
        try:
            if self.selected_paper == "input_60":
                file_path_base = resource_path("base/page_white.jpg")
            elif self.selected_paper == "input_50":
                file_path_base = resource_path("base/50cm.jpg")
            else:
                self.logger.warning("Không có khổ giấy này.")
            self.logger.info("Khởi động Photoshop...")
            ps_app = open_photoshop(self.logger)
            
            if ps_app is not None:
                self.logger.info("Photoshop đã mở thành công.")
                ps_base = open_existing_document(ps_app, file_path_base)
                
                if ps_base is not None:
                    self.logger.info(f"Mở file: {file_path_base}")
                    self.logger.info(f"Bắt đầu thiết kế")
                    if self.selected_paper == "input_60" and self.selected_size == "input_1015":
                        drag_images_from_folder_10x10(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_60" and self.selected_size == "input_69":
                        drag_images_from_folder_6x9(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_60" and self.selected_size == "input_710":
                        drag_images_from_folder_7x10(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_60" and self.selected_size == "input_913":
                        drag_images_from_folder_9x13(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_60" and self.selected_size == "input_1318":
                        drag_images_from_folder_13x18(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_50" and self.selected_size == "input_69":
                        self.logger.info("Không có kích thước đã chọn")
                    elif self.selected_paper == "input_50" and self.selected_size == "input_710":
                        drag_images_from50_folder_7x10(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_50" and self.selected_size == "input_1015":
                        drag_images_from50_folder_10x15(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_50" and self.selected_size == "input_913":
                        drag_images_from50_folder_9x13(self.logger, ps_app, self.url, ps_base)
                    elif self.selected_paper == "input_50" and self.selected_size == "input_1318":
                        drag_images_from50_folder_13x18(self.logger, ps_app, self.url, ps_base)
                    else:
                        self.logger.info("Không có kích thước đã chọn")
                    self.logger.info(f"Thiết kế xong.")
                    self.designBtn.emit()
                    self.designCompleted.emit("Thiết kế hoàn tất!")
                else:
                    self.logger.error("Không thể mở tài liệu trong Photoshop.")
            else:
                self.logger.error("Không thể mở ứng dụng photoshop Photoshop.")
        except Exception as e:
            self.logger.error(f"Gặp lỗi: {str(e)}")

    

class Login(QMainWindow):
    def __init__(self):
        super().__init__()          
        ui_path = self.resource_path('gui/main.ui')
        logo_path = self.resource_path('base/logo.ico')
        self.ui = uic.loadUi(ui_path, self)
        
        self.setup_logger()
        
        self.setWindowTitle('Auto Designer')
        self.setWindowIcon(QIcon(logo_path))
        
        self.ui.btn_url_1010.clicked.connect(self.select_url)
        self.ui.btn_submit.clicked.connect(self.handleSubmit)
        
    def setup_logger(self):
        logTextBox = QPlainTextEditLogger(self)
        self.activity = self.ui.activity  # Connect logTextBox to UI activity widget
        logTextBox.appendPlainText.connect(self.activity.append)
        self.logger = logging.getLogger('AppLogger')
        self.logger.addHandler(logTextBox)
        self.logger.setLevel(logging.DEBUG)
        self.logger.info("Opening the Auto Designer App")
    
    def handleSubmit(self):
        selected_size = self.get_selected_radio_button(["input_69", "input_710", "input_1015", "input_913", "input_1318"])
        selected_paper = self.get_selected_radio_button(["input_60", "input_50"])
        check_msg = self.get_selected_radio_button(["msg_no", "msg_yes"])
        
        if hasattr(self, 'url_1010'):
            self.ui.btn_submit.setEnabled(False)
            self.ui.btn_submit.setText("Đang chạy...")
            self.ui.btn_submit.setStyleSheet("background-color: #f99d20;border-radius: 10px;color: #fff;font-size: 14px;padding: 5px;max-width: 120px;")
            self.design_thread = DesignThread(self.url_1010, self.logger, selected_size, selected_paper)
            self.design_thread.designBtn.connect(self.showBtn)  
            if check_msg == "msg_yes":
                self.design_thread.designCompleted.connect(self.showCompletionMessage)  
                self.design_thread.start()
            else: 
                self.design_thread.start()
        else:
            self.logger.warning("Vui lòng chọn một thư mục trước.")
        
    def get_selected_radio_button(self, names):
        for name in names:
            radio_button = self.findChild(QRadioButton, name)
            if radio_button and radio_button.isChecked():
                return name
        return None
    
    def resource_path(self, relative_path):
        try:
            base_path = sys._MEIPASS if hasattr(sys, '_MEIPASS') else os.path.abspath('.')
        except Exception:
            base_path = os.path.abspath(".")

        return os.path.join(base_path, relative_path)
    
    def select_url(self):
        folder_path = QFileDialog.getExistingDirectory(self, "Chọn thư mục", "/")
        if folder_path:
            self.url_1010 = folder_path
            self.ui.btn_url_1010.setText(folder_path)
            self.logger.info(f"Đã chọn thư mục: {folder_path}")
            
    def showCompletionMessage(self, message):
        # Hàm hiển thị QMessageBox thông báo thiết kế hoàn tất
        msgBox = QMessageBox()
        msgBox.setIcon(QMessageBox.Icon.Information)  # Sử dụng icon thông báo
        msgBox.setText(message)
        msgBox.setWindowTitle("Thông Báo")
        msgBox.setStandardButtons(QMessageBox.StandardButton.Ok)
        msgBox.setWindowFlags(msgBox.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)  # Đặt cửa sổ luôn ở trên cùng
        msgBox.exec()
        
    def showBtn(self):
        self.ui.btn_submit.setText("Bắt đầu")
        self.ui.btn_submit.setStyleSheet("background: #0d6efd;border-radius: 10px;color: #fff;font-size: 14px;padding: 5px;max-width: 120px;")
        self.ui.btn_submit.setEnabled(True)
        
if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = Login()
    window.setWindowTitle("Auto designer")
    window.show()
    sys.exit(app.exec())
