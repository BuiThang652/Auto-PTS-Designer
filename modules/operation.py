import time
import os
import win32com.client
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# Kiểm tra nếu Photoshop đã mở
def is_photoshop_open():
    try:
        ps_app = win32com.client.GetActiveObject("Photoshop.Application")
        return True, ps_app
    except:
        return False, None

# Mở ứng dụng photoshop nếu chưa mở
def open_photoshop(logger):
    try:
        already_open, ps_app = is_photoshop_open()
        if already_open:
            logger.info("Photoshop đã được mở trước đó.")
            return ps_app
        else:
            ps_app = win32com.client.Dispatch("Photoshop.Application")
            return ps_app
    except Exception as e:
        logger.error("Không thể mở Photoshop:", e)
        return None

# Tắt ứng dụng photoshop với delay nhất định
def close_photoshop_after_delay(ps_app, delay):
    try:
        # print(f"Đợi {delay} giây trước khi đóng Photoshop...")
        time.sleep(delay)
        ps_app.Quit()
        # print("Photoshop đã được đóng.")
    except Exception as e:
        print("Không thể đóng Photoshop:", e)

# Mở một file Photoshop
def open_existing_document(ps_app, file_path):
    if ps_app is None:
        # print("Ứng dụng Photoshop chưa được khởi tạo.")
        return None
    try:
        already_open, ps_doc = is_document_open(ps_app, file_path)
        if already_open:
            # print(f"Tài liệu '{file_path}' đã được mở trước đó.")
            return ps_doc
        else:
            ps_doc = ps_app.Open(file_path)
            # print(f"Tài liệu '{file_path}' đã được mở trong Photoshop.")
            return ps_doc
    except Exception as e:
        # print("Không thể mở tài liệu:", e)
        return None


# Mở toàn bộ file trong một folder
def open_all_documents_in_folder(ps_app, folder_path):
    try:
        # Kiểm tra xem đường dẫn là một thư mục
        if not os.path.isdir(folder_path):
            # print(f"{folder_path} không phải là một thư mục.")
            return

        # Lặp qua tất cả các tệp trong thư mục
        for filename in os.listdir(folder_path):
            # Xác định đường dẫn đầy đủ của tệp
            file_path = os.path.join(folder_path, filename)
            
            # Mở tệp nếu là một tệp hợp lệ
            if os.path.isfile(file_path):
                open_existing_document(ps_app, file_path)
    except Exception as e:
        # print("Đã xảy ra lỗi khi mở các tệp trong thư mục:", e)
        pass

# Kiểm tra nếu tài liệu đã được mở trong window pts chưa (Window pts, Đường dẫn)
def is_document_open(ps_app, file_path):
    try:
        docs = ps_app.Documents
        for doc in docs:
            if doc.FullName == file_path:
                return True, doc
        return False, None
    except Exception as e:
        # print("Lỗi khi kiểm tra tài liệu:", e)
        return False, None

def drag_image_to_position(ps_app, src_doc, dst_doc, x, y):
    try:
        # Kích hoạt tài liệu nguồn và sao chép nội dung
        ps_app.ActiveDocument = src_doc
        src_doc.ActiveLayer.Copy()
        
        # Kích hoạt tài liệu đích và dán nội dung
        ps_app.ActiveDocument = dst_doc
        new_layer = dst_doc.Paste()
        
        # Di chuyển layer mới đến vị trí mong muốn
        js_code = f"""
        var x = {x};
        var y = {y};
        var doc = app.activeDocument;
        var layer = doc.activeLayer;
        layer.translate(x - layer.bounds[0], y - layer.bounds[1]);
        """
        ps_app.DoJavaScript(js_code)
        
        # print("Hình ảnh đã được chuyển và đặt tại vị trí", x, y, "từ", src_doc.FullName, "sang", dst_doc.FullName)
    except Exception as e:
        print("Không thể chuyển ảnh:", e)

def drag_image_to_fit(ps_app, src_doc, dst_doc, target_width=14.41, target_height=9.41):
    try:
        # Kích hoạt tài liệu nguồn và sao chép nội dung
        ps_app.ActiveDocument = src_doc
        src_doc.ActiveLayer.Copy()

        # Kích hoạt tài liệu đích và dán nội dung
        ps_app.ActiveDocument = dst_doc
        new_layer = dst_doc.Paste()

        # Tính toán tỷ lệ thu nhỏ/phóng to dựa trên tỷ lệ chiều rộng/chiều cao cụ thể và căn giữa hình ảnh
        js_code = """
        var targetWidth = """ + str(target_width) + """;
        var targetHeight = """ + str(target_height) + """;
        var doc = app.activeDocument;
        var layer = doc.activeLayer;
        var docWidth = doc.width;
        var docHeight = doc.height;
        var originalWidth = layer.bounds[2] - layer.bounds[0];
        var originalHeight = layer.bounds[3] - layer.bounds[1];
        var originalRatio = originalWidth / originalHeight;
        var targetRatio = targetWidth / targetHeight;
        var scale;
        if (originalRatio < targetRatio) {
            scale = (targetHeight / originalHeight) * 100;
        } else {
            scale = (targetWidth / originalWidth) * 100;
        }
        layer.resize(scale, scale, AnchorPosition.TOPLEFT);
        var centerX = (docWidth - originalWidth * scale / 100) / 2;
        var centerY = (docHeight - originalHeight * scale / 100) / 2;
        layer.translate(centerX - layer.bounds[0], centerY - layer.bounds[1]);
        """
        ps_app.DoJavaScript(js_code)
        
        dst_doc.Flatten()

        # print("Hình ảnh đã được điều chỉnh kích thước và căn giữa trong toàn bộ tài liệu mà không bị méo.")
    except Exception as e:
        print("Không thể điều chỉnh kích thước và căn giữa hình ảnh:", e)


# Thiết kế (10x10) - Chưa check kích thước kéo vào.
def drag_images_from_folder_10x10(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 15, 30, 45]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "10x10.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)

    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 10 

        if check_10x10(src_doc):
            set_image_size(src_doc, 15, 10)
            time.sleep(0.2)
            drag_image_to_position(ps_app, src_doc, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
        else:
            dst_doc_10x10 = open_existing_document(ps_app, base_image_path)
            if src_doc and dst_doc_10x10:
                drag_image_to_fit(ps_app, src_doc, dst_doc_10x10)
                time.sleep(0.2)
                drag_image_to_position(ps_app, dst_doc_10x10, dst_doc, current_x, current_y)
                time.sleep(0.2)
                src_doc.Close(2)
                time.sleep(0.2)
                dst_doc_10x10.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")
        
def drag_images_from50_folder_10x15(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 10, 20, 30, 40]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "15x10.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)

    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 15 

        if check_15x10(src_doc):
            set_image_size(src_doc, 10, 15)
            time.sleep(0.2)
            drag_image_to_position(ps_app, src_doc, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
        else:
            dst_doc_10x10 = open_existing_document(ps_app, base_image_path)
            if src_doc and dst_doc_10x10:
                drag_image_to_fit(ps_app, src_doc, dst_doc_10x10, 9.41, 14.41)
                time.sleep(0.2)
                drag_image_to_position(ps_app, dst_doc_10x10, dst_doc, current_x, current_y)
                time.sleep(0.2)
                src_doc.Close(2)
                time.sleep(0.2)
                dst_doc_10x10.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")

def drag_images_from_folder_7x10(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 10, 20, 30, 40, 50]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "7x10.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)
    
    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 7  # Tăng y lên 7 sau mỗi 6 hình ảnh
        dst_doc_7x10 = open_existing_document(ps_app, base_image_path)
        
        if src_doc and dst_doc_7x10:
            drag_image_to_fit(ps_app, src_doc, dst_doc_7x10, 9.41, 6.41)
            time.sleep(0.2)
            drag_image_to_position(ps_app, dst_doc_7x10, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
            time.sleep(0.2)
            dst_doc_7x10.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")
        
def drag_images_from50_folder_7x10(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 10, 20, 30, 40]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "7x10.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)
    
    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 7  # Tăng y lên 7 sau mỗi 6 hình ảnh
        dst_doc_7x10 = open_existing_document(ps_app, base_image_path)
        
        if src_doc and dst_doc_7x10:
            drag_image_to_fit(ps_app, src_doc, dst_doc_7x10, 9.41, 6.41)
            time.sleep(0.2)
            drag_image_to_position(ps_app, dst_doc_7x10, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
            time.sleep(0.2)
            dst_doc_7x10.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")

def drag_images_from_folder_6x9(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 6, 12, 18, 24, 30, 36, 42, 48, 54]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "6x9.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)
    
    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 9
        dst_doc_6x9 = open_existing_document(ps_app, base_image_path)
        
        if src_doc and dst_doc_6x9:
            drag_image_to_fit(ps_app, src_doc, dst_doc_6x9, 5.40, 8.41)
            time.sleep(0.2)
            drag_image_to_position(ps_app, dst_doc_6x9, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
            time.sleep(0.2)
            dst_doc_6x9.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")

def drag_images_from_folder_13x18(logger, ps_app, folder_path, dst_doc):
    x_positions = [10, 22.5, 35, 47.5]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "13x18.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)

    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 17.6  # Tăng y lên 10 sau mỗi 4 hình ảnh

        if check_13x18(src_doc):
            set_image_size(src_doc, 12.5, 17.6)
            time.sleep(0.2)
            drag_image_to_position(ps_app, src_doc, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
        else:
            dst_doc_13x18 = open_existing_document(ps_app, base_image_path)
            if src_doc and dst_doc_13x18:
                drag_image_to_fit(ps_app, src_doc, dst_doc_13x18, 11.91, 17)
                time.sleep(0.2)
                drag_image_to_position(ps_app, dst_doc_13x18, dst_doc, current_x, current_y)
                time.sleep(0.2)
                src_doc.Close(2)
                time.sleep(0.2)
                dst_doc_13x18.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")
        
def drag_images_from50_folder_13x18(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 12.5, 25, 37.5]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "13x18.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)

    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 17.6  # Tăng y lên 10 sau mỗi 4 hình ảnh

        if check_13x18(src_doc):
            set_image_size(src_doc, 12.5, 17.6)
            time.sleep(0.2)
            drag_image_to_position(ps_app, src_doc, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
        else:
            dst_doc_13x18 = open_existing_document(ps_app, base_image_path)
            if src_doc and dst_doc_13x18:
                drag_image_to_fit(ps_app, src_doc, dst_doc_13x18, 11.91, 17)
                time.sleep(0.2)
                drag_image_to_position(ps_app, dst_doc_13x18, dst_doc, current_x, current_y)
                time.sleep(0.2)
                src_doc.Close(2)
                time.sleep(0.2)
                dst_doc_13x18.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")

def drag_images_from_folder_9x13(logger, ps_app, folder_path, dst_doc):
    x_positions = [10, 22.5, 35, 47.5]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "9x13.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)

    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 8.8  # Tăng y lên 10 sau mỗi 4 hình ảnh

        if check_9x13(src_doc):
            set_image_size(src_doc, 12.5, 8.8)
            time.sleep(0.2)
            drag_image_to_position(ps_app, src_doc, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
        else:
            dst_doc_9x13 = open_existing_document(ps_app, base_image_path)
            if src_doc and dst_doc_9x13:
                drag_image_to_fit(ps_app, src_doc, dst_doc_9x13, 11.9, 8.21)
                time.sleep(0.2)
                drag_image_to_position(ps_app, dst_doc_9x13, dst_doc, current_x, current_y)
                time.sleep(0.2)
                src_doc.Close(2)
                time.sleep(0.2)
                dst_doc_9x13.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")
        
def drag_images_from50_folder_9x13(logger, ps_app, folder_path, dst_doc):
    x_positions = [0, 12.5, 25, 37.5]  # Các vị trí x cố định
    y_position = 0  # Y bắt đầu từ 0
    # Lấy đường dẫn của thư mục hiện tại của script
    script_dir = os.path.dirname(os.path.abspath(__file__))  # Đường dẫn đến thư mục hiện tại của file
    base_dir = os.path.dirname(script_dir)  # Lùi lại một cấp thư mục để ra khỏi thư mục 'modules'
    # Tạo đường dẫn tương đối đến thư mục 'base'
    base_image_path = os.path.join(base_dir, "base", "9x13.jpg")
    # Kết hợp đường dẫn tương đối với thư mục hiện tại của script
    files = os.listdir(folder_path)

    for i, filename in enumerate(files):
        file_path = os.path.join(folder_path, filename)
        src_doc = open_existing_document(ps_app, file_path)
        # Tính toán vị trí x và y hiện tại trước khi kiểm tra điều kiện
        current_x = x_positions[i % len(x_positions)]
        current_y = y_position + (i // len(x_positions)) * 8.8  # Tăng y lên 10 sau mỗi 4 hình ảnh

        if check_9x13(src_doc):
            set_image_size(src_doc, 12.5, 8.8)
            time.sleep(0.2)
            drag_image_to_position(ps_app, src_doc, dst_doc, current_x, current_y)
            time.sleep(0.2)
            src_doc.Close(2)
        else:
            dst_doc_9x13 = open_existing_document(ps_app, base_image_path)
            if src_doc and dst_doc_9x13:
                drag_image_to_fit(ps_app, src_doc, dst_doc_9x13, 11.9, 8.21)
                time.sleep(0.2)
                drag_image_to_position(ps_app, dst_doc_9x13, dst_doc, current_x, current_y)
                time.sleep(0.2)
                src_doc.Close(2)
                time.sleep(0.2)
                dst_doc_9x13.Close(2)
        time.sleep(0.2)
                
        logger.info(f"Thiết kế ảnh thứ {i+1}: {file_path}")

def check_10x10(doc):
    try:
        width = doc.Width
        height = doc.Height
        resolution = doc.Resolution
        
        # Tính tỷ lệ giữa chiều rộng và chiều cao
        aspect_ratio = width / height
        
        # Kiểm tra tỷ lệ này có xấp xỉ 1.5 hay không (ví dụ: chấp nhận sai số 1%)
        if abs(aspect_ratio - 1.5) / 1.5 < 0.01 and resolution == 300:
            # print(f"Tài liệu '{doc.FullName}' có tỷ lệ xấp xỉ 15:10 và độ phân giải 300 dpi.")
            return True
        else:
            # print(f"Tài liệu '{doc.FullName}' không đạt tỷ lệ 15:10 yêu cầu hoặc sai độ phân giải.")
            return False
    except Exception as e:
        # print("Lỗi khi truy cập thông tin tài liệu:", e)
        return False
   
def check_15x10(doc):
    try:
        width = doc.Width
        height = doc.Height
        resolution = doc.Resolution
        
        # Tính tỷ lệ giữa chiều rộng và chiều cao
        aspect_ratio = height / width
        
        # Kiểm tra tỷ lệ này có xấp xỉ 1.5 hay không (ví dụ: chấp nhận sai số 1%)
        if abs(aspect_ratio - 1.5) / 1.5 < 0.01 and resolution == 300:
            # print(f"Tài liệu '{doc.FullName}' có tỷ lệ xấp xỉ 15:10 và độ phân giải 300 dpi.")
            return True
        else:
            # print(f"Tài liệu '{doc.FullName}' không đạt tỷ lệ 15:10 yêu cầu hoặc sai độ phân giải.")
            return False
    except Exception as e:
        # print("Lỗi khi truy cập thông tin tài liệu:", e)
        return False   

def check_13x18(doc):
    try:
        width = doc.Width
        height = doc.Height
        resolution = doc.Resolution
        
        # Tính tỷ lệ giữa chiều rộng và chiều cao
        aspect_ratio = height / width
        
        # Kiểm tra tỷ lệ này có xấp xỉ 1.5 hay không (ví dụ: chấp nhận sai số 1%)
        if abs(aspect_ratio - 1.408) / 1.408 < 0.01 and resolution == 300:
            # print(f"Tài liệu '{doc.FullName}' có tỷ lệ xấp xỉ 15:10 và độ phân giải 300 dpi.")
            return True
        else:
            # print(f"Tài liệu '{doc.FullName}' không đạt tỷ lệ 15:10 yêu cầu hoặc sai độ phân giải.")
            return False
    except Exception as e:
        # print("Lỗi khi truy cập thông tin tài liệu:", e)
        return False
    
def check_9x13(doc):
    try:
        width = doc.Width
        height = doc.Height
        resolution = doc.Resolution
        
        # Tính tỷ lệ giữa chiều rộng và chiều cao
        aspect_ratio = width / height
        
        # Kiểm tra tỷ lệ này có xấp xỉ 1.5 hay không (ví dụ: chấp nhận sai số 1%)
        if abs(aspect_ratio - 1.42) / 1.42 < 0.01 and resolution == 300:
            # print(f"Tài liệu '{doc.FullName}' có tỷ lệ xấp xỉ 15:10 và độ phân giải 300 dpi.")
            return True
        else:
            # print(f"Tài liệu '{doc.FullName}' không đạt tỷ lệ 15:10 yêu cầu hoặc sai độ phân giải.")
            return False
    except Exception as e:
        # print("Lỗi khi truy cập thông tin tài liệu:", e)
        return False

def set_image_size(doc, width, height):
    try:
        # Đảm bảo có tài liệu đang mở
        if doc is None:
            print("Tài liệu không tồn tại.")
            return

        # Thiết lập đơn vị đo lường trong Photoshop thành centimeters
        doc.Application.Preferences.RulerUnits = 3

        # Thay đổi kích thước và độ phân giải của hình ảnh
        doc.ResizeImage(Width=width, Height=height, Resolution=300)
        # print(f"Hình ảnh đã được thay đổi kích thước thành {width}x{height} cm và độ phân giải là {300} dpi.")
    except Exception as e:
        print(f"Lỗi khi thay đổi kích thước và độ phân giải của hình ảnh: {str(e)}")
        
        
        