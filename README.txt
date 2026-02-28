HƯỚNG DẪN CÀI ĐẶT VÀ SỬ DỤNG TOOL QUẢN LÝ HÓA ĐƠN
=====================================================

1. Yêu cầu hệ thống:
   - Hệ điều hành: Windows 10/11
   - Python: Phiên bản 3.9 trở lên (https://www.python.org/downloads/)
   - Tesseract OCR: Để nhận diện file PDF scan (ảnh).

2. Cài đặt Tesseract OCR:
   - Tải bộ cài tại: https://github.com/UB-Mannheim/tesseract/wiki
   - Chạy file cài đặt.
   - QUAN TRỌNG: Ghi nhớ đường dẫn cài đặt (Mặc định là "C:\Program Files\Tesseract-OCR"). 
   - Tool sẽ tự động tìm trong thư mục mặc định này.

3. Cài đặt thư viện:
   - Mở cmd hoặc terminal tại thư mục chứa code này.
   - Chạy lệnh: 
     pip install -r requirements.txt

4. Chạy chương trình:
   - Mở cmd tại thư mục này.
   - Chạy lệnh:
     streamlit run app.py
   
   - Trình duyệt sẽ tự động mở trang web quản lý hóa đơn.

5. Lưu ý:
   - Thư mục 'poppler-24.08.0' đi kèm là cần thiết để xử lý file PDF. Đừng xóa nó.
   - File 'hoadon_tonghop.xlsx' dùng để định dạng file xuất ra.

Chúc bạn thành công!
