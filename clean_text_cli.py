import os
import re
import unicodedata

def clean_text_v4(text: str) -> str:
    if not text:
        return ""
    
    # Bước 1: Đồng bộ Unicode - Ép toàn bộ Tổ hợp về Dựng sẵn (Chuẩn NFKC)
    # Bước này thay thế hoàn toàn cho hàm ToHopSangDungSan dài loằng ngoằng trong VBA
    text = unicodedata.normalize('NFKC', text)
    
    # Bước 2: Định nghĩa BLACKLIST (Danh sách đen các ký tự rác)
    blacklist_patterns = [
        r'[\x00-\x08\x0B\x0C\x0E\x1F\x7F]',  # Control Characters & Delete char
        r'\x09',                             # Tab character (hay dính khi OCR)
        r'[\u0080-\u009F]',                  # C1 Control block
        r'[\u200B\u200C\u200D\uFEFF]',        # Zero-width chars (ZWSP, ZWNJ, ZWJ, BOM)
        r'[\u0300-\u036F]'                   # Orphaned combining marks (Dấu tổ hợp sót lại)
    ]
    
    # Tiến hành xóa sạch rác trong danh sách đen
    for pattern in blacklist_patterns:
        text = re.sub(pattern, '', text)
    
    # Bước 3: Đồng bộ khoảng trắng Web (NBSP \u00A0) thành khoảng trắng thường
    text = text.replace('\u00A0', ' ')
    
    # Cắt khoảng trắng thừa đầu cuối và trả về kết quả
    return text.strip()

def main():
    print("="*60)
    print("  CHƯƠNG TRÌNH KHỬ NHIỄU VĂN BẢN VÀ ĐỒNG BỘ UNICODE v4")
    print("="*60)
    
    # Nhận input đường dẫn file từ người dùng
    file_path = input("[?] Nhập hoặc kéo thả file cần xử lý vào đây: ").strip().strip("'\"")
    
    if not os.path.exists(file_path):
        print("[X] Lỗi: Đường dẫn file không tồn tại!")
        return

    try:
        # Đọc file với bảng mã utf-8 để cân mọi loại tiếng Việt
        with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
            content = f.read()
        
        # Chạy pipeline khử nhiễu tinh hoa
        cleaned_content = clean_text_v4(content)
        
        # Tạo file mới có hậu tố _cleaned để tránh đè mất file gốc của người dùng
        base, ext = os.path.splitext(file_path)
        output_path = f"{base}_cleaned{ext}"
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(cleaned_content)
            
        print(f"\n[✓] THÀNH CÔNG RỰC RỠ!")
        print(f"[>] File sạch đã được lưu tại: {output_path}")
        
    except Exception as e:
        print(f"[X] Đã xảy ra lỗi hệ thống: {str(e)}")

if __name__ == "__main__":
    main()