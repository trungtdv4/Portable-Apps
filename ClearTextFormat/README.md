# Text Cleaner — Khử nhiễu Unicode

Công cụ dọn dẹp văn bản bị nhiễu Unicode — đặc biệt hữu ích cho file từ OCR, copy-paste từ web, hoặc trao đổi giữa các hệ thống khác nhau.

**Xử lý được:**
- Ký tự tổ hợp (Unicode NFD) → dựng sẵn (NFC) → VLOOKUP nhận ra nhau
- Ký tự ẩn: Zero-width space, BOM, control chars
- Khoảng trắng rác: NBSP, tab thừa, space thừa

**Hỗ trợ:** `.xlsx` `.csv` `.txt` `.md` `.docx`

---

## Cài đặt (làm 1 lần)

Yêu cầu: **Python 3.8+**  
Kiểm tra: mở terminal/PowerShell, gõ `python --version`

### Windows (PowerShell)
```powershell
# 1. Tải về
Invoke-WebRequest https://github.com/YOUR_USERNAME/text-cleaner/archive/refs/heads/main.zip -OutFile cleaner.zip
Expand-Archive cleaner.zip
cd text-cleaner-main

# 2. Cài thư viện
pip install -r requirements.txt
```

### macOS / Linux (Terminal)
```bash
# 1. Tải về
curl -L https://github.com/YOUR_USERNAME/text-cleaner/archive/refs/heads/main.zip -o cleaner.zip
unzip cleaner.zip && cd text-cleaner-main

# 2. Cài thư viện
pip install -r requirements.txt
```

---

## Sử dụng

### Cách 1 — Kéo thả file vào terminal (dễ nhất)
```bash
python cleaner.py
# → Chương trình hỏi đường dẫn, kéo thả file vào là xong
```

### Cách 2 — Truyền đường dẫn trực tiếp
```bash
python cleaner.py "C:\Users\ten\Documents\data.xlsx"
python cleaner.py "/Users/ten/Documents/data.xlsx"
```

Chương trình sẽ hỏi bạn muốn:
- Tạo file mới `data_cleaned.xlsx` *(khuyên dùng)*
- Ghi đè file gốc *(tự động tạo file backup `.bak`)*

---

## Ví dụ

```
╔══════════════════════════════════════════╗
║        TEXT CLEANER  —  Unicode Fixer    ║
║   Hỗ trợ: .xlsx  .csv  .txt  .md  .docx ║
╚══════════════════════════════════════════╝

Nhập đường dẫn file cần xử lý: data.xlsx

   File      : data.xlsx
   Loại      : .xlsx
   Kích thước: 24.3 KB

Lưu kết quả:
  1  Tạo file mới  (data_cleaned.xlsx)
  2  Ghi đè file gốc
Chọn (1/2) [mặc định: 1]:

   Đang xử lý...

✔  Hoàn thành!
   Đã quét  : 1,240 ô
   Có nhiễu : 18 ô
   File mới  : data_cleaned.xlsx
```

---

## Câu hỏi thường gặp

**VLOOKUP vẫn không nhận ra nhau sau khi clean?**  
Chạy clean cho *cả hai* file nguồn và đích — nhiễu có thể ở bất kỳ đâu.

**Lỗi `pip: command not found`?**  
Thử `pip3` thay vì `pip`.

**File sau khi clean bị lỗi font?**  
Không xảy ra — chương trình chỉ xóa ký tự ẩn, không chạm font hay định dạng.
