# 📄 Word to LaTeX Converter (Đồ án Cơ sở 2)

> **Phần mềm chuyển đổi tự động tập tin Microsoft Word (.docx) sang chuẩn mã lệnh LaTeX (.tex)**
> **Khoa Công nghệ Thông tin - Trường Đại học Nam Cần Thơ** (Khóa K10 - 03/2026)

---

## 📌 Giới thiệu đề tài
Đề tài **"CHUYỂN ĐỔI TỰ ĐỘNG TẬP TIN WORD HOÀN THIỆN VỚI MẪU MS WORD (MS WORD TEMPLATE) SANG MẪU LATEX (LATEX TEMPLATE)"** được xây dựng nhằm giải quyết rào cản về cú pháp lệnh phức tạp khi soạn thảo tài liệu học thuật bằng LaTeX. Phần mềm ứng dụng máy tính (Desktop Application) này giúp tự động hóa quy trình chuyển đổi tài liệu từ `.docx` sang `.tex`, bảo toàn cấu trúc văn bản, định dạng căn lề, công thức toán học, trích xuất hình ảnh và giải quyết triệt để lỗi phông chữ Tiếng Việt.

---

## 👨‍💻 Thông tin thực hiện
* **Sinh viên thực hiện:** Võ Trung Hậu
* **Mã số sinh viên:** 225635
* **Ngành học:** Công nghệ Thông tin (Mã ngành: `7480201`)
* **Giảng viên hướng dẫn:** TS. Ngô Hồ Anh Khôi
* **Đơn vị:** Khoa Công nghệ Thông tin – Trường Đại học Nam Cần Thơ
* **Thời gian thực hiện:** Tháng 03/2026

---

## 🛠️ Công nghệ & Thư viện sử dụng
* **Ngôn ngữ lập trình:** Python
* **Giao diện người dùng (GUI):** Thư viện Tkinter 
* **Công cụ chuyển đổi lõi:** Pandoc (`pypandoc >= 1.13`)
* **Xử lý tài liệu Word:** `python-docx`
* **Thuật toán xử lý:** Biểu thức chính quy (Regex) & Chuẩn hóa Unicode NFC
* **Đóng gói ứng dụng:** PyInstaller (`Word2Latex_DoAn2.spec`)

---

## 🌟 Các tính năng chính của phần mềm

- 🖥️ **Giao diện Dual-Preview song song:** Cho phép xem trước nội dung văn bản Word thô và mã nguồn LaTeX sinh ra cùng lúc.
- 📐 **Bảo toàn căn lề & Định dạng:** Sử dụng thuật toán tiêm thẻ quy ước (`@@GIUA@@`, `@@PHAI@@`) ở giai đoạn Tiền xử lý để giữ nguyên căn lề văn bản.
- 🧮 **Chuyển đổi công thức Toán học:** Ánh xạ tự động các công thức từ chuẩn OMML của MS Word sang cú pháp Math Mode (`$ ... $`) chuẩn của LaTeX.
- 🖼️ **Trích xuất & Quản lý hình ảnh:** Tự động bóc tách hình ảnh nhúng trong file Word và lưu trữ khoa học vào thư mục `extracted_media`.
- 🇻🇳 **Chuẩn hóa Tiếng Việt (Unicode NFC):** Khắc phục lỗi phông chữ Tiếng Việt bằng cách ép chuẩn NFC và cấu hình tự động các gói lệnh `vietnam`, `utf8`.
- 📑 **Tùy biến Template:** Cho phép nạp thêm tệp mẫu `.tex` để định dạng mã nguồn đầu ra theo chuẩn của nhà trường hoặc các tạp chí.

---

## ⚙️ Quy trình xử lý Pipeline 3 giai đoạn
1. **Giai đoạn 1 - Tiền xử lý (Pre-processing):** Can thiệp vào cấu trúc XML của file Word bằng `python-docx` để nhận diện vị trí căn lề và tiêm các thẻ quy ước nhằm tránh tràn định dạng.
2. **Giai đoạn 2 - Chuyển đổi lõi (Core Conversion):** Gọi công cụ Pandoc bóc tách văn bản theo Cây cú pháp trừu tượng (AST) và tự động xuất hình ảnh ra thư mục `extracted_media`.
3. **Giai đoạn 3 - Hậu xử lý (Post-processing):** Dùng Regex chuyển thẻ quy ước thành lệnh LaTeX tương ứng, ép chuẩn Unicode NFC cho Tiếng Việt và tự động khai báo gói thư viện.

---

## 📁 Cấu trúc thư mục mã nguồn

```text
Word2Latex_DoAn2/
├── main.py                  # File khởi chạy chính của ứng dụng
├── setup.py                 # File cấu hình cài đặt package 'word-to-latex-converter'
├── requirements.txt         # Danh sách các thư viện Python phụ thuộc
├── Word2Latex_DoAn2.spec    # File cấu hình đóng gói PyInstaller ra file .exe
└── src/                     # Mã nguồn ứng dụng
    ├── gui/                 # Module giao diện người dùng (Tkinter - MainWindow)
    └── converter/           # Module xử lý thuật toán (Pre-process, Pandoc, Post-process)
``

---
