import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pypandoc
import os
from src.converter.latex_generator import LatexGenerator

class MainWindow:
    def __init__(self, root):
        self.root = root
        self.root.title("Công cụ Chuyển đổi Word sang LaTeX - Đồ Án 2")
        self.root.geometry("1000x650") # Mở rộng cửa sổ lớn hơn
        
        self.generator = LatexGenerator()

        self.word_path = tk.StringVar()
        self.template_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self.setup_ui()

    def setup_ui(self):
        # ================= KHU VỰC CHỌN FILE (PHÍA TRÊN) =================
        top_frame = tk.Frame(self.root)
        top_frame.pack(fill="x", padx=10, pady=10)

        # 1. File Word
        tk.Label(top_frame, text="1. File Word (.docx):").grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(top_frame, textvariable=self.word_path, state='readonly', width=50).grid(row=0, column=1, padx=5, pady=5)
        tk.Button(top_frame, text="Duyệt...", command=self.load_word).grid(row=0, column=2, pady=5)

        # 2. File Template
        tk.Label(top_frame, text="2. Template (.tex):").grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(top_frame, textvariable=self.template_path, state='readonly', width=50).grid(row=1, column=1, padx=5, pady=5)
        tk.Button(top_frame, text="Duyệt...", command=self.load_template).grid(row=1, column=2, pady=5)

        # 3. Nơi lưu
        tk.Label(top_frame, text="3. Lưu kết quả tại:").grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(top_frame, textvariable=self.output_path, state='readonly', width=50).grid(row=2, column=1, padx=5, pady=5)
        tk.Button(top_frame, text="Lưu...", command=self.save_output).grid(row=2, column=2, pady=5)

        # Nút Chạy
        tk.Button(top_frame, text="CHUYỂN ĐỔI", bg="#4CAF50", fg="white", font=("Arial", 10, "bold"), 
                  command=self.process_conversion, width=15).grid(row=1, column=3, rowspan=2, padx=20)

        # ================= KHU VỰC REVIEW CHIA ĐÔI (PHÍA DƯỚI) =================
        review_frame = tk.Frame(self.root)
        review_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # Thiết lập cột chia đều 50-50
        review_frame.columnconfigure(0, weight=1)
        review_frame.columnconfigure(1, weight=1)

        # Khung Trái: Preview Word
        tk.Label(review_frame, text="📄 Xem trước nội dung Word (Văn bản thô):", font=("Arial", 10, "bold"), fg="#333").grid(row=0, column=0, sticky="w")
        self.txt_word_preview = scrolledtext.ScrolledText(review_frame, wrap=tk.WORD, width=40, height=20, bg="#f9f9f9")
        self.txt_word_preview.grid(row=1, column=0, sticky="nsew", padx=(0, 5), pady=5)

        # Khung Phải: Preview LaTeX
        tk.Label(review_frame, text="📝 Xem trước Code LaTeX sinh ra:", font=("Arial", 10, "bold"), fg="#333").grid(row=0, column=1, sticky="w")
        self.txt_latex_preview = scrolledtext.ScrolledText(review_frame, wrap=tk.WORD, width=40, height=20, bg="#2b2b2b", fg="#a9b7c6", insertbackground="white")
        self.txt_latex_preview.grid(row=1, column=1, sticky="nsew", padx=(5, 0), pady=5)


    # ================= CÁC HÀM XỬ LÝ =================
    def load_word(self):
        file = filedialog.askopenfilename(filetypes=[("Word Docs", "*.docx")])
        if file: 
            self.word_path.set(file)
            self.preview_word_content(file) # Tự động load nội dung lên khung review

    def load_template(self):
        file = filedialog.askopenfilename(filetypes=[("LaTeX Templates", "*.tex")])
        if file: self.template_path.set(file)

    def save_output(self):
        file = filedialog.asksaveasfilename(defaultextension=".tex", filetypes=[("LaTeX Files", "*.tex")])
        if file: self.output_path.set(file)

    def preview_word_content(self, filepath):
        """Dùng Pandoc dịch nhanh Word sang Plain Text để xem trước"""
        try:
            self.txt_word_preview.delete('1.0', tk.END)
            self.txt_word_preview.insert(tk.END, "Đang tải dữ liệu...")
            self.root.update()

            # Dịch sang text thô để review
            text_content = pypandoc.convert_file(filepath, to='plain')
            
            self.txt_word_preview.delete('1.0', tk.END)
            self.txt_word_preview.insert(tk.END, text_content)
        except Exception as e:
            self.txt_word_preview.delete('1.0', tk.END)
            self.txt_word_preview.insert(tk.END, f"Không thể xem trước file. Lỗi: {str(e)}")

    def process_conversion(self):
        if not self.word_path.get() or not self.output_path.get():
            messagebox.showwarning("Thiếu thông tin", "Vui lòng chọn file đầu vào và nơi lưu!")
            return

        # Làm sạch khung review LaTeX
        self.txt_latex_preview.delete('1.0', tk.END)
        self.txt_latex_preview.insert(tk.END, "Đang xử lý chuyển đổi...")
        self.root.update()

        success, msg = self.generator.convert(
            input_docx=self.word_path.get(),
            output_tex=self.output_path.get(),
            template_tex=self.template_path.get() if self.template_path.get() else None
        )

        if success:
            messagebox.showinfo("Hoàn tất", msg)
            self.load_latex_preview(self.output_path.get())
        else:
            self.txt_latex_preview.delete('1.0', tk.END)
            messagebox.showerror("Lỗi", msg)

    def load_latex_preview(self, filepath):
        """Đọc file LaTeX vừa tạo và hiển thị lên khung bên phải"""
        try:
            with open(filepath, 'r', encoding='utf-8') as f:
                content = f.read()
            self.txt_latex_preview.delete('1.0', tk.END)
            self.txt_latex_preview.insert(tk.END, content)
        except Exception as e:
            self.txt_latex_preview.delete('1.0', tk.END)
            self.txt_latex_preview.insert(tk.END, f"Không thể tải code LaTeX. Lỗi: {str(e)}")

def run_app():
    root = tk.Tk()
    app = MainWindow(root)
    root.mainloop()