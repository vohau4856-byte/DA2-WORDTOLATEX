import pypandoc
import re
import os
import unicodedata
import tempfile
import docx
from docx.enum.text import WD_ALIGN_PARAGRAPH
from pathlib import Path
from src.utils.constants import PANDOC_ARGS, MEDIA_DIR_NAME

class LatexGenerator:
    def __init__(self):
        try:
            pypandoc.get_pandoc_version()
        except OSError:
            pypandoc.download_pandoc()

    def pre_process_word(self, input_docx):
        """
        [BƯỚC 1] TIỀN XỬ LÝ: Tạo các đoạn văn độc lập chứa thẻ căn lề
        để "cách ly" hoàn toàn khỏi lệnh in đậm/in nghiêng của Word.
        """
        import docx
        from docx.enum.text import WD_ALIGN_PARAGRAPH
        import tempfile
        import os
        
        doc = docx.Document(input_docx)
        
        # Tạo một danh sách lưu lại các thao tác cần làm
        # (Không nên vừa lặp vừa thêm đoạn văn mới để tránh lỗi vòng lặp vô tận)
        actions = []

        for para in doc.paragraphs:
            if not para.text.strip() or not para.runs:
                continue
            
            alignment = para.alignment
            if alignment is None and para.style and para.style.paragraph_format:
                alignment = para.style.paragraph_format.alignment

            # Ghi chú lại đoạn văn nào cần thẻ gì
            if alignment == WD_ALIGN_PARAGRAPH.CENTER:
                actions.append((para, '@@GIUA@@', '@@/GIUA@@'))
            elif alignment == WD_ALIGN_PARAGRAPH.RIGHT:
                actions.append((para, '@@PHAI@@', '@@/PHAI@@'))
            elif alignment == WD_ALIGN_PARAGRAPH.LEFT:
                actions.append((para, '@@TRAI@@', '@@/TRAI@@'))

        # Thực thi việc chèn thẻ bằng cách TẠO ĐOẠN VĂN MỚI
        for para, open_tag, close_tag in actions:
            # 1. Tạo một đoạn văn trắng tinh chứa thẻ mở nằm NGAY TRÊN đoạn hiện tại
            para.insert_paragraph_before(open_tag)
            
            # 2. Thủ thuật tạo đoạn văn nằm NGAY DƯỚI đoạn hiện tại
            # (Tạo một đoạn trước, sau đó dùng lệnh XML đẩy nó xuống dưới)
            close_p = para.insert_paragraph_before(close_tag)
            para._p.addnext(close_p._p)

        temp_dir = tempfile.gettempdir()
        temp_path = os.path.join(temp_dir, "temp_doan2_processing.docx")
        doc.save(temp_path)
        return temp_path

    def post_process_latex(self, tex_path):
        """[BƯỚC 3] HẬU XỬ LÝ: Dọn dẹp tiếng Việt và phân giải thẻ"""
        with open(tex_path, 'r', encoding='utf-8') as f:
            content = f.read()

        # 1. Ép chuẩn Unicode tiếng Việt
        content = unicodedata.normalize('NFC', content)
        content = content.replace('\u00A0', ' ')
        content = content.replace('\u200B', '')

        # 2. Bơm thư viện Tiếng Việt nếu chưa có
        if r'\begin{document}' in content and r'T5' not in content:
            vn_packages = (
                "\\usepackage[T5]{fontenc}\n"
                "\\usepackage[utf8]{inputenc}\n"
                "\\usepackage[vietnamese]{babel}\n"
                "\\begin{document}"
            )
            content = content.replace(r'\begin{document}', vn_packages)

        # 3. Quét dấu chấm liên tục
        content = re.sub(r'(\.|\\[l]?dots(?:\{\})?){4,}', r' \\dotfill ', content)

        # 4. Phân giải các thẻ căn lề (đã được tự động tiêm từ bước Tiền xử lý)
        content = re.sub(r'@@GIUA@@(.*?)@@/GIUA@@', r'\\begin{center}\n\1\n\\end{center}', content, flags=re.DOTALL)
        content = re.sub(r'@@PHAI@@(.*?)@@/PHAI@@', r'\\begin{flushright}\n\1\n\\end{flushright}', content, flags=re.DOTALL)
        content = re.sub(r'@@TRAI@@(.*?)@@/TRAI@@', r'\\begin{flushleft}\n\1\n\\end{flushleft}', content, flags=re.DOTALL)

        # 5. Phân giải Tiêu đề (nếu người dùng có dùng thẻ H1, H2 thủ công)
        content = re.sub(r'@@H1@@(.*?)@@/H1@@', r'\\section{\1}', content)
        content = re.sub(r'@@H2@@(.*?)@@/H2@@', r'\\subsection{\1}', content)

        with open(tex_path, 'w', encoding='utf-8') as f:
            f.write(content)

    def convert(self, input_docx, output_tex, template_tex=None):
        """[BƯỚC 2] LÕI XỬ LÝ CHÍNH"""
        try:
            input_path = Path(input_docx).absolute()
            output_path = Path(output_tex).absolute()
            
            # --- GỌI HÀM TIỀN XỬ LÝ TRƯỚC ---
            temp_docx = self.pre_process_word(str(input_path))

            original_cwd = os.getcwd()
            os.chdir(output_path.parent)

            media_rel_dir = MEDIA_DIR_NAME
            Path(media_rel_dir).mkdir(parents=True, exist_ok=True)

            args = PANDOC_ARGS.copy()
            args.append(f'--extract-media={media_rel_dir}')
            
            if template_tex and Path(template_tex).is_file():
                args.append(f'--template={Path(template_tex).absolute()}')

            # Đưa file TẠM (đã được tiêm code) vào Pandoc thay vì file gốc
            pypandoc.convert_file(
                source_file=temp_docx,
                to='latex',
                outputfile=output_path.name,
                extra_args=args
            )
            
            os.chdir(original_cwd)
            
            # Xóa bỏ file tạm sau khi dùng xong cho sạch máy
            if os.path.exists(temp_docx):
                os.remove(temp_docx)

            # --- GỌI HÀM HẬU XỬ LÝ ---
            self.post_process_latex(output_path)

            return True, f"Thành công!\nFile lưu tại: {output_path.absolute()}"
        
        except Exception as e:
            if 'original_cwd' in locals():
                os.chdir(original_cwd)
            return False, f"Lỗi kỹ thuật: {str(e)}"