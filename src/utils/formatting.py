# Chứa các hàm hỗ trợ định dạng (Tùy chọn mở rộng sau này)
def clean_file_path(path_string):
    """Xóa khoảng trắng thừa ở hai đầu đường dẫn."""
    return path_string.strip() if path_string else ""