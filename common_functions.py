import os
from PyQt6.QtGui import QFont

def count_files(root_dir):
    total = 0
    for _, _, files in os.walk(root_dir):
        total += len(files)
    return total
def code_editor_font(f_type="Consolas", f_size=10):
    return QFont(f_type, f_size)