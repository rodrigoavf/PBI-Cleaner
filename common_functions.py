import os

def count_files(root_dir):
    total = 0
    for _, _, files in os.walk(root_dir):
        total += len(files)
    return total