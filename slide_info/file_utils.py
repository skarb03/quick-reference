from pathlib import Path

def list_ppt_files(directory_path):
    ppt_files = [str(file) for file in Path(directory_path).rglob('*.pptx') if file.is_file()]
    return sorted(ppt_files)
