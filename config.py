import configparser
import os
import sys

def read_config(filename='config.ini'):
    config = configparser.ConfigParser()

    # 1. 우선: 현재 실행 파일과 같은 경로에서 config.ini 찾기
    exe_dir = os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__)
    external_path = os.path.join(exe_dir, filename)

    if os.path.exists(external_path):
        config.read(external_path, encoding='utf-8')
    else:
        # 2. fallback: PyInstaller 내에 포함된 config.ini
        if getattr(sys, 'frozen', False):
            base_path = sys._MEIPASS
            internal_path = os.path.join(base_path, filename)
            config.read(internal_path, encoding='utf-8')
        else:
            raise FileNotFoundError(f"{filename} 파일을 찾을 수 없습니다.")

    return config['DEFAULT']
