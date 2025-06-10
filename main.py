import os
import subprocess
import logging
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox

from config import read_config
from slide_info.ppt_extractor import extract_slide_data
from slide_info.file_utils import list_ppt_files
from slide_info.data_processor import merge_data, save_to_excel
from rfp_group.rfp_processor import process_excel_file

# 로그 설정
logging.basicConfig(
    filename="log.txt",
    level=logging.ERROR,
    format="%(asctime)s [%(levelname)s] %(message)s",
    encoding="utf-8"
)

def run_slide_info(ppt_folder_path, save_path, config):
    ppt_files = list_ppt_files(ppt_folder_path)
    extract_config = {
        'Title': config['Title'],
        'Rfp': config['Rfp'],
        'PageNo': config['PageNo'],
        'Navigation': config['Navigation']
    }
    ppts_data = [extract_slide_data(ppt_file, extract_config) for ppt_file in ppt_files]
    merged_data = merge_data(ppts_data)
    save_to_excel(merged_data, save_path)

def run_rfp_group(input_file, output_file):
    process_excel_file(input_file, output_file)

def show_message(title, message, is_error=False):
    root = tk.Tk()
    root.withdraw()
    if is_error:
        messagebox.showerror(title, message)
    else:
        messagebox.showinfo(title, message)
    root.destroy()

def choose_task_window():
    window = tk.Tk()
    window.title("작업 선택")
    window.geometry("300x230")
    window.eval('tk::PlaceWindow . center')

    label = tk.Label(window, text="실행할 작업을 선택하세요", font=("맑은 고딕", 12))
    label.pack(pady=15)

    choice = {"value": None}

    def select_slide():
        choice["value"] = "slide"
        window.destroy()

    def select_rfp():
        choice["value"] = "rfp"
        window.destroy()

    def open_config():
        try:
            subprocess.Popen(["notepad.exe", "config.ini"])
        except Exception as e:
            show_message("오류", f"config.ini 열기 실패:\n{e}", is_error=True)

    tk.Button(window, text="📂 Slide Info", width=22, command=select_slide).pack(pady=5)
    tk.Button(window, text="📄 RFP Group", width=22, command=select_rfp).pack(pady=5)
    tk.Button(window, text="⚙ 설정(config.ini) 수정", width=22, command=open_config).pack(pady=15)

    window.mainloop()
    return choice["value"]

def main():
    while True:
        task = choose_task_window()
        if not task:
            break

        config = read_config()

        if task == "slide":
            ppt_folder = filedialog.askdirectory(
                title="PPT 폴더 선택", initialdir=config.get('SlideInfoFolder', '.')
            )
            if not ppt_folder:
                continue

            save_path = filedialog.asksaveasfilename(
                title="슬라이드 정보 저장 경로 선택",
                defaultextension=".xlsx",
                filetypes=[("Excel 파일", "*.xlsx")],
                initialfile="제안서_슬라이드_정보.xlsx"
            )
            if not save_path:
                continue

            try:
                run_slide_info(ppt_folder, save_path, config)
                show_message("완료", f"슬라이드 정보가 저장되었습니다:\n{save_path}")
            except Exception:
                logging.error(traceback.format_exc())
                show_message("오류", "Slide Info 처리 중 오류가 발생했습니다.\nlog.txt를 확인해주세요.", is_error=True)

        elif task == "rfp":
            input_file = filedialog.askopenfilename(
                title="슬라이드 정보 엑셀 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx")],
                initialdir=os.path.dirname(config.get('RfpInputFile', '.'))
            )
            if not input_file:
                continue

            output_file = filedialog.asksaveasfilename(
                title="RFP 정보 저장 경로 선택",
                defaultextension=".xlsx",
                filetypes=[("Excel 파일", "*.xlsx")],
                initialfile="제안서_슬라이드_정보(rfp).xlsx"
            )
            if not output_file:
                continue
            try:
                run_rfp_group(input_file, output_file)
            except Exception as e:
                with open("log.txt", "a", encoding="utf-8") as f:
                    f.write("RFP 처리 에러:\n")
                    f.write(traceback.format_exc())
                print("❌ RFP 처리 실패. log.txt를 확인하세요.")
                return

if __name__ == "__main__":
    main()
