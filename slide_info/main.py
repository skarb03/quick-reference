import os
from config import read_config
from ppt_extractor import extract_slide_data
from file_utils import list_ppt_files
from data_processor import merge_data, save_to_excel

def main():
    config = read_config()
    directory_path = config['DirectoryPath']
    excel_file_path = os.path.join(directory_path, '제안서_슬라이드_정보.xlsx')
    
    ppt_files = list_ppt_files(directory_path)
    ppts_data = [extract_slide_data(ppt_file, config) for ppt_file in ppt_files]
    
    merged_data = merge_data(ppts_data)
    save_to_excel(merged_data, excel_file_path)

if __name__ == '__main__':
    main()
