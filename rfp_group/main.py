import os
from rfp_processor import process_excel_file
from config import read_config


def main():
    config = read_config()
    directory_path = config['DirectoryPath']
    input_path = os.path.join(directory_path, '제안서_슬라이드_정보.xlsx')
    output_path = os.path.join(directory_path, '제안서_슬라이드_정보(rfp).xlsx')
    process_excel_file(input_path, output_path)

if __name__ == '__main__':
    main()
