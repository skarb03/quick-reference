import pandas as pd
import re

def expand_range(prefix, start, end):
    """주어진 접두사와 시작, 끝 번호로 범위를 확장하여 리스트를 반환합니다."""
    start_num = int(start)
    end_num = int(end)
    return [f"{prefix}{str(num).zfill(len(start))}" for num in range(start_num, end_num + 1)]

def parse_and_expand(input_str):
    """콤마로 구분된 문자열을 받아 확장된 리스트를 반환합니다."""
    parts = input_str.split(',')
    expanded_list = []

    for part in parts:
        if '~' in part:
            match = re.match(r"([a-zA-Z]+-)(\d+)~(\d+)", part)
            if match:
                prefix, start, end = match.groups()
                expanded_list.extend(expand_range(prefix, start, end))
        else:
            match = re.match(r"([a-zA-Z]+-)?(\d+)", part)
            if match:
                prefix, number = match.groups()
                if prefix is None:
                    # 이전 접두사를 사용하여 번호를 확장합니다.
                    prev_prefix, prev_number = re.match(r"([a-zA-Z]+-)(\d+)", expanded_list[-1]).groups()
                    expanded_list.append(f"{prev_prefix}{number.zfill(len(prev_number))}")
                else:
                    expanded_list.append(f"{prefix}{number}")

    return expanded_list

def process_excel_file(input_path, output_path):
    """엑셀 파일을 읽고 RFP 데이터를 확장하여 새로운 엑셀 파일로 저장합니다."""
    df = pd.read_excel(input_path)

    rfp_data = {
        'rfp': [],
        '목차': [],
        '페이지': []
    }

    # 각 행을 처리하여 RFP 데이터를 확장합니다.
    for row in df.itertuples(index=False, name='Pandas'):
        result = parse_and_expand(str(row.rfp))
        for rfp in result:
            rfp_data['rfp'].append(rfp)
            rfp_data['목차'].append(row.제목)
            rfp_data['페이지'].append(row.페이지)

    df_rfp = pd.DataFrame(rfp_data)
    df_rfp.to_excel(output_path, index=False)
    print(f'DataFrame이 {output_path} 파일로 저장되었습니다.')

def main():
    excel_file_path = '/Users/ngk/Downloads/test/제안서_슬라이드_정보.xlsx'
    out_file_path = '/Users/ngk/Downloads/test/제안서_슬라이드_정보(rfp)2.xlsx'
    process_excel_file(excel_file_path, out_file_path)

if __name__ == '__main__':
    main()
