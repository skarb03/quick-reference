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

def format_pages(pages):
    def page_key(page):
        match = re.match(r"(\D+)-(\d+)", page)
        return (match.group(1), int(match.group(2)))
    
    pages = sorted(pages, key=page_key)
    ranges = []
    range_start = pages[0]
    range_end = pages[0]
    prev_key = page_key(pages[0])

    for i in range(1, len(pages)):
        current_key = page_key(pages[i])
        if current_key[0] == prev_key[0] and current_key[1] == prev_key[1] + 1:
            range_end = pages[i]
        else:
            if range_start == range_end:
                ranges.append(prev_key[0] + '-' + str(page_key(range_start)[1]))
            else:
                ranges.append(f"{prev_key[0]}-{page_key(range_start)[1]}~{page_key(range_end)[1]}")
            range_start = pages[i]
            range_end = pages[i]
        prev_key = current_key
    
    if range_start == range_end:
        ranges.append(prev_key[0] + '-' + str(page_key(range_start)[1]))
    else:
        ranges.append(f"{prev_key[0]}-{page_key(range_start)[1]}~{page_key(range_end)[1]}")
    
    return ','.join(ranges)

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


# 병합된 데이터를 저장할 리스트
    merged_data = {
        'rfp': [],
        '목차': [],
        '페이지': []
    }

    for name, group in df_rfp.groupby('rfp'):
        if len(group) == 1:
            merged_data['rfp'].append(group.iloc[0]['rfp'])
            merged_data['목차'].append(group.iloc[0]['목차'])
            merged_data['페이지'].append(group.iloc[0]['페이지'])
        else:
            pages = group['페이지'].tolist()
            formatted_pages = format_pages(pages)
            
            merged_data['rfp'].append(name)
            merged_data['목차'].append(group.iloc[0]['목차'])
            merged_data['페이지'].append(formatted_pages)
            
            for i in range(1, len(group)):
                merged_data['rfp'].append('')
                merged_data['목차'].append(group.iloc[i]['목차'])
                merged_data['페이지'].append('')


# 병합된 데이터로 새로운 데이터프레임 생성
    merged_df = pd.DataFrame(merged_data)
    merged_df.to_excel(output_path, index=False)
    print(f'DataFrame이 {output_path} 파일로 저장되었습니다.')

def main():
    excel_file_path = '/Users/ngk/Downloads/test/제안서_슬라이드_정보.xlsx'
    out_file_path = '/Users/ngk/Downloads/test/제안서_슬라이드_정보(rfp).xlsx'
    process_excel_file(excel_file_path, out_file_path)

if __name__ == '__main__':
    main()
