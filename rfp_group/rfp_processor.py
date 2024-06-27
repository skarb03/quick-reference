import pandas as pd
from utils import parse_and_expand, format_pages

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
