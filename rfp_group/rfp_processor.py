import pandas as pd
from rfp_group.utils import parse_and_expand, format_pages
import traceback
import logging
logging.basicConfig(filename="log.txt", level=logging.DEBUG, encoding='utf-8')
logging.debug("📌 rfp_processor.py 실행 시작됨")

def process_excel_file(input_path, output_path):
    logging.debug(f"엑셀 파일 읽기 시작: {input_path}")
    try:
        df = pd.read_excel(input_path)

        rfp_data = {
            'rfp': [],
            '목차': [],
            '페이지': []
        }

        for row in df.itertuples(index=False, name='Pandas'):
            result = parse_and_expand(str(row.rfp))
            for rfp in result:
                rfp_data['rfp'].append(rfp)
                rfp_data['목차'].append(row.제목)
                rfp_data['페이지'].append(row.페이지)

        df_rfp = pd.DataFrame(rfp_data)

        merged_data = {
            'rfp': [],
            '목차': [],
            '페이지': []
        }

        for name, group in df_rfp.groupby('rfp'):
            combined_titles = "\n".join(group['목차'])
            if len(group) == 1:
                merged_data['rfp'].append(group.iloc[0]['rfp'])
                merged_data['목차'].append(combined_titles)
                merged_data['페이지'].append(group.iloc[0]['페이지'])
            else:
                pages = group['페이지'].tolist()
                formatted_pages = format_pages(pages)
                merged_data['rfp'].append(name)
                merged_data['목차'].append(combined_titles)
                merged_data['페이지'].append(formatted_pages)
        logging.debug("병합된 데이터 저장 시작")
        merged_df = pd.DataFrame(merged_data)
        merged_df.to_excel(output_path, index=False)

    except Exception as e:
        with open("log.txt", "a", encoding="utf-8") as f:
            f.write("\n[process_excel_file 에러]\n")
            f.write(traceback.format_exc())
        raise  # 예외를 다시 main.py로 넘겨줘야 GUI가 알 수 있음
