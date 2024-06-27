import pandas as pd

def merge_data(ppts_data):
    merge_data = {
        '파일명': [],
        '슬라이드번호': [],
        '제목': [],
        '평가항목': [],
        '페이지': [],
        'rfp': []
    }
    
    for data in ppts_data:
        merge_data['파일명'].extend(data['file_name'])
        merge_data['슬라이드번호'].extend(data['slide_no'])
        merge_data['제목'].extend(data['title'])
        merge_data['평가항목'].extend(data['eval_item'])
        merge_data['페이지'].extend(data['page_no'])
        merge_data['rfp'].extend(data['rfp'])
    
    return merge_data

def save_to_excel(data, excel_file_path):
    df = pd.DataFrame(data)
    df.to_excel(excel_file_path, index=False)
    print(f'DataFrame이 {excel_file_path} 파일로 저장되었습니다.')
