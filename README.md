# Quick reference 
 파워포인트에서 조견표에 맞는 데이터를 추출하여 관련정보를 엑셀로 만들어주는 코드 
 
## 필수조건
1. ppt 내 shape naming 작업 
2. ppt명이 로마자로 시작되어야함 
- 페이지번호에 사용됨
3. 엑셀파일 생성후 같은 위치에 둬야 함 
- 사용방법 4번이 실행되려면 같은 위치에 있어양함

## 사용방법
1. python 설치후 pip install pip or pip3 로 실행
```
pip3 install python-pptx pandas openpyxl pathlib requests
```

2. config.ini 파일 설정
```
[DEFAULT]
PageNo = 페이지번호(없어도됨)
Title = 제목
Rfp = rfp정보
Navigation= 네비게이션
DirectoryPath = ppt파일경로
```

3. slide_info/main.py 실행 ppt 파일 경로에 동일하게 엑셀파일 생성

4. rfp_group/main.py 실행

5. 윈도우 실행  configparser UnicodeDecodeError: 'cp949' 에러 발생시
   환경변수 PYTHONUTF8 추가 아래 그림과 같이
![image](https://github.com/user-attachments/assets/920f42f5-0972-49e0-a973-3bf903ae9af0)

