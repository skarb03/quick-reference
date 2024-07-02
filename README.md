# Quick reference 
 파워포인트에서 조견표에 맞는 데이터를 추출하여 관련정보를 엑셀로 만들어주는 코드 
 
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
EvalItem= 평가항목명
DirectoryPath = ppt파일경로
```

3. slide_info  main.py 실행 ppt 파일 경로에 동일하게 엑셀파일 생성

4. rfp_group main.py 실행 3번 엑세파일 현재 위치 그대로 둬야 작동함.



