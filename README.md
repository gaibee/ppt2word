# ppt2word
피피티를 워드로 변환해주는 프로그램


### config file
"COLUMN_HEADER_WIDTH" : 2.41, # 슬라이드 이름을 나타내는 셀의 가로 크기(cm)<br>
"COLUMN_PPT_WIDTH" : 6.75, # ppt 그림을 나타내는 셀의 가로 크기(cm)<br>
"COLUMN_NOTE_WIDTH" : 6.75, # 슬라이드 노트를 나타내는 셀의 가로 크기(cm)<br>
"FONT": "맑은 고딕", # 폰트<br>
"FONT_SIZE": 10 # 폰트 크기(pt)<br>

### 사용법
python main.py<br>
실행 후 변환하고 싶은 pptx 파일 선택<br>
변환이 완료되면 해당 폴더에 save.docx로 저장됨<br>

### Requirements
pip install python-pptx<br>
pip install python-docx<br>
pip install python-pptx-interface<br>

### 변환 이미지
<img src="https://user-images.githubusercontent.com/64792575/180704860-ae021ef5-6f03-4f1c-ac7b-28f048f673c9.png" width=800>
