from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
import json
import os
from tkinter import filedialog
from components import ppt


# 설정 불러오기
config = {}
try:
    with open('config.json', 'r', encoding='utf-8') as f:
        config = json.load(f)
except:
    print('config.json 파일을 찾을 수 없습니다.')
    quit()

# ppt 파일 불러오기
file_path = filedialog.askopenfilename(initialdir=os.getcwd(), title = "ppt 파일을 선택 해 주세요", filetypes = (("*.pptx","*pptx"),("*.ppt","*ppt")))
if file_path == '':
    quit()

print(f'선택된 ppt 파일 : {file_path}')

# ppt 파일 -> png로 변환
if not os.path.exists('png'):
    os.makedirs('png')
    print('png 폴더 생성')

png_full_path = os.path.join( os.getcwd(), 'png' )
print(f'{file_path} ppt 파일을 {png_full_path} 폴더에 png 파일로 저장중...')
ppt.slide_to_image(png_full_path, file_path)
print('- 완료')

print('슬라이드 노트 불러오는 중...')
notes = ppt.get_slides_note(file_path)
print('- 완료')

print('Docx 파일 생성중...')
doc = Document()
table = doc.add_table(rows=len(notes), cols=3)
for i, row in enumerate(table.rows):
    print(f'슬라이드{i+1} 변환중...')
    
    cells = row.cells
    
    cells[0].width = Cm(config['COLUMN_HEADER_WIDTH'])
    cells[1].width = Cm(config['COLUMN_PPT_WIDTH'])
    cells[2].width = Cm(config['COLUMN_NOTE_WIDTH'])

    run = cells[0].paragraphs[0].add_run()
    run.text = f'슬라이드{i+1}'
    run.font.name = config['FONT']
    run.font.size = Pt(config['FONT_SIZE'])
    run._element.rPr.rFonts.set(qn('w:eastAsia'), config['FONT'])
    
    run = cells[1].paragraphs[0].add_run()
    run.add_picture(f'png/슬라이드{i+1}.PNG', width=Cm(config['COLUMN_PPT_WIDTH']))
    
    run = cells[2].paragraphs[0].add_run()
    run.text = notes[i]
    run.font.name = config['FONT']
    run.font.size = Pt(config['FONT_SIZE'])
    run._element.rPr.rFonts.set(qn('w:eastAsia'), config['FONT'])

    cells[0].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
    cells[1].paragraphs[0].paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER

    cells[0].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cells[1].vertical_alignment = WD_ALIGN_VERTICAL.CENTER
    cells[2].vertical_alignment = WD_ALIGN_VERTICAL.CENTER

doc.save('save.docx')
print('save.docx로 저장 완료')