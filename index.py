import re
import requests
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor

def create_presentation():
    return Presentation()

def add_title_slide(prs, title_text):
    if not title_text:
        return
    title_slide_layout = prs.slide_layouts[6] 
    slide = prs.slides.add_slide(title_slide_layout)
    title = slide.shapes.add_textbox(left=Inches(1), top=Inches(0.5), width=Inches(14), height=Inches(4))
    title.text = title_text
    title.text_frame.text = title_text
    index = 0
    title.text_frame.margin_top = Inches(2.75)

    for paragraph in title.text_frame.paragraphs:
        if index == 0:    
            paragraph.alignment = PP_ALIGN.CENTER
            title.text_frame.word_wrap = True
            if paragraph.runs:
                title_format = paragraph.runs[0].font
            title_format.bold = True
            title_format.name = 'Century Gothic'
            title_format.size = Pt(72)
            title_format.color.rgb = RGBColor(255, 255, 255)
        else:
            paragraph.alignment = PP_ALIGN.CENTER
            title.text_frame.word_wrap = True
            if paragraph.runs:
                title_format = paragraph.runs[0].font
            title_format.bold = False
            title_format.name = 'Century Gothic'
            title_format.size = Pt(60)
            title_format.color.rgb = RGBColor(255, 255, 255)
            title.text_frame.margin_top = Inches(1)

        index = index + 1
    
    slide.background.fill.solid()
    slide.background.fill.fore_color.rgb = RGBColor(31,56,100)

def add_content_slide(prs, content_text, max_chars_per_line=220):
    if not content_text:
        return
    
    content_slide_layout = prs.slide_layouts[6]
    words = content_text.split()
    lines = []
    current_line = ''

    for word in words:
        if word == '<br>':
            lines.append(current_line)
            current_line = ''
        else:
            if len(current_line) + len(word) + 1 <= max_chars_per_line:
                current_line += word + ' '
            else:
                lines.append(current_line.strip())
                current_line = word + ' '

    if current_line:
        lines.append(current_line.strip())

    for line in lines:
        if not line:
            line = " "

        slide = prs.slides.add_slide(content_slide_layout)
        content = slide.shapes.add_textbox(left=Inches(0.5), top=Inches(0.5), width=Inches(15), height=Inches(8))
        content.text = line
        content.text_frame.text = line        

        slide.background.fill.solid()
        slide.background.fill.fore_color.rgb = RGBColor(31,56,100)

        if "AS:" in line:
            adjust_text_box_size(content, True, len(line) > 86)
        else:
            adjust_text_box_size(content, False, False)

def adjust_text_box_size(shape, isCenter, isLarge):
    text_frame = shape.text_frame

    for paragraph in shape.text_frame.paragraphs:
        if not paragraph.runs:
            return
        font_format = paragraph.runs[0].font
        font_format.color.rgb = RGBColor(255,255,255)
        font_format.bold = isCenter
        font_format.name = 'Century Gothic'
        if isCenter:
            font_format.size = Pt(65)
        else:    
            font_format.size = Pt(60)
        paragraph.alignment = PP_ALIGN.LEFT

    if isCenter:
        if isLarge:
            text_frame.margin_top = Inches(1.25) 
        else:
            text_frame.margin_top = Inches(3) 


    text_frame.word_wrap = True
    text_frame.margin_right = 0  
    text_frame.margin_left = 0  


def save_presentation(prs, file_path):
    prs.save(file_path)

def get_parts(html, htmlId, get_text, get_sub = False):
    sub = ''
    text = ''
    title = ''
    div = html.find(id=f"{htmlId}")
    title_founded = ''

    if not div:
        return {'title': title, 'text': text, 'sub': sub}
    if div.find('b'):
        title_founded = div.find('b').getText()
    elif div.find('strong'):
        title_founded = div.find('strong').getText()
    else:
        ImportError()
        
    if get_sub:
        sub_founded = div.find_all('p')[1].getText()
        if sub_founded:
            sub = sub_founded

    if title_founded:
        title = title_founded 

    if get_text:
        paragraphs = div.find_all('p')[1:]
        text = '\n'.join([p.getText() for p in paragraphs])
        text = re.sub(r'(?<=[a-zA-Zá-üÁ-Ü\s])\d+|\d+(?=[a-zA-Zá-üÁ-Ü\s])', '', text, flags=re.UNICODE)

    return {'title': title, 'text': text, 'sub': sub}

def get_txt_file(path):
    try:
        with open(path, 'r', encoding='utf-8') as file:
            text = file.read()
        return text
    except FileNotFoundError:
        print(f"O arquivo '{path}' não foi encontrado.")
    except Exception as e:
        print(f"Ocorreu um erro ao ler o arquivo: {e}")

def main():
    ###### ALTERE AQUI
    ###### ALTERE AQUI
    ###### ALTERE AQUI
    ###### ALTERE AQUI
    # Após alteração precione "CTRL + S"
    # Após alteração precione "CTRL + S"
    url = 'https://liturgia.cancaonova.com/pb/liturgia/5a-semana-da-quaresma-quinta-feira-6/?sDia=21&sMes=03&sAno=2024' # ver site da canção nova
    oracao = get_txt_file('oracoes_eucaristicas/eucaristica_2_anunciamos.txt') # ver pasta de arquivos "oracoes_eucaristicas"
    comGloria = False 
    comCreio = True
    # Após alteração precione "CTRL + S"
    # Após alteração precione "CTRL + S"
    ###### ALTERE AQUI
    ###### ALTERE AQUI
    ###### ALTERE AQUI
    ###### ALTERE AQUI
    
    response = requests.get(url)
    html_content = response.content
    html = BeautifulSoup(html_content, 'html.parser')
    title = html.select_one('h1.entry-title').getText().strip()

    l1 = get_parts(html, 'liturgia-1', False, True)
    s = get_parts(html, 'liturgia-2', False, True)
    l2 = get_parts(html, 'liturgia-3', False, True)
    e = get_parts(html, 'liturgia-4', False, False)

    pai_nosso = get_txt_file('outras_oracoes/pai_nosso.txt')
    creio = get_txt_file('outras_oracoes/creio.txt')

    presentation = create_presentation()
    presentation.slide_width = Inches(16)
    presentation.slide_height = Inches(9)  

    add_title_slide(presentation, title)
    add_title_slide(presentation, "CANTO DE ENTRADA")    
    add_title_slide(presentation, " ")    
    add_title_slide(presentation, "ATO PENITENCIAL")    
    add_title_slide(presentation, " ")
    if comGloria:    
        add_title_slide(presentation, "GLÓRIA")    
        add_title_slide(presentation, " ")
    add_title_slide(presentation, l1['title'] + '\n \n' + l1['sub'])
    add_content_slide(presentation, l1['text'])
    add_title_slide(presentation, " ")    
    add_title_slide(presentation, s['title'])
    add_title_slide(presentation, s['sub'])
    add_content_slide(presentation, s['text'])
    add_title_slide(presentation, " ")    

    if l2['title']:
        add_title_slide(presentation, l2['title'] + '\n \n' + l2['sub'])
        add_content_slide(presentation, l2['text'])

    add_title_slide(presentation, " ")    
    add_title_slide(presentation, "ACLAMAÇÃO AO EVANGELHO")    
    add_title_slide(presentation, e['title'])
    add_content_slide(presentation, e['text'])
    add_title_slide(presentation, " ")    
    add_title_slide(presentation, "HOMILIA")  
    if comCreio:
        add_content_slide(presentation, creio)
        add_title_slide(presentation, " ")    
    add_title_slide(presentation, "PRECES DA COMUNIDADE")    
    add_title_slide(presentation, " ")    
    add_title_slide(presentation, "OFERTÓRIO")    
    add_title_slide(presentation, " ")    
    add_title_slide(presentation, "ORAÇÃO EUCARISTICA")  
    add_content_slide(presentation, oracao)
    add_title_slide(presentation, " ")  
    add_content_slide(presentation, pai_nosso)
    add_title_slide(presentation, " ")  
    add_title_slide(presentation, "CORDEIRO")  
    add_title_slide(presentation, " ")  
    add_title_slide(presentation, "COMUNHÃO")  
    add_title_slide(presentation, " ")  
    add_title_slide(presentation, " ")  

    
    save_presentation(presentation, "Missa textos.pptx")

if __name__ == "__main__":
    main()
