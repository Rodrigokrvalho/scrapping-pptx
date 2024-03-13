import requests
import re
from bs4 import BeautifulSoup

# URL da página a ser scrapeada
url = 'https://liturgia.cancaonova.com/pb/'

# Faça o download do conteúdo da página
response = requests.get(url)
html_content = response.content

html = BeautifulSoup(html_content, 'html.parser')
title = html.select('h1.entry-title')[0].getText().strip()

def getEucaristicPrey(name):
    with open(name, 'r', encoding='utf-8') as file:
        prey = file.read()
    return prey

def getParts(htmlId, getText):
    text = ''
    div = html.find(id=f"{htmlId}")
    title = div.find('strong').getText()

    if getText:
        paragraphs = div.findAll('p')[1:]
        text = '\n'.join([p.getText() for p in paragraphs])
        text = re.sub(r'(?<=[a-zA-Zá-üÁ-Ü\s])\d+|\d+(?=[a-zA-Zá-üÁ-Ü\s])', '', text, flags=re.UNICODE)

    return {'title': title, 'text': text}

l1 = getParts('liturgia-1', True)
s = getParts('liturgia-2', True)
l2 = getParts('liturgia-3', True)
e = getParts('liturgia-4', True)
prey = getEucaristicPrey("eucaristica_1_domingo.txt")

# Salve os dados em um arquivo de texto
with open('dados_scraping.txt', 'w', encoding='utf-8') as file:
    file.write(f'Título: {title}\n')
    file.write(f'Primeira leitura: {l1["title"]}\n')
    file.write(f'Conteúdo da Primeira leitura:\n{l1["text"]}\n\n')
    file.write(f'Salmo: {s["title"]}\n')
    file.write(f'Conteúdo do Salmo:\n{s["text"]}\n\n')
    file.write(f'Segunda leitura: {l2["title"]}\n')
    file.write(f'Conteúdo da Segunda leitura:\n{l2["text"]}\n\n')
    file.write(f'Evangelho: {e["title"]}\n')
    file.write(f'Conteúdo do Evangelio:\n{e["text"]}\n')
    file.write(f'Oração eucaristica:\n{prey}\n')
