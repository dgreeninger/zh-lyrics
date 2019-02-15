import os, re, json
from pprint import pprint
try:
    from xml.etree.cElementTree import XML
except ImportError:
    from xml.etree.ElementTree import XML
import zipfile
import sys
reload(sys)
sys.setdefaultencoding('utf8')

"""
Module that extract text from MS XML Word document (.docx).
(Inspired by python-docx <https://github.com/mikemaccana/python-docx>)
"""

WORD_NAMESPACE = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
PARA = WORD_NAMESPACE + 'p'
TEXT = WORD_NAMESPACE + 't'


def get_docx_text(path):
    """
    Take the path of a docx file as argument, return the text in unicode.
    """
    document = zipfile.ZipFile(path)
    xml_content = document.read('word/document.xml')
    document.close()
    tree = XML(xml_content)

    paragraphs = []
    for paragraph in tree.getiterator(PARA):
        texts = [node.text
                 for node in paragraph.getiterator(TEXT)
                 if node.text]
        if texts:
            paragraphs.append(''.join(texts))

    return '\n\n'.join(paragraphs)



zh = {}
for filename in sorted(os.listdir('./lyrics'), key=int):
    file_name = './lyrics/'+filename
    song = get_docx_text(file_name).splitlines()
    if "emember" in song[0]:
        song.pop(0)
        first_line = song[1].split(' ', 1)
        song.pop(1)
        verses = song
    else:
        first_line = song[0].split(' ', 1)
        song.pop(0)
        verses = song

    song_number = int(first_line[0])
    zh[song_number] = {}
    title = first_line[1].encode('utf-8').replace('\xe2', "'").replace('\x99', "").replace('\x80','')
    zh[song_number]['title'] = str(title)

    zh[song_number]['verses'] = {}
    for verse in verses:
        verse=verse.encode('utf-8').replace('\xe2', "'").replace('\x99', "").replace('\x80','')
        pattern = re.compile(r'[0-9]+\.')
        match = re.findall(pattern, verse)
        if match:
            verse_number = match[0].translate(None, '.')
            verse = verse.split(' ', 1)[-1].replace('\xe2', "'")
            zh[song_number]['verses'][verse_number] = verse

print(json.dumps(zh, ensure_ascii=False, indent=2))
