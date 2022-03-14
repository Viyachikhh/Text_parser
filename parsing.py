import re
import xml.etree.ElementTree as ET
import zipfile

import spacy


def printRecur(root, list_sentence, label_sentence):
    global indent
    global previous_tag

    if root.text is not None and re.sub(r'{.*}', '', root.tag.title()) != 'Instrtext':
        element = root.attrib.get('name', root.text)
        #  element = re.sub(r'(^ +?)|( +?$)', '', element)  # remove spaces from start and end
        #  element = re.sub(r'\([a-z]+?\)', '', element)
        #  element = re.sub(r"[\[\]();:●£*$/’]", '', element)  # remove ()[]
        list_sentence.append(element)
        if previous_tag == 'Color':
            label_sentence.append(element)
    previous_tag = re.sub(r'{.*}', '', root.tag.title())
    #print(' ' * indent + f"{re.sub(r'{.*}', '', root.tag.title())}: {root.attrib.get('name', root.text)}")

    indent += 4
    for elem in root.getchildren():
        printRecur(elem, list_sentence, label_sentence)
    indent -= 4

    return list_sentence, label_sentence


def parse_file(path):
    with zipfile.ZipFile(path) as docx:
        tree = docx.read('word/document.xml').decode('utf-8')
    element = ET.fromstring(tree)

    sentences = []
    labels = []

    sentences, labels = printRecur(element, sentences, labels)
    sentences = list(map(lambda x: re.sub(r'\d+[.,]?\d*?', '', x), sentences))
    sentences = list(map(lambda x: re.sub(r'\([a-z]{1,2}\)', '', x), sentences))
    sentences = list(map(lambda x: re.sub(r'[…]', '', x), sentences))
    sentences = list(map(lambda x: re.sub(r'(^ +?)|( +?$)', '', x), sentences))
    sentences = [el for el in sentences if el not in ('', ',', ';', ' ', '-')]

    i = 0
    while i < len(sentences):
        if sentences[i] == '.':
            elem = sentences[i - 1] + sentences[i]
            sentences[i - 1] = elem
            del sentences[i]
        i += 1

    sentences = ' '.join(sentences)

    labels = [el for el in labels if el not in ('', ' ')]

    nlp = spacy.load('en_core_web_sm')

    res = nlp(sentences)
    #for token in res:
    #    if token.is_stop:
    #        print(token.text, token.lemma_, token.pos_, token.is_stop)

    return sentences, labels


indent = 0
previous_tag = None

archive = zipfile.ZipFile('docx.zip')
all_files = archive.namelist()
files = [el for el in all_files if el[-5:] == '.docx']

sentence, labels = parse_file(files[0])
print(len(sentence), '\n', len(labels))
print(labels)



