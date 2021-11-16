import openpyxl as excel
import nltk,random
from nltk.corpus import wordnet as wn
from nltk.tokenize import word_tokenize

varb = ['VBD','VBG','VBN','VBZ','VBZ','VBP']
wb = excel.load_workbook('databank_englishB.xlsx')
sheet = wb['シート1']
question = {}
pos = []

for line in sheet.iter_rows():
    question.setdefault(line[0].row,[line[0].value,line[1].value,line[2].value])

## 英文分割関数
def division(sentence):
    sentence = word_tokenize(str(sentence))
    pos = nltk.pos_tag(sentence)
    for i in range(1,len(pos)):
        if pos[i][1] in varb:
            if wn.morphy(pos[i][0],'v') != None:
                sentence[i] = wn.morphy(pos[i][0],'v')
    return sentence

def sort():
    for s in range(len(question)):
        sentence = division(question[s+1][2])   ##英文分割部分
        for i in range(10):
            random.shuffle(sentence)            ##英文並び替え部分
        char = ''
        for c in sentence:
            char = char + c + ' / '
        sheet['D'+str(s+1)] = char[:-3]     ##excelデータに出力

    wb.save('databank_englishB-copy.xlsx')  ##excelデータの書き出し

sort()
