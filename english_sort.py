import openpyxl as excel
import nltk,random
from nltk.corpus import wordnet as wn
from nltk.tokenize import word_tokenize
import re

varb = ['VBD','VBG','VBN','VBZ','VBZ','VBP']
shortened = {
    '\'m': ' am',
    '\'re': ' are',
    'aren\'t': 'are not',
    'is\'nt': 'is not',
    'it\'s': 'it is',
    'It\'s': 'It is',
    'wasn\'t': 'was not',
    'weren\'t': 'were not',
    'don\'t': 'do not',
    'Don\'t': 'Do not',
    'doesn\'t': 'does not',
    'Doesn\'t': 'Does not',
    'didn\'t': 'did not',
    'Didn\'t': 'Did not',
    'won\'t': 'will not',
    'haven\'nt': 'have not',
    'can\'t': 'can not',
    'let\'s': 'let as',
    'Let\'s': 'Let as'
}
shortened_re = re.compile('(?:' + '|'.join(map(lambda x: '\\b' + x + '\\b', shortened.keys())) + ')')   #re.compile関数に辞書のキーを入力したものを定義
wb = excel.load_workbook('databank_englishB.xlsx')
sheet = wb['シート1']
question = {}
pos = []
exception = ['Did','Stop','Do','Whether','Please','Dark']   #文頭にある小文字にする名詞
initial_check = False   #文頭にある名詞をチェックするかどうか


for line in sheet.iter_rows():
    question.setdefault(line[0].row,[line[0].value,line[1].value,line[2].value])

## 英文分割関数
def division(sentence):
    sentence = word_tokenize(str(sentence))
    pos = nltk.pos_tag(sentence)
    for i in range(len(pos)):
        if (pos[i][1] != 'NNP' and sentence[i] != 'I' or sentence[i] in exception):
            sentence[i] = sentence[i].lower()       #大文字を小文字に変換する
        elif initial_check and sentence[i] != 'I':
            print(pos[i])
        if pos[i][1] in varb:
            archetypes = wn.morphy(pos[i][0],'v')
            if archetypes != None:
                sentence[i] = archetypes
    return sentence

def sort():
    for s in range(len(question)):
        sentence = shortened_re.sub(lambda x: shortened[x.group(0)], question[s+1][2])  #省略形をもとに戻す
        sentence = division(sentence)   ##英文分割部分
        for i in range(10):
            random.shuffle(sentence)            ##英文並び替え部分
        char = ''
        for c in sentence:
            char = char + c + ' / '
        sheet['D'+str(s+1)] = char[:-3]     ##excelデータに出力

    wb.save('databank_englishB-copy.xlsx')  ##excelデータの書き出し

sort()
