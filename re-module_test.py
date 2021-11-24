import re

shortened = {
    '\'m': ' am',
    '\'re': ' are',
    'is\'nt': 'is not',
    'it\'s': 'it is',
    'don\'t': 'do not',
    'doesn\'t': 'does not',
    'didn\'t': 'did not',
    'won\'t': 'will not',
    'haven\'nt': 'have not',
    'can\'t': 'can not'
}

shortened_re = re.compile('(?:' + '|'.join(map(lambda x: '\\b' + x + '\\b', shortened.keys())) + ')')   #re.compile関数に辞書のキーを入力したものを定義

sentence = "I didn’t know you were a friend of the famous writer."
sentence = "I didn't know you were a friend of the famous writer."

sentence = shortened_re.sub(lambda x: shortened[x.group(0)], sentence)

print(sentence)