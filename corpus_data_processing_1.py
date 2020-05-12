import xlrd
import xlwt
import os
from xlrd import xlsx
import xlwt
from xlwt import *
import nltk
from nltk import tokenize
import re
from nltk.tokenize import WordPunctTokenizer
import numpy as np

sentence_list = []
Convergent_questions = []
Convergent_questions_label = []
Convergent_questions_before_revision=[]
Convergent_questions_after_revision=[]
Divergent_questions_before_revision=[]
Divergent_questions_after_revision=[]
Convergent_words_search = ["Which", "which", "What", "what", "Do", "Does", "Is", "Are", "Has", "Have", "Who", "whom",
                           "How", "how", "When", "when", "Why"]
Divergent_questions = []
Divergent_questions_label = []
Divergent_words_search = ["Could", "such", " Can you ", "If"]
label_list = ['rel', '?', 'rel', 'rel', 'rel', 'rel', 'rel', '?', '?', 'rel', 'rel', 'irrel', 'rel', '?', 'rel', 'rel',
              'rel', 'rel', 'rel', 'rel', 'rel', '?', 'rel', 'rel', 'rel', 'irrel', 'rel', 'rel', 'rel', 'rel', 'rel',
              'rel', 'rel', 'rel', 'rel', '?', 'rel', '?', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'irrel', 'rel',
              'rel', 'rel', 'rel', '?', 'rel', 'rel', 'rel', 'rel', 'skip', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel',
              'rel', 'rel', 'rel', 'skip', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel',
              'rel', 'rel', 'irrel', 'rel', 'rel', '?', 'irrel', 'rel', 'rel', 'irrel', 'rel', 'rel', 'rel', 'rel',
              'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'skip', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel',
              'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', 'rel', '?']


def data_preprocess():
    data = xlrd.open_workbook(r'data of feedback with raising questions.xlsx')
    sheet = data.sheet_names()
    del sheet[0]
    print('All sheet:%s' % sheet)
    before_revision0 = []
    after_revision0 = []
    before_revision = []
    after_revision = []
    teacher_feedback = []
    label = []
    label_index = []
    for i in range(len(sheet)):
        i = 1
        sheet[i] = data.sheet_by_index(i)
        before_revision0.append(sheet[i].col_values(2))
        after_revision0.append(sheet[i].col_values(3))
        teacher_feedback.append(sheet[i].col_values(4))
        label.append(sheet[i].col_values(9))

    for sentence_rev in before_revision0:
        for sentence_rev0 in sentence_rev:
            before_revision.append(sentence_rev0)
    #print(before_revision)
    for sentence_rev in after_revision0:
        for sentence_rev0 in sentence_rev:
            after_revision.append(sentence_rev0)
    #print(after_revision)


    print("This is the overall list:", teacher_feedback)
    # label_index.append(label_real)
    # print("This is certain label:",label_list[5])
    print("This is certain label:", label)

    i = 0

    for sentences in teacher_feedback:
        i += 1
    print("This is the simplified list:", i, sentences)
    k = 0
    m = 0
    n = 0
    for index in sentences:
        k += 1
        sentence = index
        # print(sentence_with_words)
        sen_tokenizer = nltk.data.load('tokenizers/punkt/english.pickle')
        sentence_seperate = nltk.word_tokenize(sentence)

        for words in sentence_seperate:

            if words in Convergent_words_search:
                #m += 1
                Convergent_questions.append(sentence)
                b = sentences.index(sentence)
                Convergent_questions_label.append(label_list[b])
                Convergent_questions_before_revision.append(before_revision[b])
                Convergent_questions_after_revision.append(after_revision[b])
                # print(Convergent_questions_label)
                continue
            elif words in Divergent_words_search:
                n += 1
                Divergent_questions.append(sentence)
                c = sentences.index(sentence)
                Divergent_questions_label.append(label_list[c])
                Divergent_questions_before_revision.append(before_revision[c])
                Divergent_questions_after_revision.append(after_revision[c])

    print("This is convergent question:", Convergent_questions)
    print("This is divergent question:", Divergent_questions)

    print("The ratio convergent questions to divergent questions:", m / n)
    print("The  proportion of convergent questions :", m / k)
    print("The  proportion of divergent questions :", n / k)
    return m,n,k

def caculate():
    p = 0
    q = 0
    f = 0
    y = 0
    for label_word in Convergent_questions_label:
        if "rel" in label_word:
            p += 1
        elif "irrel" in label_word:
            q += 1
        elif "?" in label_word:
            f += 1
        elif "skip" in label_word:
            y += 1
    print("rel in Convergent questions tag:", p / m)
    print("irrel in Convergent questions tag:", q / m)
    print(" ? in Convergent questions tag:", f / m)
    print("skip in Convergent questions tag:", y / m)
    p1 = 0
    q1 = 0
    f1 = 0
    y1 = 0
    for label_word in Divergent_questions_label:
        if "rel" in label_word:
            p1 += 1
        elif "irrel" in label_word:
            q1 += 1
        elif "?" in label_word:
            f1 += 1
        elif "skip" in label_word:
            y1 += 1
    print("rel in Divergent questions tag:", p1 / n)
    print("irrel in Divergent questions tag:", q1 / n)
    print(" ? in Divergent questions tag:", f1 / n)
    print("skip in Divergent questions tag:", y1 / n)


### writing into the txt file
# file=open('convergent question.txt','w')
# for single_sentence in Convergent_questions:
#     file.write(str(single_sentence));
#     file.write("\n")
# file.close()
#
# file=open('divergent question.txt','w')
# for single_sentence in Divergent_questions:
#     file.write(str(single_sentence));
#     file.write("\n")
# file.close()
def set_style(name, height, bold=False):
    style = xlwt.XFStyle()
    font = xlwt.Font()
    font.name = name
    font.bold = bold
    font.color_index = 4
    font.height = height
    style.font = font
    return style


### writing into the excel file
def write_excel():
    f = xlwt.Workbook()
    # cell_format = f.add_format({'bold': True})
    sheet1 = f.add_sheet('Convergent_questions', cell_overwrite_ok=True)
    sheet2 = f.add_sheet('Divergent_questions', cell_overwrite_ok=True)

    row0 = ["Before revision","After revision","Teacher feedback", "Label"]
    colum000 = Convergent_questions_before_revision
    colum0000= Convergent_questions_after_revision
    colum0 = Convergent_questions
    colum00 = Convergent_questions_label

    colum111 =Divergent_questions_before_revision
    colum1111 =Divergent_questions_after_revision
    colum1 = Divergent_questions
    colum11 = Divergent_questions_label
    c = 0
    for c in range(0, len(row0)):
        # sheet1.col(c).width = 8888
        sheet1.write(0, c, row0[c], set_style('Times New Roman', 220, True))
        sheet2.write(0, c, row0[c], set_style('Times New Roman', 220, True))

    for c in range(0, len(colum0)):
        # sheet1.col(c).width = 256 * 20
        # sheet1.set_row(0, 20, cell_format)
        sheet1.write(c + 1, 0, colum000[c], set_style('Times New Roman', 220, True))
        sheet1.write(c + 1, 1, colum0000[c], set_style('Times New Roman', 220, True))
        sheet1.write(c + 1, 2, colum0[c], set_style('Times New Roman', 220, True))
        sheet1.write(c + 1, 3, colum00[c], set_style('Times New Roman', 220, True))
    for c in range(0, len(colum1)):
        # sheet1.col(c).width = 256 * 20
        sheet2.write(c + 1, 0, colum111[c], set_style('Times New Roman', 220, True))
        sheet2.write(c + 1, 1, colum1111[c], set_style('Times New Roman', 220, True))
        sheet2.write(c + 1, 2, colum1[c], set_style('Times New Roman', 220, True))
        sheet2.write(c + 1, 3, colum11[c], set_style('Times New Roman', 220, True))
    f.save('Corpus data.xls')


if __name__ == '__main__':
    data_preprocess()
    #caculate()
    write_excel()
