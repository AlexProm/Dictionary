# -*- coding: utf-8 -*-
import re
import os
import xlrd
import time

reg = re.compile('[^a-zA-Z ]')

def Coise_var_input_Text():
    Coise_input_Text = input('Ввести адрес файла = 1, Ввести в строку = 2: ')
    if Coise_input_Text == '1':
        namefile = input('Введите адрес файла: ')
        f = open(namefile, "r")
        Txt_file = f.read()
        f.close()
    elif Coise_input_Text == '2':
        Txt_file = input('Скопируй сюда текст: ')
    else:
        print('\nВведён неверный запрос\n')
        Coise_var_input_Text()
    return str(Txt_file)

def File_Open_name(filname):
    '''Считывание Массива, содержащего другие массивы, находящегося по адресу'''
    filname = str(filname)
    f = open(filname, "r")
    Txt_file = f.readline()
    f.close()
    # Разделение входных
    Txt_file = Txt_file.split()
    Txt_file = list(Txt_file)
    return Txt_file

def Open_Exel_DB():
    '''Открывает базу данных записанную в файле Excel'''
    # открываем файл
    f = xlrd.open_workbook('DB.xls', formatting_info=True)
    sheet = f.sheet_by_index(0)
    # получаем список значений из всех записей ( Если вместо sheet.nrows будет 1 - только первую строчку)
    vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    vals = list(vals)
    return vals

def Search_index_Word(Word, List):
    index = -1
    for idx in range(len(List)):
        if Word == List[idx]:
            index = idx
    return index
# С английского  на знакомый язык
def New_words_En_Ua(Slova):
    #Часть получени новых слов
    New_words = []
    for word in range(len(Slova)):
        Word = Slova[word - 1]
        Searh = Search_index_Word(Word, Original_Word)
        if Searh > -1:
            Slova[word - 1] = Translate_Word[Searh]
        elif Searh == -1:
            xLeng = len(New_words)
            New_words.append(Slova[word - 1])
    New_words = set(New_words)
    New_words = list(New_words)
    return New_words

def File_Open():
    '''Открываем переводимый файл и уменьшаем регистр'''
    f = open("TEXT.txt", "r")
    TEXT = f.read()
    f.close()
    TEXT = reg.sub('', TEXT)
    return TEXT.lower()

def Level():
    words = len(Original_Word)
    if words < 1000:
        level = 'A1 Beginner'
    elif words < 1500:
        level = 'A1 Elementary'
    elif words < 2000:
        level = 'A2 Pre-Intermediate'
    elif words < 3000:
        level = 'B1 Intermediate'
    elif words < 4000:
        level = 'B2 Upper-Intermediate'
    elif words < 6000:
        level = 'C1 Advanced'
    else: level = 'C2 Proficient'
    return level

def LevelUp():
    words = len(Original_Word)
    if words < 1000:
        levelup = 'До уровня A1 Elementary, осталось {0} слов, или {1:2.2f}% от требуемого'.format(1000 - words, words/1000)
    elif words < 1500:
        levelup = 'До уровня A2 Pre-Intermediate, осталось {0} слов, или {1:2.2f}% от требуемого'.format(1500 - words, words/1500)
    elif words < 2000:
        levelup = 'До уровня B1 Intermediate, осталось {0} слов, или {1:2.2f}% от требуемого'.format(2000 - words, words/2000)
    elif words < 3000:
        levelup = 'До уровня B2 Upper-Intermediate, осталось {0} слов, или {1:2.2f}% от требуемого'.format(3000 - words, words / 3000)
    elif words < 4000:
        levelup = 'До уровня C1 Advanced, осталось {0} слов, или {1:2.2f}% от требуемого'.format(4000 - words, words /4000)
    elif words < 6000:
        levelup = 'До уровня C2 Proficient, осталось {0} слов, или {1:2.2f}% от требуемого'.format(6000 - words, words / 6000)
    else: levelup = 'Ты C2 Proficient, и этим всё сказанно... Учи новый язык или не парься'
    return levelup

def Report():
    '''Последний отчёт/показатель - Самый свежий'''
    f = open("English Last Raport.txt", "w")
    xxx = '-----------------------------------------------------------------------------------------------------------------------------'
    f.write('''
{0}
Дата: {15}
{0}
Оригинал текста (Slova), Всего слов: {1}
{2}
{0}
Переведённый текст: {3}
{4}
{0}
Процент изученных слов в тексте (по количеству слов в тексте): ...{5:+2.2f}%
Процент незнакомых слов: ...........{6:+2.2f}%
Процент изученных слов в тексте (по уникальным словам): ..........{7:+2.2f}%
Процент незнакомых слов: ...........{8:+2.2f}%
Впервые встреченных слов к тексту: ...............................{9:+2.2f}%
По отношению к известным словам: ...{10:-2.2f}%
{0}
Новые слова для изучения: Всего {11} (Всего изученно в словаре: {12}),(Процентов от известного: {13:+5.3f}%)\n{14}
{0}
Уровень владения Английским языком: {16}
{17}
{0}

'''.format(xxx, str(len(Slova)), String, TEXT_NAME, Text_after_Translate, KnowPercent_text, -notKnowPercent_text, KnowPercent_unikal, -notKnowPercent_unikal, Whath_words_Percent_text, -Whath_words_Percent_all, str(len(New_words)), str(len(Original_Word)), ((len(New_words) / (len(Original_Word))) * 100), (" ".join(New_words)), time.strftime(("%Y.%m.%d"), time.gmtime()), Level(), LevelUp() ) )
    #       0         1           2       3                 4                   5                    6                   7                   8                           9                       10                         11                     12                                   13                                  14                          15                            16        17
    f.close()


#Инициализируем переменные
Txt_file = Open_Exel_DB()
Original_Word = []
Translate_Word = []
for ids in range(len(Txt_file)):
    Original_Word.append(Txt_file[ids][1])
    Translate_Word.append(Txt_file[ids][3])
Whath_words = File_Open_name("Встречавшиеся неизвестные слова.txt")
New_words = []

#Открываем файл для перевода(очищаем от лишних символов) и делаем из него массив
String = File_Open()
Slova = String.split()

#Перечень неизвестных слов
New_words = New_words_En_Ua(Slova)

add_Wath_words = []
#Сохранение новых никода не встречавшихся слов в общий список встреченных/виденных слов
for word in range(len(New_words)):
    Word = New_words[word]
    Searh = Search_index_Word(Word, Whath_words)
    if Searh == -1:
        Whath_words.append(New_words[word])
        add_Wath_words.append(New_words[word])

#Сохранение встречавшихся слов
Whath_words = set(Whath_words)
Whath_words = list(Whath_words)
f = open("Встречавшиеся неизвестные слова.txt", "w")
Whath_words_F = " ".join(Whath_words)
f.write(Whath_words_F)
f.close()

#Часть перевода текста:
#Slova
Translate_text = []
for word in range(len(Slova)):
    Word = Slova[word]
    Searh = Search_index_Word(Word, Original_Word)
    if Searh > -1:
        Translate_text.append(Translate_Word[Searh])
    elif Searh == -1:
        Translate_text.append(Slova[word])

Text_after_Translate = " ".join(Translate_text)

#Имя переведённого текста - из первфх 10 слов в нём
TEXT_NAME = ''
for i in range(10):
    TEXT_NAME = TEXT_NAME + str(Translate_text[i]) + ' '

Translate_text = set(Translate_text)

#Восстанавливаем Slova, в которых часть слов теперь переведены в исходный текст
Slova = String.split()

#Подсчёт % понимание текста по весовым категорим слов
Know_text = 0
notKnow_text = 0
for word in range(len(Slova)):
    Word = str(Slova[word])
    Searh = Search_index_Word(Word, Original_Word)
    if Searh > -1:
        Know_text = Know_text +1
    elif Searh == -1:
        notKnow_text = notKnow_text +1

#Подготовка отчёта
xxx = len(Slova)
KnowPercent_text = (Know_text / xxx) * 100
notKnowPercent_text = (notKnow_text / xxx) * 100
Translate_text = list(Translate_text)
KnowPercent_unikal = ((len(Translate_text) -len(New_words)) / len(Translate_text)) * 100
notKnowPercent_unikal = (len(New_words) / len(Translate_text)) * 100
Whath_words_Percent_all = (len(add_Wath_words) / len(Whath_words)) * 100
Whath_words_Percent_text = (len(add_Wath_words) / len(Translate_text)) * 100
Translate_text = " ".join(Translate_text)

#Сохроняем результаты
Report()
print("Выполнено!")