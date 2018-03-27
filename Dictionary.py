import os
import time
from random import randint
import xlrd, xlwt

clear = lambda: os.system('cls')

def First_big_Word(word):
    word = list(word)
    word[0] = word[0].upper()
    word = ''.join(word)
    return word

def Search_index_Word_En(Word, List):
    '''Находит Индекс искомого объекта, или выдаёт значение -1, если такого слова нет Английские Слова'''
    index = -1
    for idx in range(len(List)):
        if Word == List[idx][1]:
            index = idx
    return index

def Add_New_word():
    '''Добавление нового слова'''
    clear()
    Number_word = len(Slovar) + 1
    New_word = [Number_word, input('Введи Английское слово: '), input('Введи Транскрипцию: '), input('Введи Перевод: '), time.strftime("%Y%m%d", time.gmtime())]
    Slovar.append(New_word)
    #File_Safe_Dictionary()

def Step_print_word():
    '''Печать на экран новое следующее слово с новой строки'''
    clear()
    for idx in range(len(Slovar)):
        print(str(Slovar[idx][0]) + ' ' + str(Slovar[idx][1]) + ' ' + str(Slovar[idx][2]) + ' ' + str(Slovar[idx][3]))

def print_One_Line_Word(index):
    '''Печатает на экран Слово, по индексу'''
    print(str(Slovar[index][0]) + ' ' + str(Slovar[index][1]) + ' ' + str(Slovar[index][2]) + ' ' + str(Slovar[index][3]) + '\n')

def Repeat_all_word():
    '''Повторения всего словаря'''
    clear()
    for idx in range(len(Slovar)):
        print_One_Line_Word(idx)
        Next_word = input('Для следующего слова введи что-нибудь, Введёшь 0 и прервёшь повтор')
        clear()
        if Next_word == '0':
            print('\nПовторение прерванно\n')
            break

def Repeat_10_random_word():
    '''Повторение 10 слов, выбранных случайно (через random)'''
    clear()
    for idx in range(10):
        idx_random = randint(0, len(Slovar)-1)
        print('\n' + str(idx + 1) + ' ' + str(Slovar[idx_random][1]) + ' ' + str(Slovar[idx_random][2]) + ' ' + str(Slovar[idx_random][3]) + '\n')
        if idx == 9:
            break
        Next_word = input('\nДля следующего слова введи что-нибудь, Введёшь 0 и прервёшь повтор')
        if Next_word == '0':
            print('\nПовторение прерванно\n')
            break
        clear()
    Next_word = input('Закончить повтор')

def Repeat_10_LowRepeat_word():
    '''Повторение 10 слов, которые я меньше всего повторял'''
    #Сделать переменную с количеством обращением к ней во время повторения, куда запись будет идти во время повторения...
    #Только числа, которые соответствуют кол-во повторений, а индекс его номеру в массиве, и так его и вызывать.
    #В отдельный файл записывать... Хотя если исспользовать сортировку, то номера нужны.
    clear()
    print('Пока нет\n')

def Enter_True_Letter_word():
    Language = input('Выберите режим работы.\n Со значением на известном языке 1\n Только слово с пропущенными буквами 2 ')
    if Language == '1':
        Enter_True_Letter_word_with_translate()
    elif Language == '2':
        Enter_True_Letter_word_without_translate()

def print_list_in_Propusk_Letter(List):
    STR_List = ''
    for idx in range(len(List)):
        STR_List = STR_List + str(List[idx] + '')
    return STR_List

def Propusk_Letter_Delete(String_i):
    '''Заменяет две случайные буквы, кроме последней из введённой строки/слова'''
    String_o = list(String_i)
    Len_STR = int(len(String_o))
    index_propyska = randint(0, Len_STR - 2)
    String_o[index_propyska] = '_'
    index_propyska = randint(0, Len_STR - 2)
    String_o[index_propyska] = '_'
    String_o = print_list_in_Propusk_Letter(String_o)
    return str(String_o)

def Enter_True_Letter_word_with_translate():
    '''Повторение в 10 раз больше, чем есть в словаре, выбранных случайно (через random). Вод английского перевда (пропущенных букв...'''
    clear()
    for idx in range(10*len(Slovar)):
        idx_random = randint(0, len(Slovar)-1)
        Full_word = Propusk_Letter_Delete(Slovar[idx_random][1])
        Full_word2 = input('\n' + str(idx + 1) + ' ' + (Slovar[idx_random][3]) + ' \"' + str(Full_word) + '\" Введите правильное слово: ' )
        if Full_word2 == str(Slovar[idx_random][1]):
            print('Верно')
        else: print('Не верно, верно: ' + str(Slovar[idx_random][1]))
        if idx == (10*len(Slovar)):
            break
        Next_word = input('\nДля следующего слова введи что-нибудь, Введёшь 0 и прервёшь повтор')
        if Next_word == '0':
            print('\nПовторение прерванно\n')
            break
        clear()
    Next_word = input('Закончить повтор')

def Enter_True_Letter_word_without_translate():
    '''Повторение в 10 раз больше, чем есть в словаре, выбранных случайно (через random). Вод английского перевда (пропущенных букв...'''
    clear()
    for idx in range(10*len(Slovar)):
        idx_random = randint(0, len(Slovar)-1)
        Full_word = Propusk_Letter_Delete(Slovar[idx_random][1])
        Full_word2 = input('\n' + str(idx + 1) + ' \"' + str(Full_word) + '\" Введите правильное слово: ' )
        if Full_word2 == str(Slovar[idx_random][1]):
            print('Верно')
        else: print('Не верно, верно: ' + str(Slovar[idx_random][1]))
        if idx == (10*len(Slovar)):
            break
        Next_word = input('\nДля следующего слова введи что-нибудь, Введёшь 0 и прервёшь повтор')
        if Next_word == '0':
            print('\nПовторение прерванно\n')
            break
        clear()

def Enter_True_word():
    '''Блок Выбора режима работы'''
    Language = input('Выберите режим работы.\n Выводить слова на Изучаемом языке, Вводить слова на Знакомом языке, нажмите 1\n Выводить слова на Знакомом языке, Вводить слова на Изучаемом языке, нажмите 2 ')
    if Language == '1':
        Enter_True_word_Old_New()
    elif Language == '2':
        Enter_True_word_New_Old()

def Enter_True_word_Old_New():
    '''Повторение в 10 раз больше, чем есть в словаре, выбранных случайно (через random). Вод руского перевда...'''
    clear()
    for idx in range(10*len(Slovar)):
        idx_random = randint(0, len(Slovar)-1)
        Translate_word = input('\n' + str(idx + 1) + ' ' + str(Slovar[idx_random][1]) + ' Переводится, как: ' )
        if Translate_word == str(Slovar[idx_random][3]):
            print('Верно')
        else: print('Не верно, верно: ' + str(Slovar[idx_random][3]))
        if idx == (10*len(Slovar)):
            break
        Next_word = input('\nДля следующего слова введи что-нибудь, Введёшь 0 и прервёшь повтор')
        if Next_word == '0':
            print('\nПовторение прерванно\n')
            break
        clear()
    Next_word = input('Закончить повтор')

def Enter_True_word_New_Old():
    '''Повторение в 10 раз больше, чем есть в словаре, выбранных случайно (через random). Вод английского перевда...'''
    clear()
    for idx in range(10*len(Slovar)):
        idx_random = randint(0, len(Slovar)-1)
        Translate_word = input('\n' + str(idx + 1) + ' ' + str(Slovar[idx_random][3]) + ' Переводится, как: ' )
        if Translate_word == str(Slovar[idx_random][1]):
            print('Верно')
        else: print('Не верно, верно: ' + str(Slovar[idx_random][1]))
        if idx == (10*len(Slovar)):
            break
        Next_word = input('\nДля следующего слова введи что-нибудь, Введёшь 0 и прервёшь повтор')
        if Next_word == '0':
            print('\nПовторение прерванно\n')
            break
        clear()
    Next_word = input('Закончить повтор')

def Search_Word_in_Slovar():
    Serching_Word = input('Введите слово для перевода: ')
    for idx in range(len(Slovar)):
        if Serching_Word == Slovar[idx][1]:
            print_One_Line_Word(idx)

def Delete_Word():
    Del_Number = int(input('Введите номер слова')) - 1
    Proverka = input('Для подтверждения удаления нажмите 1')
    if Proverka =='1':
        del Slovar[Del_Number]
        for idx in range(len(Slovar)):
            Slovar[idx-1][0] = idx
        Slovar[len(Slovar)-1][0] = len(Slovar)
        print('Удалено.')
    else: print('Удаление отменено')

def Start_Learning():
    '''Меню выбора обучения'''
    clear()
    Learning_Choise = input('Мой словарь ::Режим повторения::\n\nПовторить весь словарь, нажми 1\nПовторить 10 случайных слов, нажми 2\nПовторить 10 самых неповторяемых слов, нажми 3\nПропущенные буквы... ввести всё слово, нажми 4\nНаписать правильное слово. нажми 5\nВернуться в главное меню, нажми 0\n')
    while Learning_Choise != '0':
        if Learning_Choise == '1':
            Repeat_all_word()
        elif Learning_Choise == '2':
            Repeat_10_random_word()
        elif Learning_Choise == '3':
            Repeat_10_LowRepeat_word()
        elif Learning_Choise == '4':
            Enter_True_Letter_word()
        elif Learning_Choise == '5':
            Enter_True_word()
        else: print('\nВведён неверный запрос\n')
        clear()
        Learning_Choise = input('Мой словарь ::Режим повторения::\n\nПовторить весь словарь, нажми 1\nПовторить 10 случайных слов, нажми 2\nПовторить 10 самых неповторяемых слов, нажми 3\nПропущенные буквы... ввести всё слово, нажми 4\nНаписать правильное слово. нажми 5\nВернуться в главное меню, нажми 0\n')

def Start_Dictnionary():
    '''Меню выбора действий'''
    Start_Choise = input('Если хочешь добавить слово, нажми 1\nЕсли хочешь повторить слова нажми 2\nЕсли хочешь вывести весь список, нажми 3\nЕсли хочешь ввести 10 слов по плану, нажми 4\nПромежуточное сохранение, нажми 5\nЕсли хочешь Удалить слово, нажми 6\nЕсли хочешь Найти слово, нажми 7\nЕсли хочешь закрыть программу нажми 0\n')
    while Start_Choise != '0':
        if Start_Choise == '1':
            Add_New_word()
        elif Start_Choise == '2':
            Start_Learning()
        elif Start_Choise == '3':
            Step_print_word()
        elif Start_Choise == '4':
            for idx in range(10):
                Add_New_word()
        elif Start_Choise == '5':
            Safe_Dictionary_Exel_DB()
            clear()
            print('Сохранено')
        elif Start_Choise == '6':
            Delete_Word()
        elif Start_Choise == '7':
            Search_Word_in_Slovar()
        else: print('\nВведён неверный запрос\n')
        input('\nВ главное меню...\n')
        clear()
        Start_Choise = input('Это мой словарь\n\nЕсли хочешь добавить слово, нажми 1\nЕсли хочешь повторить слова нажми 2\nЕсли хочешь вывести весь список, нажми 3\nЕсли хочешь ввести 10 слов по плану, нажми 4\nПромежуточное сохранение, нажми 5\nЕсли хочешь Удалить слово, нажми 6\nЕсли хочешь Найти слово, нажми 7\nЕсли хочешь закрыть программу нажми 0\n')

def Out_Dictionary(Step_Out, Speed_Out, Clear_Screen_Out):
    '''Корректное завершение работы программы'''
    Safe_Dictionary_Exel_DB()
    for idx in range(Step_Out):
        print('\nЗавершение работы словаря... ' + str(Step_Out))
        Step_Out = int(Step_Out) - 1
        time.sleep(Speed_Out)
        if Clear_Screen_Out:
            clear()

def Dictionary():
    print('Это мой словарь\n')
    Start_Dictnionary()
    clear()
    Out_Dictionary(0, 0.5, True)

def File_Step_write_word():
    '''Печать на экран новое следующее слово с новой строки'''
    clear()
    for idx in range(len(Slovar)):
        print(Slovar[idx])

def Open_Exel_DB():
    '''Открывает базу данных записанную в файле Excel'''
    # открываем файл
    f = xlrd.open_workbook('DB.xls', formatting_info=True)
    sheet = f.sheet_by_index(0)
    # получаем список значений из всех записей ( Если вместо sheet.nrows будет 1 - только первую строчку)
    vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    vals = list(vals)
    for idx in range(len(vals)):
        vals[idx][0] = int(vals[idx][0])
        vals[idx][4] = int(vals[idx][4])
        vals[idx][1] = First_big_Word(vals[idx][1])
        vals[idx][2] = First_big_Word(vals[idx][2])
        vals[idx][3] = First_big_Word(vals[idx][3])
    return vals

def Safe_Dictionary_Exel_DB():
    '''Сохранение всех данных в таблицу Excel'''
    for idx1 in range(len(Slovar)):
        for idx2 in range(4):
            wordl = str(Slovar[idx1][idx2])
            wordl = list(wordl)
            words = ''
            for idx3 in range(len(wordl)):
                words1 = str(wordl[idx3])
                words = words + words1.lower()
                Slovar[idx1][idx2] = words
    wb = xlwt.Workbook()
    sheet1 = wb.add_sheet('Sheet 1')
    for stolb in range(len(Slovar[0])):
        for stroki in range(len(Slovar)):
            row = Slovar[stroki]
            x = row[stolb]
            sheet1.write(stroki, stolb, x)
    wb.save('DB.xls')

Slovar = list(Open_Exel_DB())
Dictionary()
