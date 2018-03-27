import time
import xlrd

def Open_Exel_DB():
    '''Открывает базу данных записанную в файле Excel'''
    # открываем файл
    f = xlrd.open_workbook('DB.xls', formatting_info=True)
    sheet = f.sheet_by_index(0)
    # получаем список значений из всех записей ( Если вместо sheet.nrows будет 1 - только первую строчку)
    vals = [sheet.row_values(rownum) for rownum in range(sheet.nrows)]
    vals = list(vals)
    return vals

def Level(words):
    if words < 2000:
        level = 'A1 Beginner'
    elif words < 3000:
        level = 'A1 Elementary'
    elif words < 4000:
        level = 'A2 Pre-Intermediate'
    elif words < 6000:
        level = 'B1 Intermediate'
    elif words < 8000:
        level = 'B2 Upper-Intermediate'
    elif words < 12000:
        level = 'C1 Advanced'
    else: level = 'C2 Proficient'
    return level

def LevelUp(words):
    if words < 2000:
        levelup = 'До уровня A1 Elementary, осталось {0} слов, или {1:2.2f}% от требуемого'.format(2000 - words, words/2000)
    elif words < 3000:
        levelup = 'До уровня A2 Pre-Intermediate, осталось {0} слов, или {1:2.2f}% от требуемого'.format(3000 - words, words/3000)
    elif words < 4000:
        levelup = 'До уровня B1 Intermediate, осталось {0} слов, или {1:2.2f}% от требуемого'.format(4000 - words, words/4000)
    elif words < 6000:
        levelup = 'До уровня B2 Upper-Intermediate, осталось {0} слов, или {1:2.2f}% от требуемого'.format(6000 - words, words/6000)
    elif words < 8000:
        levelup = 'До уровня C1 Advanced, осталось {0} слов, или {1:2.2f}% от требуемого'.format(8000 - words, words/8000)
    elif words < 12000:
        levelup = 'До уровня C2 Proficient, осталось {0} слов, или {1:2.2f}% от требуемого'.format(12000 - words, words/12000)
    else: levelup = 'Ты C2 Proficient, и этим всё сказанно... Учи новый язык или не парься'
    return levelup

DATES = Open_Exel_DB()
for idx in range(len(DATES)):
    DATES[idx] = int(DATES[idx][4])

Day = min(DATES)
Kol = 0
vremR = []
only_val = []
DayKol = (int(time.strftime("%Y%m%d", time.gmtime())) - min(DATES) + 1)
for idxD in range(DayKol):
    Day = min(DATES) + idxD
    lstt = []
    for idx in range(len(DATES)):
        val = DATES[idx]
        if val == Day:
            Kol = Kol + 1
    lstt.append(idxD + 1)
    lstt.append(Kol)
    only_val.append(Kol)
    vremR.append(lstt)
    Kol = 0

TIME = int(input('Введите на сколько дней рассчитывать будущее значение:  '))

x = sum(only_val)
sred = x/len(only_val)
mes = int(sred * TIME + x)
print('Через {0} дней из {1} слов выйдет {2} слова|   По среднему за всё время Оптимистичный прогноз                         |    В среднем за день: {3:-2.2f} |'.format(TIME, x, mes, sred))

lv_Optim = mes

y = sum(only_val[-7:])
sred7 = y / 7
mes = int(sred7 * TIME + y)
print('Через {0} дней из {1} слов выйдет {2} слова|   По среднему за последнюю неделю Моментный прогноз                      |    В среднем за день: {3:-2.2f} |'.format(TIME, x, mes, sred7))

lv_Moment = mes

for idx in range(TIME):
    only_val.append(0)
    Predikt = 0
    sum_Predikt = 0
    XXX = 0
    WeekBefor = int(len(only_val) / 7 + 0.5)
    Srednee = sum(only_val) / len(only_val)
    for idxIn in range(WeekBefor + 1):
        IdxIn = 7 * ( 1 + idxIn )
        iiddxx = len(only_val) - 1 - IdxIn
        xxx = int(only_val[iiddxx])
        sum_Predikt = sum_Predikt + xxx
        XXX = XXX + 1
    Predikt = int((sum_Predikt / XXX) + 0.5)
    only_val[len(only_val) - 1] = Predikt
mes = sum(only_val)
sred = sum(only_val)/len(only_val)
print('Через {0} дней из {1} слов выйдет {2} слова|   По среднему за єтот день недели Динамический прогноз                   |    В среднем за день: {3:-2.2f} |'.format(TIME, x, mes, sred))

lv_Din = mes

y = sum(only_val[-7:])
sred7 = y / 7
mes = int(sred7 * TIME + y)
print('Через {0} дней из {1} слов выйдет {2} слова|   По среднему за последнюю неделю динамического прогноза Реальный прогноз|    В среднем за день: {3:-2.2f} |'.format(TIME, x, mes, sred7))

lv_Real = mes

z = input('Нажмите ввод, что бы перейти далее.')

print('''Уровень владения Английским языком:
Сейчас = {1},   {2}
Через {0} дней. По Оптимистичный прогноз = {3},   {4}
Через {0} дней. По Моментный прогноз = {5},   {6}
Через {0} дней. По Динамический прогноз = {7},   {8}
Через {0} дней. По Реальный прогноз = {9},   {10}
'''.format(TIME, Level(x), LevelUp(x), Level(lv_Optim), LevelUp(lv_Optim), Level(lv_Moment), LevelUp(lv_Moment), Level(lv_Din), LevelUp(lv_Din), Level(lv_Real), LevelUp(lv_Real)))


z = input('Нажмите ввод, для выхода.')