import random
from tkinter import *
from tkinter import messagebox
import openpyxl
import tkinter as tk
import matplotlib

matplotlib.use('TkAgg')

from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import (
    FigureCanvasTkAgg,
    NavigationToolbar2Tk
)

letters = {'газпром': 'B', 'татнефть': 'C', 'сбербанк': 'D', 'втб': 'E', 'алроса': 'F', 'аэрофлот': 'G',
           'русгидро': 'H',
           'московская биржа': 'I', 'нлмк': 'J', 'северсталь': 'K', 'детский мир': 'L', 'полиметалл': 'M',
           'яндекс': 'N',
           'афк': 'O', 'группа лср': 'P', 'ленэнерго': 'Q', 'лукойл': 'R', 'мтс': 'S', 'новатэк': 'T',
           'пик': 'U'}


def excel(letter):
    values = []
    wb = openpyxl.load_workbook('data.xlsx')
    sheet = wb['A1']

    for i in range(2, 1012):
        value = sheet['%s%d' % (letter, i)].value
        values.append(value)

    return values


def vinzo(values):
    for i in range(len(values)):
        if values[i] is None:
            if i < len(values) - 1:
                g = i
                while values[g] is None:
                    if g == len(values) - 1:
                        g = i
                        while values[g] is None:
                            g -= 1
                        else:
                            values[i] = values[g]
                    else:
                        g += 1
                else:
                    values[i] = values[g]
            else:
                g = i
                while values[g] is None:
                    g -= 1
                else:
                    values[i] = values[g]

    return values


def lin_app(values):
    for i in range(len(values)):
        if values[i] is None:

            if len(values) - 1 > i > 0:
                g = i
                while values[g] is None:
                    if g == len(values) - 1:
                        values[i] = values[i - 1]
                        break
                    else:
                        g += 1
                else:
                    koef_x = (values[g] - values[i - 1]) / (g - i + 1)
                    for j in range(i, g):
                        values[j] = values[j - 1] + koef_x

            elif i == 0:
                g = i
                while values[g] is None:
                    g += 1
                else:
                    values[i] = values[g]
            else:
                values[i] = values[i - 1]

    return values


def kor_voss(values):
    only_first = []
    only_second = []
    for i in range(len(values)):
        if values[i] is None:
            continue
        else:
            only_first.append(i)
            only_second.append(values[i])
    sredn_znach1, sredn_znach2 = sum(only_first) / len(only_first), sum(only_second) / len(only_second)
    otklonenie1, otklonenie2 = 0, 0
    for g in range(len(only_first)):
        otklonenie1 += (float(only_first[g]) - sredn_znach1) ** 2
        otklonenie2 += (float(only_second[g]) - sredn_znach2) ** 2
    srednee_kvad_otkl1 = (otklonenie1 / (len(only_first))) ** (1 / 2)
    srednee_kvad_otkl2 = (otklonenie2 / (len(only_second))) ** (1 / 2)
    umnozh = 0
    for l in range(len(only_first)):
        umnozh += (float(only_first[l]) * float(only_second[l]))

    kor = ((umnozh / len(only_first)) - (sredn_znach1 * sredn_znach2)) / (
            srednee_kvad_otkl1 * srednee_kvad_otkl2)

    for i in range(len(values)):
        if values[i] is None:

            if len(values) - 1 > i > 0:
                g = i
                while values[g] is None:
                    if g == len(values) - 1:
                        values[i] = values[i - 1]
                        break
                    else:
                        g += 1
                else:
                    koef_x = (values[g] - values[i - 1]) / (g - i + 1)
                    for j in range(i, g):
                        values[j] = (values[j - 1] + koef_x) * kor

            elif i == 0:
                g = i
                while values[g] is None:
                    g += 1
                else:
                    values[i] = values[g]
            else:
                values[i] = values[i - 1]

    return values


def koef_kor_pre(go, tick):
    dates = excel('A')
    h = 1
    s = 1
    m, k = go.split()

    while (dates[-h].split('.')[1] + '.' + dates[-h].split('.')[2]) != str(m):
        h += 1
    while (dates[-s].split('.')[1] + '.' + dates[-s].split('.')[2]) != str(k):
        s += 1

    values = []
    prices = excel(str(letters[str(tick).lower()]))

    while h <= s:
        values.append(prices[-h])
        h += 1

    for i in range(len(values) // 3):
        values.insert(random.randint(0, len(values)), None)

    return values


def koef_kor(go, tick, types):
    dates = excel('A')
    h = 1
    s = 1
    m, k = go.split()

    while (dates[-h].split('.')[1] + '.' + dates[-h].split('.')[2]) != str(m):
        h += 1
    while (dates[-s].split('.')[1] + '.' + dates[-s].split('.')[2]) != str(k):
        s += 1

    values = []
    prices = excel(str(letters[str(tick).lower()]))

    while h <= s:
        values.append(prices[-h])
        h += 1

    for i in range(len(values) // 3):
        values.insert(random.randint(0, len(values)), None)

    if str(types).lower() == '1':
        values = vinzo(values)
    elif str(types).lower() == '2':
        values = lin_app(values)
    elif str(types).lower() == '3':
        values = kor_voss(values)

    return values


def sglazh(values, type, window):
    if str(type).lower() == '1':
        n = 1
        result = []
        summa = [values[0]]
        while n < len(values):
            if n % 3 == 0:
                result.append((summa[0] + summa[1] * 2 + summa[2]) / 4)
                summa = [values[n]]
                n += 1
            else:
                summa.append(values[n])
                n += 1
        else:
            result.append(sum(summa) / len(summa))
    elif str(type).lower() == '2':
        n = 1
        result = []
        summa = [values[0]]
        while n < len(values):
            if n % int(window) == 0:
                result.append(sum(summa) / int(window))
                summa = [values[n]]
                n += 1
            else:
                summa.append(values[n])
                n += 1
        else:
            result.append(sum(summa)/int(window))
    return result


def start_pre():
    ticker_start = ticker.get()
    time_start = time.get()
    values = koef_kor_pre(time_start, ticker_start)
    messagebox.showinfo(title='Значения перед восстановлением', message=values)


root = Tk()
root['bg'] = '#2F4F4F'
root.geometry('1500x1500')
root.title('Анализ акций для инвестирования by Daniil')

frame = Frame(root, bg='#293133', bd=5)
frame.place(relwidth=0.9, relheight=0.9, relx=0.05, rely=0.05)

label = Label(frame, text='Выберите тикер акции:',
              bg='#293133',
              fg='white')
label.config(font=("Courier", 20))
label.pack()

label_dop = Label(frame, text=(
        '(Газпром, Татнефть, Сбербанк, ВТБ, Алроса, Аэрофлот, РусГидро, Московская Биржа, НЛМК,' + '\n' + 'Северсталь, Детский Мир, Полиметалл, Яндекс, АФК, Система, Группа ЛСР, Ленэнерго,' + '\n' + 'Лукойл, ''МТС, Новатэк и ПИК)'),
                  bg='#293133',
                  fg='white', )
label_dop.config(font=("Courier", 10))
label_dop.pack()

ticker = Entry(frame, bg='white')
ticker.pack()

space = Frame(frame, bg='#293133', bd=5, width=10, height=50)
space.pack()

label1 = Label(frame, text='Выберите временной промежуток:',
               bg='#293133',
               fg='white')
label1.config(font=("Courier", 20))
label1.pack()

label_dop1 = Label(frame, text=('(01.2016 12.2019)'),
                   bg='#293133',
                   fg='white', )
label_dop1.config(font=("Courier", 10))
label_dop1.pack()

time = Entry(frame, bg='white')
time.pack()

space = Frame(frame, bg='#293133', bd=5, width=10, height=50)
space.pack()

label2 = Label(frame, text='Выберите метод восстановления пропущенных данных: (1, 2, 3)',
               bg='#293133',
               fg='white')
label2.config(font=("Courier", 20))
label2.pack()

label_dop2 = Label(frame, text='(винзорирование, линейная аппроксимация, корреляционное восстановление)',
                   bg='#293133',
                   fg='white', )
label_dop2.config(font=("Courier", 10))
label_dop2.pack()

types = Entry(frame, bg='white')
types.pack()

space = Frame(frame, bg='#293133', bd=5, width=10, height=50)
space.pack()

label3 = Label(frame, text='Выберите метод сглаживания данных: (1, 2)',
               bg='#293133',
               fg='white')
label3.config(font=("Courier", 20))
label3.pack()

label_dop3 = Label(frame, text='(взвешенный метод скользящего среднего, метод скользящего среднего со скользящим '
                               'окном наблюдения)',
                   bg='#293133',
                   fg='white', )
label_dop3.config(font=("Courier", 10))
label_dop3.pack()

types_1 = Entry(frame, bg='white')
types_1.pack()

types_2 = Entry(frame, bg='white')
types_2.pack()

space = Frame(frame, bg='#293133', bd=5, width=10, height=50)
space.pack()

btn = Button(frame, text='Build pre data', bg='yellow', command=start_pre)
btn.pack()

space = Frame(frame, bg='#293133', bd=5, width=10, height=50)
space.pack()


class App(tk.Tk):
    def __init__(self):
        super().__init__()

        self.title('App')

        # prepare data
        ticker_start = ticker.get()
        time_start = time.get()
        types_start = types.get()
        types_1_start = types_1.get()
        window = types_2.get()
        values = koef_kor(time_start, ticker_start, types_start)
        nice = values
        values = sglazh(values, types_1_start, window)
        first = nice
        second = [i for i in range(len(nice))]

        # create a figure
        figure = Figure(figsize=(6, 4), dpi=100)

        # create FigureCanvasTkAgg object
        figure_canvas = FigureCanvasTkAgg(figure, self)

        # create the toolbar
        NavigationToolbar2Tk(figure_canvas, self)

        # create axes
        axes = figure.add_subplot()

        # create the barchart
        axes.plot(second, first)
        if types_1_start == '1':
            axes.plot([i*3 for i in range(len(values))], values)
        else:
            axes.plot([i*int(window) for i in range(len(values))], values)
        print([i*3 for i in range(len(values))], [i for i in range(len(nice))])
        axes.set_title('Акции')
        axes.set_ylabel('Цена')

        figure_canvas.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=1)


btn_result = Button(frame, text='Build', bg='yellow', command=App)
btn_result.pack()

root.mainloop()
