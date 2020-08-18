from bs4 import BeautifulSoup #подключение библиотеки для поиска по html
import requests #подключение библиотеки для отправки и получения запросов
from openpyxl import load_workbook #подключение библиотеки для работы с excel
import re #подключение библиотеки для работы с текстом
from fake_useragent import UserAgent
from random import choice
q = 5 #стартовое значение
o = input('До какой записи хотите проверить? ') #конечное значение (по какую строку считывать данные)
#объявление переменных для ИНН, названия компании, номера телефона и электронной почты
PHONEBD = ''
INN = ''
NAME = ''
PHONE = ''
EMAIL = ''
#обращение к базе данных excel
wbopen = load_workbook(filename='finblock.xlsx', data_only=True)
sheet = wbopen['Лист1']

while int(q)<int(o): #цикл от стартового до конечного значения(радиус поиска)
    try:
        q = str(q)
        ua = UserAgent().random

        form = sheet['A'+q].value
        proxies = open('proxxx.txt').read().split('\n')
        for i in range(10):
            proxy = {'http': 'http://' + choice(proxies)}
            print(proxy)
        print (form)
        FIRSTURL = 'https://www.list-org.com/search?type=inn&val='+form #FIRSTURL для сравнения ИНН по поиску
        URL='https://www.list-org.com' #URL для дальнейшей работы с ссылками
        print(FIRSTURL)

        HEADERS = {'user-agent': ua,
                   'accept': '*/*'} #заголовки для того, чтобы посылаемый запрос выдавал нас за реального пользователя
        print(HEADERS)
        #подключение запросов
        def get_html(url, params=None, proxy=None):
            r = requests.get(url, headers=HEADERS, params=params, proxies=proxy)
            return r

        #первое сравнение данных по поиску ИНН
        def get_content_first(html):
            soup = BeautifulSoup(html, 'html.parser')
            items = soup.find_all('div', class_='content')
            link = {}
            for item in items:
                try:
                    link.update({
                        'link': item.find('div', class_='org_list').find('a').get('href')
                    })

                except: #если поиск не дает результатов, печатается 'Не найдено'
                    print('Не найдено')


            print(link['link'])

            global URL
            URL = URL+link['link']
            print(link['link'])
        #получение ответа поиска по сайту
        def open_first():
            html = get_html(FIRSTURL)
            if html.status_code == 200:
                get_content_first(html.text)
            else:
                print('Ошибка')

        open_first()

        #если в поиске выдало результат, приступаем к проверке внутри найденной страницы
        def get_content_second(html):
            soup = BeautifulSoup(html, 'html.parser')
            items = soup.find_all('div', class_='content')
            for item in items:
                #забираем указанные значения, в ином случае получаем 'Пусто'
                global INN, NAME, PHONE, EMAIL
                try:
                    NAME = item.find('div', class_='c2m').find('a', class_='upper').get_text(),
                except:
                    NAME = 'Пусто'
                try:
                    PHONE = item.find('div', class_='c2m').findNext('div', class_='c2m').find('a',
                                                                                               class_='nwra lbs64').get_text(),
                except:
                    PHONE = 'Пусто'
                try:
                    INN = item.find('div', class_='c2m').findNext('div', class_='c2m').findNext('div', class_='c2m').find(
                        'p').get_text(),
                except:
                    INN = 'Пусто'

                try:
                    EMAIL = item.find('div', class_='c2m').findNext('div', class_='c2m').find('a',
                                                                                                   rel='nofollow').get_text()
                except AttributeError:
                    EMAIL = 'Пусто'

        def open_second():
            html = get_html(URL)
            if html.status_code == 200:
                get_content_second(html.text)
            else:
                print('Ошибка')

        open_second()
        #приводим данные к строковому виду
        INN = str(INN)
        PHONE = str(PHONE)
        NAME = str(NAME)
        EMAIL = str(EMAIL)
        #убираем лишние символы
        INN = re.sub(r"[ИНН:()': ]", "", INN)
        PHONE = re.sub(r"[()', +-]", "", PHONE)
        NAME = re.sub(r"[()':]", "", NAME)
        EMAIL = re.sub(r"[()': ]", "", EMAIL)
        #записываем телефон и email напротив исследуемого ИНН
        sheet['G'+str(q)]=PHONE
        sheet['H'+str(q)]=EMAIL
        PHONEBD = sheet['E' + str(q)].value
        PHONEBD = re.sub(r"[()', +-]", "", PHONEBD)

        if PHONE == 'Пусто': #проверка на наличие номера на list org
            sheet['J' + str(q)] = 'Номер не найден'
        else: #если номер есть, сверяется номер с базой данных и данными с list org
            if PHONEBD[7:11] == PHONE[7:11]:
                sheet['J' + str(q)] = 'Одинаковый номер'
            else:
                sheet['J' + str(q)] = 'Номер изменен'

        q = q+1 #если всё успешно, приступаем к следующей строке

    except: #в случае пустого результата или иной ошибки, пропускаем исследуему строку базы данных и переходим к следующей
        q = int(q)
        q = q + 1
        print(q)
        print('Пропуск..')

wbopen.save('gotovo.xlsx') #записываем данные в новый файл excel
print('Готово!')