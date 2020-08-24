from bs4 import BeautifulSoup #подключение библиотеки для поиска по html
import requests #подключение библиотеки для отправки и получения запросов
from openpyxl import load_workbook #подключение библиотеки для работы с excel
import re #подключение библиотеки для работы с текстом
from urllib import request

#объявление переменных для ИНН, названия компании, номера телефона и электронной почты
capcha = ''
HEADERS = {}
PHONEBD = ''
INN = ''
NAME = ''
PHONE = ''
EMAIL = ''
URL = ''
FIRSTURL = ''
#обращение к базе данных excel
wbopen = load_workbook(filename='finblock.xlsx', data_only=True)
sheet = wbopen['Лист1']

def get_html(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params)
    return r

#первое сравнение данных по поиску ИНН
def get_content_first(html):
    soup = BeautifulSoup(html, 'html.parser')

    items = soup.find_all('div', class_='content')
    link = {'link': ''}

    for item in items:
        try:
            link.update({
                'link': item.find('div', class_='org_list').find('a').get('href')
            })

        except: #если поиск не дает результатов, печатается 'Не найдено'
            try:
                images = soup.findAll('img') #поиск img в html страницы
                for image in images:
                    print(image['src']) #поиск scr для формирования ссылки капчи
                    global URL, capcha
                    capcha = URL + image['src'] #формируем ссылку на капчу
                    print('Возможно, капча! Скачиваем файл..')
                    request.urlretrieve(capcha, 'out.jpg') #скачиваем капчу
                    print('Файл скачан')
                    s = requests.Session()
                    #указываем шапку для запроса к странице капчи
                    HEADERBOT = {
                        'accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
                        'accept-encoding': 'gzip, deflate, br',
                        'accept-language': 'ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7',
                        'cookie': '_ga=GA1.2.357585501.1597317931; __gads=ID=f9937b7742642e25:T=1597318362:S=ALNI_Maw63MLAk48soPvkxBhcQsnb2nPLg; user=5f353cbb6110c2060267; _gid=GA1.2.1294085579.1598102037; PHPSESSID=6p7sh35auj8t0be4fm4g90fk3r',
                        'sec-fetch-dest': 'document',
                        'sec-fetch-mode': 'navigate',
                        'sec-fetch-site': 'same-origin',
                        'sec-fetch-user': '?1',
                        'upgrade-insecure-requests': '1',
                        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.135 Safari/537.36'
                    }
                    s.headers = HEADERBOT
                    s.get('https://www.list-org.com')

                    #сбор данных Form Data для отправки на сайт
                    data = {'keystring': '',
                            'submit': ' Проверить! '
                    }

                    incapcha = input('Введите код с картинки: ')
                    data['keystring'] = incapcha
                    z = s.post('https://www.list-org.com/bot', data=data) #отправка данных на сайт

                    return z

            except: #в ином случае выведет соответствующее сообщение
                print('Не найдено')

#получение ответа поиска по сайту
def open_first(FIRSTURL):
    html = get_html(FIRSTURL)
    if html.status_code == 200:
        get_content_first(html.text)
    else:
        print('Ошибка')

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

def open_second(URL):
    html = get_html(URL)
    if html.status_code == 200:
        get_content_second(html.text)
    else:
        print('Ошибка')

#цикл, который беребирает ИНН и дает возможность выбрать диапазон поиска
def cycle():
    q = int(input('С какой строки начать проверку? '))
    o = input('До какого числа хотите проверить? ')
    while q < int(o):

        q = str(q)

        form = sheet['A'+q].value
        print(form)

        global INN, NAME, PHONE, PHONEBD, EMAIL, HEADERS, FIRSTURL, URL
        FIRSTURL = 'https://www.list-org.com/search?type=inn&val=' + form  # FIRSTURL для сравнения ИНН по поиску
        URL = 'https://www.list-org.com'  # URL для дальнейшей работы с ссылками
        HEADERS = {'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.135 Safari/537.36',
                   'accept': '*/*'}  # заголовки для того, чтобы посылаемый запрос выдавал нас за реального пользователя

        print(HEADERS)
        print(FIRSTURL)

        open_first(FIRSTURL)
        open_second(URL)
        # приводим данные к строковому виду
        INN = str(INN)
        PHONE = str(PHONE)
        NAME = str(NAME)
        EMAIL = str(EMAIL)
        # убираем лишние символы
        INN = re.sub(r"[ИНН:()': ]", "", INN)
        PHONE = re.sub(r"[()', +-]", "", PHONE)
        NAME = re.sub(r"[()':]", "", NAME)
        EMAIL = re.sub(r"[()': ]", "", EMAIL)
        # записываем телефон и email напротив исследуемого ИНН
        NEXT = int(q) + 1
        NEXT = str(NEXT)
        sheet['G' + NEXT].value = PHONE
        sheet['H' + NEXT].value = EMAIL
        PHONEBD = sheet['E' + str(q)].value
        PHONEBD = re.sub(r"[()', +-]", "", PHONEBD)


        if PHONE == 'Пусто':  # проверка на наличие номера на list org
            sheet['J' + NEXT] = 'Номер не найден'
        else:  # если номер есть, сверяется номер с базой данных и данными с list org
            if PHONEBD[6:10] == PHONE[6:10]:
                sheet['J' + NEXT] = 'Одинаковый номер'
            else:
                sheet['J' + NEXT] = 'Номер изменен'
        q = int(q) + 1
        print(q)
        print('Далее.. ')
cycle()
wbopen.save('gotovo.xlsx') #записываем данные в новый файл excel
print('Готово!')