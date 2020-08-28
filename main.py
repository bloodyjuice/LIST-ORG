from bs4 import BeautifulSoup  # подключение библиотеки для поиска по html
import requests  # подключение библиотеки для отправки и получения запросов
from openpyxl import load_workbook  # подключение библиотеки для работы с excel
import re  # подключение библиотеки для работы с текстом
from urllib import request


def get_html(url, params=None):
    r = requests.get(url, headers=HEADERS, params=params)
    return r


# первое сравнение данных по поиску ИНН
def get_content_first(html):  # URL для дальнейшей работы с ссылками
    soup = BeautifulSoup(html, 'html.parser')

    items = soup.find_all('div', class_='content')
    link = {'link': None}
    URL = 'https://www.list-org.com'  # URL для дальнейшей работы с ссылками
    for item in items:

        try:
            link.update({
                'link': item.find('div', class_='org_list').find('a').get('href')
            })
            URL = URL + link['link']
            print(URL)
            html2 = get_html(URL)
            if html2.status_code == 200:
                get_content_second(html2.text)
            else:
                print('Ошибка 02')
        except AttributeError:  # проверяется на наличие капчи, в ином случае печатается 'Не найдено'
            try:
                images = soup.findAll('img')  # поиск img в html страницы
                for image in images:  # поиск scr для формирования ссылки капчи
                    capcha = URL + image['src']  # формируем ссылку на капчу
                    print('Возможно, капча! Скачиваем файл..')
                    request.urlretrieve(capcha, 'out.jpg')  # скачиваем капчу
                    print('Файл скачан')
                    s = requests.Session()
                    # указываем шапку для запроса к странице капчи
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
                    # сбор данных Form Data для отправки на сайт
                    data = {'keystring': '',
                            'submit': ' Проверить! '
                            }
                    incapcha = input('Введите код с картинки: ')
                    data['keystring'] = incapcha
                    z = s.post('https://www.list-org.com/bot', data=data)  # отправка данных на сайт
                    return z
            except AttributeError:  # в ином случае выведет соответствующее сообщение
                print('Не найдено')


# если в поиске выдало результат, приступаем к проверке внутри найденной страницы
def get_content_second(html2):
    html = html2
    soup = BeautifulSoup(html, 'html.parser')
    items = soup.find_all('div', class_='content')
    for item in items:  # поиск данных в полученном html коде
        # объявляем переменные для телефона с БД, ИНН, названия компании, телефона и эл.почты
        # забираем указанные значения, в ином случае получаем 'Пусто'
        try:
            NAME = item.find('div', class_='c2m').find('a', class_='upper').get_text(),
        except AttributeError:
            NAME = 'Пусто'
        try:
            PHONE = item.find('div', class_='c2m').findNext('div', class_='c2m').find('a',
                                                                                      class_='nwra lbs64').get_text(),
        except AttributeError:
            PHONE = 'Пусто'
        try:
            INN = item.find('div', class_='c2m').findNext('div', class_='c2m').findNext('div', class_='c2m').find(
                'p').get_text(),
        except AttributeError:
            INN = 'Пусто'

        try:
            EMAIL = item.find('div', class_='c2m').findNext('div', class_='c2m').find('a',
                                                                                      rel='nofollow').get_text()
        except AttributeError:
            EMAIL = 'Пусто'

        global next, sq
        # убираем лишние символы
        INN = re.sub(r"[ИНН:()': ]", "", str(INN))
        PHONE = re.sub(r"[()', +-]", "", str(PHONE))
        NAME = re.sub(r"[()':]", "", str(NAME))
        EMAIL = re.sub(r"[()': ]", "", str(EMAIL))
        # записываем телефон и email напротив исследуемого ИНН
        next = str(next)
        sheet[f'G{next}'].value = PHONE
        sheet[f'H{next}'].value = EMAIL
        PHONEBD = sheet[f'E{sq}'].value
        PHONEBD = re.sub(r"[()', +-]", "", PHONEBD)

        if PHONE == 'Пусто':  # проверка на наличие номера на list org
            sheet['J' + next] = 'Номер не найден'
        else:  # если номер есть, сверяется номер с базой данных и данными с list org
            if PHONEBD[6:10] == PHONE[6:10]:
                sheet['J' + next] = 'Одинаковый номер'
            else:
                sheet['J' + next] = 'Номер изменен'


def open(FIRSTURL):
    html = get_html(FIRSTURL)
    if html.status_code == 200:
        get_content_first(html.text)
    else:
        print('Ошибка 01')


if __name__ == '__main__':
    # обращение к базе данных excel
    wbopen = load_workbook(filename='finblock.xlsx', data_only=True)
    sheet = wbopen['Лист1']
    HEADERS = {
        'user-agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/84.0.4147.135 Safari/537.36',
        'accept': '*/*'}  # заголовки для того, чтобы посылаемый запрос выдавал нас за реального пользователя
    try:
        q = int(input('С какой строки начать проверку? '))
    except ValueError:
        q = 5
        print('Введено не число! Используется значение по умолчанию: ', q)
    try:
        o = int(input('До какого числа хотите проверить? '))
    except ValueError:
        o = q + 50
        print('Введено не число! Используется значение по умолчанию: ', o)

    # цикл, который беребирает ИНН и дает возможность выбрать диапазон поиска
    while q < o:
        next = q + 1
        sq = str(q)
        form = sheet[f'A{sq}'].value
        FIRSTURL = 'https://www.list-org.com/search?type=inn&val=' + form  # FIRSTURL для сравнения ИНН по поиску
        open(FIRSTURL)
        q = q + 1
        print('Далее.. ')

    wbopen.save('gotovo.xlsx')  # записываем данные в новый файл excel
    print('Готово!')