import requests
# для работы с приложениями windows, в том числе с почтой Outlook 2013
import win32com.client
import time, datetime

def get_token_auth():
    """Функция для получения токена аутентификации при подключении к сайту"""
    # Передаваемые заголовки
    headers = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        'Connection': 'keep-alive',
        'Content-Length': '51',
        'Content-Type': 'application/json',
        'Host': 'ip:5000',
        'Origin': 'http://ip',
        'Referer': 'http://ip/login',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:70.0) Gecko/20100101 Firefox/70.0'
    }
    # Передаваемый в параметрах json
    json = {
        'username': '',
        'password': '',
    }

    # создаем подключение с нашими заголовками и параметрами
    response = requests.post('http://ip:5000/api/users/authenticate', headers=headers, json=json)

    if response.ok:
        print('OK')

    # Получаем токен ауентификации
    token_auth = response.json()['token']

    return token_auth


def get_token():
    """Функция для получения токена после аутентификации при подключении к сайту"""
    headers_token = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        'Authorization': 'Bearer ' + token_auth,
        'Connection': 'keep-alive',
        'Host': 'ip:5000',
        'Origin': 'http://ip',
        'Referer': 'http://ip/targetinfo',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:70.0) Gecko/20100101 Firefox/70.0'
    }

    response = requests.get('http://ip:5000/api/targetinfo/gettoken', headers=headers_token)

    # Получаем токен
    token = response.json()['token']

    return token


def send_request(auto_number):
    """Направление запроса ТОЛЬКО по ГРЗ"""
    auto_number = auto_number
    headers_request = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        'Authorization': 'Bearer ' + token_auth,
        'Connection': 'keep-alive',
        'Content-Length': '251',
        'Content-Type': 'application/json',
        'Host': 'ip:5000',
        'Origin': 'http://ip',
        'Referer': 'http://ip/targetinfo',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:70.0) Gecko/20100101 Firefox/70.0'
    }

    json_request = {
        'filters': {
            'devices': [],
            'directions': [],
            'emptyGrz': 'false',
            'firstTime': '',
            'isExact': 'false',
            'lastTime': '',
            'orderBy': 'DateFix_desc',
            'searchString': auto_number,
            'uniquePlate': 'false'},
        'neededPage': '0',
        'pageSize': '70',
        'token': token
    }

    request = requests.post('http://ip:5000/api/targetinfo/getpage', headers=headers_request,
                                 json=json_request)
    return request


def get_count_auto(auto_number):
    """Направление запроса ТОЛЬКО по ГРЗ"""
    auto_number = auto_number
    headers_request = {
        'Accept': 'application/json, text/plain, */*',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        'Authorization': 'Bearer ' + token_auth,
        'Connection': 'keep-alive',
        'Content-Length': '251',
        'Content-Type': 'application/json',
        'Host': 'ip:5000',
        'Origin': 'http://ip',
        'Referer': 'http://ip/targetinfo',
        'User-Agent': 'Mozilla/5.0 (Windows NT 6.1; Win64; x64; rv:70.0) Gecko/20100101 Firefox/70.0'
    }

    json_request = {
        'filters': {
            'devices': [],
            'directions': [],
            'emptyGrz': 'false',
            'firstTime': '',
            'isExact': 'false',
            'lastTime': '',
            'orderBy': 'DateFix_desc',
            'searchString': auto_number,
            'uniquePlate': 'false'},
        'neededPage': '0',
        'pageSize': '70',
        'token': token
    }

    request = requests.post('http://ip:5000/api/targetinfo/getcount', headers=headers_request,
                            json=json_request)
    request_count_auto = int(request.text)
    return request_count_auto


def processing_request(request):
    """Функция по обработке результатов запроса. Возвращает список"""
    request = request
    i = 0
    my_list = []
    print(request_count_auto)
    if request_count_auto > 0:
        while i <= request_count_auto and i <= 20:
            try:
                ymd = request.json()['items'][i]['dateFix']
                # Проверяем направление движения автомашины (если 1 - встречное, если 2 - попутное)
                if str(request.json()['items'][i]['direction']) == '1':
                    direction = 'встречное'
                else:
                    direction = 'попутное'

                my_list.append(
                    '[Дата] ' + str(datetime.datetime.strptime(
                        f'{ymd[:4]}-{ymd[5:7]}-{ymd[8:10]}T{ymd[11:13]}:{ymd[14:16]}:{ymd[17:19]}+0000',
                        '%Y-%m-%dT%H:%M:%S%z').astimezone())[:19] + '\n' +
                    '[Место] ' + request.json()['items'][i]['location'] + '\n' +
                    '[Направление] ' + direction + '\n' + '-' * 20)
            except IndexError:
                my_list.append('ИСКЛЮЧЕНИЕ: Сработало штатное исключение при доступе '
                               'к индексу, который больше, чем количество автомашин в БД')
            i = i + 1
    else:
        my_list.append('СООБЩЕНИЕ: Авто на территории Брянской области не зафиксировано')
    return my_list


def sentReply(to_address, subject):
    """Функция для отправки письма через электронную почту"""
    # инициализируем объект outlook
    # outlook = win32com.client.Dispatch("Outlook.Application")
    Msg = outlook.CreateItem(0)
    # формируем письма, выставляя адресата, тему и текст
    Msg.To = to_address
    Msg.Subject = subject  # название темы
    Msg.Body = result
    # и отправляем
    Msg.Send()


# перерыв в секундах между повторным сканирование почты
IterationDelay = 5

# Создаем бесконечный цикл проверки входящей почты
while True:
    outlook = win32com.client.Dispatch("Outlook.Application")
    mapi = outlook.GetNamespace("MAPI")
    my_folder = mapi.Folders['mail@mail.ru'].Folders['Входящие']

    for massage in my_folder.Items:
        # получаем электронный адрес автора входящего письма
        author = []
        author.append(massage.Sender.Address)
        # получаем название темы входящего письма
        head = []
        # переводим название темы в формат string
        head_fixed_string = str(massage.Subject)
        # добавляем название темы в список head и приводим название к нижнему регистру
        head.append(head_fixed_string.lower())
        # создаем список номеров автомашин и добавляем номер из тела входящего сообщения
        input_auto_number = []
        input_auto_number.append(massage.Body)

        # Сверяем все названия тем входящих сообщений на наличие слова 'поток'
        if ('поток') in head:
            # Сохраняем полученный список в строку и заменяем символы пробела (объеденяем).
            text = ''.join(input_auto_number).replace(' ', '')
            author_result = ''.join(author)
            # удаляем указанное сообщение
            massage.Delete()
            # Подключаемся к сайту "Поток"
            # получаем ключ аутентификации
            token_auth = get_token_auth()
            # получаем ключ после аутентификации
            token = get_token()
            # Получаем ответ на запрос по определнному номеру авто
            request = send_request(text)
            request_count_auto = get_count_auto(text)

            print(request.text)

            # Получаем результат в виде списка и переводим его в тип STRING
            result = '\n'.join(processing_request(request))
            print(result)

            sentReply(author_result, 'привет...')

    time.sleep(IterationDelay)
