import logging
import subprocess
import os
import ctypes
import time
import wmi
from requests import get, post, put, delete, packages
from json import dumps
from datetime import datetime
import win32api
from base64 import b64decode


# Настройка логгирования
if not os.path.exists("c://soft"):
    os.mkdir("c://soft")
if not os.path.exists("c://soft//FINIT log"):
    os.mkdir("c://soft//FINIT log")
logging.basicConfig(level=logging.INFO, filename="c://soft//FINIT log//log.log", filemode="w",
                    format="%(levelname)s %(message)s", encoding="utf-8")


# Описание глобальных переменных
NOTE_INFO = {}
USERS_DICT = {'ahmadullin': 'Роман Ахмадуллин',
              'aleksskull': 'Александр Черепенин',
              'alex_adm': 'Александр Новосёлов',
              'ares': 'Михаил Табильский',
              'armoredfail': 'Павел Семёнов',
              'babikov': 'Дмитрий Бабиков',
              'cherepanov_m': 'Михаил Черепанов',
              'crash': 'Алексей Крашенинников',
              'dementev_e': 'Евгений Дементьев',
              'fortis': 'Сергей Фортис',
              'glazyrin': 'Евгений Глазырин',
              'glinsky': 'Евгений Глинских',
              'grenor': 'Алексей Щербаков',
              'haliullin': 'Эдуард Халиуллин',
              'ivanov_s': 'Семён Иванов',
              'iskakov': 'Халит Искаков',
              'barybin': 'Денис Барыбин',
              'panihin': 'Вячеслав Панихин',
              'rogatskiyi': 'Александр Рогацкий',
              'smoke': 'Алексей Марков',
              'usr3111': 'Гризель Денис',
              'usr3443': 'Александр Никулин'}
USER = ''
DSC = ''
DSC_2 = ''
VAULT_ID = "60f68e57917ddf485e648823"       # TEST BASE:            "654c7b281915343bf564137d"  REAL BASE:          "60f68e57917ddf485e648823"
OO_FOLDER = "641d9cdec31d255d93601453"     # TEST OO_FOLDER:       "654c7b3bfe1c4761b94d1b35"  REAL OO_FOLDER:     "641d9cdec31d255d93601453"
TOCHKA_FOLDER = "64181b053a00c740583fc0f4"  # TEST TOCHKA_FOLDER:   "654c7b331915343bf564139b"  REAL TOCHKA_FOLDER: "64181b053a00c740583fc0f4"
WN_FOLDER = "6559eec8fcf15c15624c8fa7"      # TEST WN_FOLDER:       "654c7b466d239062ee61dee5"  REAL WN_FOLDER:     "6559eec8fcf15c15624c8fa7"
URL = 'https://194.107.116.186/api/v4'
TOKEN = ''
KEY_LOCATION = [r'C:\API_KEY.txt', r'D:\API_KEY.txt', r'E:\API_KEY.txt', r'F:\API_KEY.txt', r'G:\API_KEY.txt', r'H:\API_KEY.txt', r'I:\API_KEY.txt', r'M:\API_KEY.txt', r'J:\API_KEY.txt', r'Z:\API_KEY.txt', r'X:\API_KEY.txt']
ADM_ID = '63e3782750492875340e93e4'


# Создание учетной записи Windows
def create_user_win(username, password, adm_rules):
    logging.info(f"create_user({username}, {password}, {adm_rules})")
    try:
        subprocess.run(f"net user {username} {password} /add", shell=True)
        if adm_rules:
            subprocess.run(f"net localgroup Администраторы {username} /add", shell=True)
        logging.info(f"Учетная запись {username} успешно создана!")
    except:
        logging.error("", exc_info=True)


# Проверка администратора
def is_admin():
    logging.info("is_admin()")
    try:
        # Проверяем что код запущен от имени администратора
        chk = ctypes.windll.shell32.IsUserAnAdmin()
        if not chk:
            logging.error("FINIT запущен без прав администратора!")
            print('\033[1m' + '\033[91m' + "FINIT запущен без прав администратора! Завершение работы через 5 секунд!" + '\033[0m')
            time.sleep(5)
            os._exit(0)
        logging.info("Скрипт запущен с правами администратора")

        # Проверяем что учетная запись Administrator введена верно
        output = subprocess.check_output('echo %USERNAME%', shell=True)
        current_user = output.decode("utf-8").strip()
        if current_user != 'Administrator':
            logging.error('Имя учетной записи введено неверно, требуется переустановка системы')
            print('\033[1m' + '\033[91m' + 'Имя учетной записи введено неверно, требуется переустановка системы! Завершение работы через 5 секунд!' + '\033[0m')
            time.sleep(5)
            os._exit(0)
        logging.info("Учетная запись Administrator введена верно")
    except:
        logging.error("", exc_info=True)


# Получение информации о системе и пользователе
def get_info():
    logging.info('get_info()')
    try:
        global NOTE_INFO
        # Получение серийного номера
        bios = wmi.WMI().Win32_BIOS()[0]
        NOTE_INFO['SN'] = bios.SerialNumber
        logging.info(f"S/N: {NOTE_INFO['SN']}")
        print('\033[1m' + '\033[93m' + f"S/N: {NOTE_INFO['SN']} \n" + '\033[0m')
        for model in wmi.WMI().win32_ComputerSystem():
            logging.info(f'MTM - {model.model}')
            NOTE_INFO['MTM'] = model.model
    except:
        logging.error("", exc_info=True)


# Переименование ПК
def rename_computer(new_computer_name):
    logging.info(f'rename_computer({new_computer_name})')
    try:
        c = wmi.WMI()
        computer = c.Win32_ComputerSystem()[0]
        result = computer.Rename(new_computer_name)
        logging.info(f"Имя компьютера успешно изменено на {new_computer_name}!")
    except:
        logging.error("", exc_info=True)


# Аутентификация
def auth():
    logging.info('auth()')
    input("Нажми Enter для начала работы ")
    try:
        action = '/auth/login/'
        key = ''
        for kl in KEY_LOCATION:
            try:
                with open(kl, 'r') as file:
                    logging.info(f'Ключ найден в {kl}')
                    key = file.read()
            except FileNotFoundError:
                None
        if key == '':
            logging.error("Файл ключа API passwork не найден!")
            print('\033[1m' + '\033[91m' + "Файл ключа API passwork не найден! Завершение работы через 5 секунд!" + '\033[0m')
            time.sleep(5)
            os._exit(0)
        r = post(URL + action + key, verify=False).json()
        if r['status'] == 'success':
            global TOKEN, USER, DSC, DSC_2
            TOKEN = r['data']['token']
            name = r['data']['user']['name']
            try:
                USER = USERS_DICT[name]
            except:
                USER = name
            print('\033[1m' + '\033[32m' + f"Успешная аутентификация! Здорова, {USER}! \n" + '\033[0m')
            DSC = f'Создано с помощью FINIT\n Кем: {USER}\n Когда: {get_time()}\n'
            DSC_2 = f'Перезалит с помощью FINIT\n Кем: {USER}\n Когда: {get_time()}\n'
            logging.info(f'Успешная аутентификация, {USER}')
        else:
            logging.error('Ошибка аутентификации, неверный ключ API')
            print('\033[1m' + '\033[91m' + "Ошибка аутентификации, неверный ключ API. Завершение работы через 5 секунд" + '\033[0m')
            time.sleep(5)
            os._exit(0)
    except:
        logging.error("", exc_info=True)
        print('\033[1m' + '\033[91m' + "Ошибка соеденения с сервером. Завершение работы через 5 секунд" + '\033[0m')
        time.sleep(5)
        os._exit(0)


# Получение времени (почему-то всегда отрабатывает только локальное время)
def get_time():
    logging.info('get_time()')
    try:
        r = get('https://www.timeapi.io/api/Time/current/zone?timeZone=Etc/UTC', verify=False).json()
        day_list = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
        i = 0
        for day in day_list:
            i += 1
            if r['dayOfWeek'] == day:
                dayOfWeek = i
        win32api.SetSystemTime(r['year'], r['month'], dayOfWeek, r['day'], r['hour'], r['minute'], r['seconds'], r['milliSeconds'])
        r['month'] = "{:0>2}".format(r['month']) #форматируем месяц до 05 например
        r['day'] = "{:0>2}".format(r['day']) #форматируем дату до 05 например
        logging.info(f"set system time: {r['time']} (UTC), date: {r['day']}.{r['month']}.{r['year']}")
        date = f"{r['time']}, {r['day']}.{r['month']}.{r['year']}"
        return date
    except:
        logging.error("", exc_info=True)
        logging.info('use local time')
        now = datetime.now()
        now = now.strftime('%H:%M %d.%m.%Y')
        return now


# Выход из системы
def logout():
    logging.info('logout()')
    try:
        action = '/auth/logout'
        head = {'Passwork-Auth': TOKEN}
        r = post(URL + action, headers=head, verify=False).json()
        logging.info('Сессия FINIT завершена. Запускаю скрипт через 3 секунды')
        print('\033[1m' + '\033[32m' + 'Сессия FINIT завершена. Запускаю скрипт через 3 секунды' + '\033[0m')
        time.sleep(3)
        os.system('cls')
    except:
        logging.error("", exc_info=True)


# Получение всех имен из папки в passwork
def get_pass(data):
    logging.info(f'get_pass({data})')
    try:
        if data == 'oo':
            action = f'/folders/{OO_FOLDER}/passwords'
        elif data == 'tochka':
            action = f'/folders/{TOCHKA_FOLDER}/passwords'
        else:
            logging.error(f'Неверный параметр функции: {data}')
            os._exit(0)
        head = {'Passwork-Auth': TOKEN}
        r = get(URL + action, headers=head, verify=False).json()
        de = []
        for i in r['data']:
            de.append(i)
        res = sorted(de, key=lambda d: d['name'])
        logging.debug(f"Выгрузка всех имен ноутбуков из сейфа {data}: {res}")
        return res
    except:
        logging.error("", exc_info=True)


# Поиск ноутбуков по базе passwork
def search(data):
    logging.info(f'search({data})')
    print(f"Ищу S/N: {data} в базе passwork... \n")
    try:
        action = '/passwords/search'
        head = {'Passwork-Auth': TOKEN}
        data_e = {'query': data, "vaultId": VAULT_ID}
        r = post(URL + action, data=data_e, headers=head, verify=False).json()
        if r['data'] == []:
            logging.info('Совпадений не найдено')
            return 'Совпадений не найдено'
        elif not r['data'] == []:
            search_result = []
            for i in r['data']:
                # Если ноубук был найден в "Wrong names"
                if i['folderId'] == WN_FOLDER:
                    logging.info(f"Найден в неправильных паролях: {i['name']}, S/N: {i['login']}")
                    logging.info(f"Удаляю {i['name']}")
                    action = f"/passwords/{i['id']}"
                    head = {"Passwork-Auth": TOKEN, "Accept": "application/json", "Content-Type": "application/json"}
                    r = delete(URL + action, headers=head, verify=False).json()
                    if r['status'] == 'success':
                        logging.info(f"Ноутбук: {i['name']}, S/N: {i['login']} успешно удален")
                    else:
                        logging.error(f"Ошибка удаления {i['name']}")
                elif i['folderId'] == OO_FOLDER or i['folderId'] == TOCHKA_FOLDER:
                    try:
                        res = {'name': i['name'], 'S/N': i['login'], 'tag': NOTE_INFO['MTM'], 'pass_id': i['id'],
                               'folder_id': i['folderId'], 'vault_id': i['vaultId'], 'dsc': i['description']}
                    except:
                        res = {'name': i['name'], 'S/N': i['login'], 'tag': NOTE_INFO['MTM'], 'pass_id': i['id'],
                               'folder_id': i['folderId'], 'vault_id': i['vaultId'], 'dsc': ''}
                    search_result.append(res)
                else:
                    logging.error('Не найдено такое расположение')
            if search_result == []:
                logging.info('Совпадений не найдено')
                return 'Совпадений не найдено'
            else:
                logging.info(search_result)
                return search_result
    except:
        logging.error("", exc_info=True)


# Удаление зарезервированного ноутбука
def unreserv(id):
    logging.info(f'unreserv({id})')
    try:
        action = f"/passwords/{id}"
        head = {"Passwork-Auth": TOKEN, "Accept": "application/json", "Content-Type": "application/json"}
        r = delete(URL + action, headers=head, verify=False).json()
        if r['status'] == 'success':
            logging.info(f"{id} Успешно удален")
        else:
            logging.error("Ошибка удаления зарезервированного ноутбука")
    except:
        logging.error("", exc_info=True)


# Создание пароля в passwork
def create_password(name, id, folder, sn, mtm):
    logging.info(f'edit_password({name, id, sn, mtm})')
    try:
        action = f'/passwords/{id}'
        head = {"Passwork-Auth": TOKEN, "Accept": "application/json", "Content-Type": "application/json"}
        data_e = {"name": name, "login": sn, "encryptedPassword": "123", "color": 7, "url": "",
                  "description": DSC, "masterHash": "123",
                      "folderId": folder, "vaultId": VAULT_ID, "tags": [mtm]}
        data_e = dumps(data_e)
        r = put(URL + action, data=data_e, headers=head, verify=False).json()
        if r['status'] == 'success':
            name = r['data']['name']
            logging.info(f' >>> Пароль: "{name} Login: {sn}  Tag: {mtm}" успешно создан!')
        else:
            logging.error(f"Ошибка создания пароля {name}, {sn}, {id}")
    except:
        logging.error("", exc_info=True)


# Изменение пароля в passwork
def edit_password(name, id, folder, sn, mtm, dsc):
    logging.info(f'edit_password({name, id, sn, mtm})')
    try:
        action = f'/passwords/{id}'
        dsc_new = dsc + '\n\n' + DSC_2
        head = {"Passwork-Auth": TOKEN, "Accept": "application/json", "Content-Type": "application/json"}
        data_e = {"name": name, "login": sn, "encryptedPassword": "123", "color": 7, "url": "",
                  "description": dsc_new, "masterHash": "123",
                      "folderId": folder, "vaultId": VAULT_ID, "tags": [mtm]}
        data_e = dumps(data_e)
        r = put(URL + action, data=data_e, headers=head, verify=False).json()
        if r['status'] == 'success':
            name = r['data']['name']
            logging.info(f' >>> Пароль: "{name} Login: {sn}  Tag: {mtm}" успешно изменен!')
        else:
            logging.error(f"Ошибка изменения пароля {name}, {sn}, {id}")
    except:
        logging.error("", exc_info=True)


# Получения пароля администратора из passwork
def adm_pass():
    logging.info('adm_pass()')
    try:
        action = f'/passwords/{ADM_ID}'
        head = {'Passwork-Auth': TOKEN}
        r = get(URL + action, headers=head, verify=False).json()
        if r['status'] == 'success':
            password = b64decode(r['data']['cryptedPassword']).decode('utf-8')
            logging.info('Пароль администратора успешно получен, но в логи я его не напишу :P')
            subprocess.run(f"net user Administrator {password}", shell=True)
            logging.info('Пароль администратора успешно изменен')
        else:
            logging.error('Ошибка получения пароля администратора')
            print('\033[1m' + '\033[91m' + 'Ошибка получения пароля администратора. Завершение работы через 5 секунд' + '\033[0m')
            time.sleep(5)
            os._exit(0)
    except:
        logging.error("", exc_info=True)


# Получение последних ноутбуков в папках Tochka и OO
def get_last():
    logging.info('get_last():')
    try:
        action = f'/passwords'
        tochka_list = get_pass('tochka')
        tochka_last = tochka_list[-1]['name']
        tochka_last = 'Tochka' + str(int(tochka_last[6:]) + 1)
        head = {"Passwork-Auth": TOKEN, "Accept": "application/json", "Content-Type": "application/json"}
        tochka_data = {"name": tochka_last, "login": '<reserv>', "cryptedPassword": "123", "url": "",
                       "masterHash": "123", "vaultId": VAULT_ID, "folderId": TOCHKA_FOLDER}
        tochka_data = dumps(tochka_data)
        r = post(URL + action, data=tochka_data, headers=head, verify=False).json()
        if r['status'] == 'success':
            tochka_id = r['data']['id']
            logging.info(f'{tochka_last} создан и зарезервирован')
        else:
            logging.error(f"Ошибка создания и резервирования {tochka_last}")
        return tochka_last, tochka_id
    except:
        logging.error("", exc_info=True)


# Основная функция
def work_process():
    logging.info('work_process()')
    is_admin()
    get_info()
    auth()
    search_res = search(NOTE_INFO["SN"])

    # Если не нашлось совпадений в passwork:
    if search_res == 'Совпадений не найдено':
        logging.info('Совпадений не найдено, создаю новый пароль')
        name_list = get_last()
        name = name_list[0]
        input(f"Не нашел совпадений в passwork, зарезервировал имя {name}, нажми Enter для продолжения ")
        logging.info(f"Enter нажат, начинаю создание {name}")
        create_password(name, name_list[1], TOCHKA_FOLDER, NOTE_INFO["SN"], NOTE_INFO["MTM"])
        rename_computer(name)
        create_user_win(name, '1qaz2wsx', False)

    # Если было найдено больше одной записи в passwork:
    elif len(search_res) > 1:
        logging.error("Поиск нашел несколько записей с этим S/N")
        print('\033[1m' + '\033[91m' + "\nПоиск нашел несколько записей с этим S/N, зайди в passwork и разберись с этим вручную!" + '\033[0m')
        input("\nЖми Enter для завершения работы")
        os._exit(0)

    # Если нашлось совпадение в папке Tochka:
    elif search_res[0]['folder_id'] == TOCHKA_FOLDER:
        dsc = search_res[0]['dsc']
        name = search_res[0]['name']
        id = search_res[0]['pass_id']
        logging.info(f"Нашел совпадение - {name}")
        print(f"Нашел совпадение в passwork: {name}. Ноут будет перезалит с этим именем.\n")
        input('\033[1m' + '\033[93m' + f"А ты удалил ноут {name} из Касперского??? Жми Enter что бы продолжить" + '\033[0m')
        logging.info(f"Enter нажат, начинаю создание {name}")
        edit_password(name, id, TOCHKA_FOLDER, NOTE_INFO["SN"], NOTE_INFO["MTM"], dsc)
        rename_computer(name)
        create_user_win(name, '1qaz2wsx', False)

    # Если совпадение где-то ещё:
    elif search_res[0]['folder_id'] == OO_FOLDER or search_res[0]['folder_id'] == WN_FOLDER:
        name_list = get_last()
        name = name_list[0]
        reserv_id = name_list[1]
        dsc = search_res[0]['dsc']
        old_name = search_res[0]['name']
        id = search_res[0]['pass_id']
        logging.info(f"Нашел совпадение - {old_name} в OO или Wrong names")
        print(f"Нашел совпадение - {old_name} в OO или Wrong names. Ноутбук будет перезалит с именем {name}\n")
        input('\033[1m' + '\033[93m' + f"А ты удалил ноут {name} из Касперского??? Жми Enter что бы продолжить" + '\033[0m')
        logging.info(f"Enter нажат, начинаю изменение {old_name} на {name}")
        edit_password(name, id, TOCHKA_FOLDER, NOTE_INFO["SN"], NOTE_INFO["MTM"], dsc)
        unreserv(reserv_id)
        rename_computer(name)
        create_user_win(name, '1qaz2wsx', False)

    create_user_win('TeamLead', 'k@KnnqfB', False)
    adm_pass()
    logout()
    logging.info("Пытаюсь запустить скрипт del.bat...")
    print("\nЗапускаю скрипт...\n")
    time.sleep(2)
    try:
        subprocess.Popen("c://soft//del.bat")
    except:
        logging.error("", exc_info=True)


# Старт программы
packages.urllib3.disable_warnings() # Отключаем предупреждения
os.system('color') # Включаем цветной режим консоли
work_process()