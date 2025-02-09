import requests
import json
import os
import time

# user, token, название компании скрыты
def execute_method(upload_url_method, ID, user='my_id', token='my_webhook'):
    disk_url = f'https://my_company.bitrix24.ru/rest/{user}/{token}/'
    upload_url_response = requests.get(
    disk_url + upload_url_method,{'id': ID}).json()
    return upload_url_response

def upload():
    path_to_file = os.getcwd()+'\\'+'реестр семян.xlsx'
    methods = ['disk.file.delete', 'disk.folder.getchildren', 'disk.folder.uploadFile']
    id_folder = 'id_my_folder' # id папки скрыт

    "ПРОВЕРИТЬ СОДЕРЖИМОЕ ПАПКИ И ПОЛУЧИТЬ ID ФАЙЛА. ЕСЛИ ОН ЕСТЬ, ТО УДАЛИТЬ ФАЙЛ"
    url_response = execute_method(methods[1], id_folder)
    if url_response['result']:
        id_file = url_response['result'][0]['ID']
        "УДАЛИТЬ ФАЙЛ"
        url_response = execute_method(methods[0], id_file)

    time.sleep(1)

    "ДОБАВИТЬ ФАЙЛ"
    url_response = execute_method(methods[2], id_folder)
    upload_url = (url_response['result']['uploadUrl'])
    with open(path_to_file, 'rb') as f:
            content = f.read()
    requests.post(upload_url, files={'file':[path_to_file, content]})

if __name__ == '__main__':
    upload()
    
