# Парсер для сайта https://qgold.com/

Этот скрипт представляет собой парсер данных о ювелирных изделиях с веб-страниц и сохраняет их в файлы Excel. Он использует модуль requests для выполнения HTTP-запросов и модуль openpyxl для работы с файлами Excel.

## Установка

Для использования скрипта необходимо установить зависимости, указанные в файле requirements.txt. Вы можете установить их с помощью pip следующим образом:
pip install -r requirements.txt

## Использование

Скрипт выполняет парсинг данных различных URL-адресов ювелирных изделий, каждый из которых выполняется в отдельном потоке. Данные об изделиях сохраняются в файлы Excel. Скрипт может быть настроен для задания интервала между парсингом с помощью аргумента командной строки --interval.

Пример запуска скрипта с интервалом в 900 секунд (по умолчанию):
python product_parser.py

Пример запуска скрипта с интервалом в 600 секунд:
python product_parser.py --interval 600

## Блок-схема
![Блок-схема](images/diagram.png)

## Файлы

- `product_parser.py`: Главный файл скрипта, содержащий основной код для парсинга данных и сохранения их в файлы Excel.
- `requirements.txt`: Файл, содержащий список зависимостей для установки с помощью `pip`.

## Зависимости

Скрипт использует следующие зависимости:

- `requests`: Модуль для выполнения HTTP-запросов.
- `openpyxl`: Модуль для работы с файлами Excel.

Вы можете установить зависимости с помощью команды:
pip install -r requirements.txt
