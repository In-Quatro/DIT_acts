import os
import re
import csv
import logging
import subprocess
from pathlib import Path

from docx import Document


logging.basicConfig(level=logging.INFO,
                    format='%(levelname)s - %(asctime)s  - %(message)s',
                    datefmt='%H:%M:%S')


def write_to_csv(data, output_dir, queue_num):
    """Создание CSV файла и внесение данных."""
    csv_file = f'Данные по {queue_num}-ой очереди.csv'
    csv_folder = os.path.dirname(output_dir)
    cav_file_path = os.path.join(csv_folder, csv_file)
    file_exists = os.path.isfile(cav_file_path)
    with open(cav_file_path, mode='a', newline='') as file:
        writer = csv.writer(file, delimiter=';')

        if not file_exists:
            writer.writerow(
                [
                    'Код МО',
                    'Наименование МО',
                    'Должность МО',
                    'ФИО',
                    'Основание подписания'
                 ]
            )
        writer.writerow(data)


def check_queue_num():
    """Проверка корректности очереди."""
    while True:
        queue_num = input('Укажите номер очереди (1, 4, 5 или 13) '
                          'или 0 для выхода: ')
        if queue_num in ('1', '4', '5', '13'):
            return queue_num
        if queue_num == '0':
            exit('Завершение работы')
        else:
            print('Необходимо указать правильно очередь!')


def main():
    """Главная функция."""
    print('Выбран режим сбора данных из титульных листов')
    queue_num = check_queue_num()
    folder_name = fr'Данные из титульных листов\{queue_num}-я'
    files = os.listdir(folder_name)
    count_files = len(files)
    if not files:
        exit(f'Нет файлов в папке "{folder_name}" для обработки!')

    logging.info(f'Обнаружено файлов - {count_files}. Начинаю чтение')

    patterns = {
        'kod': r'\Технический акт №\xa0(.*?)\n',
        'naimenovanie': r'\, и (.*?)\s\(',
        'dolzhnost': r'\«МО», в лице (.*?)\s[А-ЯЁ]',
        'fio': r'[А-ЯЁ][а-яё]+\s[А-ЯЁ][а-яё]+\s[А-ЯЁ][а-яё]+',
        'osnovanie': r'\основании (.*?)\,',
    }
    for file in files:
        if Path(file).suffix == '.docx':
            file_path = Path(folder_name, file)
            doc = Document(file_path)
            text = '\n\n'.join([par.text for par in doc.paragraphs])
            data = [re.findall(match, text)[-1] for match in patterns.values()]
            write_to_csv(data, folder_name, queue_num)

    logging.info('Все данные считаны')
    project_folder = os.getcwd()
    subprocess.Popen(fr'explorer {project_folder}\Данные из титульных листов')


if __name__ == "__main__":
    main()
