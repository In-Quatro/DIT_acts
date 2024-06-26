import os
import re
import csv
import logging

from docx import Document


logging.basicConfig(level=logging.INFO,
                    format='%(levelname)s - %(asctime)s  - %(message)s',
                    datefmt='%H:%M:%S')


def fill_docx_template(template_path, output_path, data):
    """Создание docx на основе шаблона."""
    doc = Document(template_path)

    for paragraph in doc.paragraphs:
        for key, value in data.items():
            if key in paragraph.text:
                inline = paragraph.runs
                for i in range(len(inline)):
                    if key in inline[i].text:
                        inline[i].text = re.sub(re.escape(key), value,
                                                inline[i].text)

    doc.save(output_path)


def main():
    """Главная функция."""
    print('Выбран режим создание Титульных листов для технических актов')
    num = input('Укажите номер очереди (1, 4, 5 или 13): ')

    if num not in ('1', '4', '5', '13'):
        exit('Необходимо указать правильно очередь!')

    folder_template = 'Шаблоны'
    template_path = f'{folder_template}/Шаблон_для_титульного_листа.docx'
    output_folder = f'Титульные_листы_{num}-я_очередь'
    date = f'word_data_{num}.csv'
    folder_date = 'Данные_для_актов'
    date_path = f'{folder_date}/{date}'

    if not os.path.exists(date_path):
        exit(f'Нет файла "{date}" для "{num}" очереди!')

    if not os.path.exists(output_folder):
        os.mkdir(output_folder)
        logging.info(f'Создаю папку "{output_folder}"')

    with open(date_path, encoding='ANSI') as csv_file:
        file_reader = csv.DictReader(csv_file, delimiter=";")
        for row in file_reader:
            kod_mo = row['kod']
            fio = row['client']
            post = row['position']
            logging.info(f'Создаю "{kod_mo}_{post}_{fio}"')
            output_path = f"{output_folder}/{kod_mo}_{post}_{fio}.docx"
            fill_docx_template(template_path, output_path, row)


if __name__ == "__main__":
    main()
