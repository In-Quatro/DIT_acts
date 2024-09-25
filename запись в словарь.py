import csv
from collections import defaultdict
from pprint import pprint

template_date = fr'Данные_для_актов\excel_data_1.csv'


def create_in_memory_csv(template):
    """Разбивка данных на структуру в памяти по подписи."""
    data = defaultdict(list)  # Словарь для хранения данных по подписи

    with open(template, newline='') as csvfile:
        reader = csv.DictReader(csvfile, delimiter=';')

        for row in reader:
            pprint(row)
    #         sign = row['Общее МО']
    #         data[sign].append(row)
    #
    # for mo, tt in data.items():
    #     pprint(mo)
    #     for t in tt:
    #         pprint(t)


create_in_memory_csv(template_date)
