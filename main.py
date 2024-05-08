import acts, title_page, incidents, analysis_acts


menu = {
    '1': acts.main,
    '2': title_page.main,
    '3': incidents.main,
    '4': analysis_acts.main,
}


def main():
    """Навигатор по скриптам."""
    num = input(
        'Что нужно сделать:\n'
        '1 - Технические акты\n'
        '2 - Титульные листы для технических актов\n'
        '3 - Занести инциденты в технические акты\n'
        '4 - Собрать данные\n\n'
        '0 - Закончить работу\n')

    if num in menu:
        menu[num]()
    if num == '0':
        exit()
    else:
        print('Можно использовать только предложенные варианты!\n')
        main()


if __name__ == "__main__":
    main()
