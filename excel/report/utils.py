def int_r(num: float) -> int:
    """
    Округление дробных чисел.
    """
    num = int(num + (0.5 if num > 0 else -0.5))
    return num


def funk_for_total_calk(row):
    """
    Проверяет налоговую базу и высчитывает исчисление всего по формуле.
    Если налоговая база не заданна пропускает строку.
    Результат вычислений округляется.
    """
    value = row.isnull()
    if value[2]:
        return row[3]
    one = row[2] / 100
    if row[2] <= 5000000:
        val = one * 13
    else:
        val = one * 15
    return int_r(val)


def funk_for_deviation(row):
    """
    Высчитывает отклонения для строки.
    """
    return row['Исчислено всего'] - row['Исчислено всего по формуле']


def highlight(value, color):
    """
    Окрашивает ячейку по заданным условиям. Если значение отсутствует окрашивание ячейки не задается.
    """
    if value == 0:
        return f'background-color: {color}'
    elif value > 0 or value < 0:
        return 'background-color: red'
    else:
        return
