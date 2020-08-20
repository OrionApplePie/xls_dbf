import argparse

import dbf
import xlrd
from dbf import READ_WRITE


def get_from_rec(rec=None, cols=None):
    res = ""

    str_list = [
        str(rec[col_name]).strip() for col_name in cols
    ]
    # TODO: переработать логику, возможно отделить
    if len(cols) == 1:
        res = "".join(str_list)
    elif len(cols) == 3:
        surname, name, patr = str_list
        res = f"{surname} {name[:1]}{patr[:1]}"
    else:
        pass

    return res


def get_data_xls(filename="", sheet_num=0, rows_num=None, col_num_key=None, col_num_value=None):
    """Получение данных из файла excel по указанным строкам и столбцам.
    rows_num: пределы строк - ("с", "по"). Если не указано, то все строки.

    col_num_key: номер поля который станет ключем в словаре и
                  по которому будет вестись поиск.
    col_num_value: номер поля из которого берется значение суммы.
    Нумерация в эксель с нуля.
    Возвращает словарь: ключ - номер договора (или ФИО или как-то еще...),
                        значение - сумма.
    """
    book = xlrd.open_workbook(
        filename=filename,
        # formatting_info=True  # не работает для xlsx
    )

    sheet = book.sheet_by_index(sheet_num)

    if not rows_num:
        rows_num = (sheet.nrows,)
    start, end = rows_num

    rows_data = [
        sheet.row_values(rownum)
        for rownum in range(start, end+1)
    ]

    book.release_resources()
    del book

    return {
        row[col_num_key].strip(): row[col_num_value]
        for row in rows_data
    }


def find_dbf(filename="", data=None, dbf_comp_cols=None, dbf_add_col=""):
    """Поиск значения по заданному полю каждой записи из DBF в data словаре и,
    если значение найдено, запись данныхы в dbf в указанное поле.

    data: словарь {
            "ключ сравнения": "данные для записи"
          }
    """
    # открытие на чтение и запись dbf файла
    table = dbf.Table(
        filename=filename,
        codepage='cp866'
    )
    table.open(mode=READ_WRITE)

    for rec in dbf.Process(table):
        compare_col_value = get_from_rec(rec=rec, cols=dbf_comp_cols)
        print(f"compare value = {compare_col_value}")
        # TODO: предусмотреть случай полных тезок
        if compare_col_value in data:
            print(f"find {compare_col_value}")
            print(data[compare_col_value])
            rec[dbf_add_col] = float(data[compare_col_value])

    table.close()


def main():
    parser = argparse.ArgumentParser(
        description="""Сборщик телефонов.
        Ищет абонента из пришедшего файла xls в файле выгрузки DBF по
        заданному полю (ФИО, или номер счета), и если находит, проставляет соотв. сумму."""
    )

    parser.add_argument(
        '-op_file',
        type=str, action='store',
        dest='op_file', default='',
        help='Файл excel от оператора связи.'
    )

    parser.add_argument(
        '-esrn_file',
        type=str, action='store',
        dest='esrn_file', default='',
        help='Файл выгрузки связи из ЭСРН.'
    )

    parser.add_argument(
        '-xck', '--xls_col_key',
        type=int, action='store',
        dest='xls_col_key',
        help='Номер столбца из файла эксель - значения для поиска.'
    )

    parser.add_argument(
        '-xcv', '--xls_col_val',
        type=int, action='store',
        dest='xls_col_val',
        help='Номер столбца из файла эксель - значения для внесения в dbf.'
    )

    parser.add_argument(
        '-xr', '--xls_rows',
        nargs="+", type=int,
        dest='xls_rows',
        help='Номера первой и последней строк из файла эксель (от и по).'
    )

    parser.add_argument(
        '-dc', '--dbf_cols', nargs='+',
        help='Имена полей в DBF файле для поиска. Напр. одно поле номера счета L_TEL, или три поля ФИО',
        required=True, dest='dbf_cols',
    )

    parser.add_argument(
        '-dcw', '--dbf_col_write',
        help='Поле в DBF файле куда записывать значение из файла эксель, если есть совпадение.',
        required=True, dest='dbf_col_write',
    )

    args = parser.parse_args()

    xls_data = get_data_xls(
        filename=args.op_file,
        rows_num=args.xls_rows,
        col_num_key=args.xls_col_key,
        col_num_value=args.xls_col_val,
    )

    find_dbf(
        filename=args.esrn_file,
        data=xls_data,
        dbf_comp_cols=args.dbf_cols,
        dbf_add_col=args.dbf_col_write,
    )


if __name__ == '__main__':
    main()
