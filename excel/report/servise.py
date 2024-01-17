import pandas as pd
from django.conf import settings
from openpyxl.reader.excel import load_workbook
from openpyxl.styles import Alignment, Font, PatternFill
from pandas._typing import AggFuncType

from report.utils import funk_for_total_calk, funk_for_deviation, highlight


class ReportService:
    FILE_NAME = 'report'

    NEW_COLUMNS = {
        'Исчислено всего по формуле': funk_for_total_calk,
        'Отклонения': funk_for_deviation
    }

    MERGE_COLUMNS_INDEX = [
        (1, 2, 1, 1),
        (1, 2, 2, 2),
        (1, 2, 3, 3),
        (1, 1, 4, 5),
        (1, 2, 6, 6),
    ]

    COLUMN_WIDTH = [
        ('A', 30),
        ('B', 20),
        ('C', 20),
        ('D', 20),
        ('E', 20),
        ('F', 20),
    ]
    ROW_HEIGHT = [
        {1: 12},
        {2: 27}
    ]
    HEADERS_RENGE = [2, 6]

    CELL_ALIGNMENT_LIST = ['A1', 'B1', 'C1', 'D1', 'D2', 'E2', 'F1']

    def __init__(self, file: str):
        self.file = file
        self.parser_file = ExcelParsers(file)

    def create_report(self):
        self.parser_file.rename_columns()
        self.parser_file.del_column_by_value()
        self.create_columns()
        self.parser_file.sort_by_value()
        self.parser_file.create_style_to_df(highlight)
        filepath = self.parser_file.df_to_excel()
        filepath = self.formation_report(filepath)
        return filepath

    def create_columns(self) -> None:
        for column_name, func in self.NEW_COLUMNS.items():
            self.parser_file.create_columns_by_func(func, column_name)

    def formation_report(self, file_path: str) -> str:
        ws = Workbook(file_path)
        ws.merge_headers_cells(self.MERGE_COLUMNS_INDEX)
        ws.set_columns_name()
        ws.create_alignment_to_cells(self.CELL_ALIGNMENT_LIST)
        ws.set_cells_width(self.COLUMN_WIDTH)
        ws.change_row_height(self.ROW_HEIGHT)
        ws.format_all_cells(self.HEADERS_RENGE[1], size=10)
        ws.format_headers_cells(self.HEADERS_RENGE, size=10, bold=True)
        file_path = ws.save_document(file_path)

        return file_path


class ExcelParsers:

    def __init__(self, report_file, engine: str = 'openpyxl'):
        self.report_file = report_file
        self.engine = engine
        self.df = pd.read_excel(self.report_file, engine=self.engine, header=1, usecols=[0, 1, 4, 5])

    def rename_columns(self) -> None:
        new_columns_name_dict = {
            'Unnamed: 0': 'Филиал',
            'Unnamed: 1': 'Сотрудник',
            'Unnamed: 4': 'Налоговая база'
        }
        self.df.rename(columns=new_columns_name_dict, inplace=True)

    def del_column_by_value(self, value: str = 'Итого', column: str | int = 'Филиал', ) -> None:
        column_index = self.df[self.df[column] == value].index
        self.df = self.df.drop(column_index, axis=0)

    def create_columns_by_func(self, func: AggFuncType, column_name: str) -> None:
        self.df[column_name] = self.df.apply(func, axis=1)

    def sort_by_value(self, value: str = 'Отклонения', ascending: bool = False) -> None:
        self.df = self.df.sort_values(by=[value], ascending=ascending)

    def create_style_to_df(self, func: AggFuncType, column_name: str = 'Отклонения', color: str = 'green') -> None:
        self.df = self.df.style.applymap(func, color=color, subset=column_name)

    def df_to_excel(self, file_name: str = 'report', engine: str = 'openpyxl') -> str:
        excel_writer = settings.MEDIA_ROOT + file_name + '.xlsx'
        self.df.to_excel(excel_writer, engine=engine, index=False, sheet_name='Лист1')
        return excel_writer


class Workbook:
    def __init__(self, file_path: str):
        self.file_path = file_path
        self.wb = load_workbook(file_path)
        self.ws = self.wb.active
        self.headers = [item.value for item in self.ws[1]]
        self.ws.insert_rows(idx=0, amount=1)

    @property
    def headers_dict(self) -> dict:
        return {
            'A1': self.headers[0],
            'B1': self.headers[1],
            'C1': self.headers[2],
            'D1': 'Налог',
            'F1': self.headers[-1]
        }

    def set_columns_name(self) -> None:

        for cell, value in self.headers_dict.items():
            self.ws[cell] = value

    def create_alignment_to_cells(self, cells_list: list[str]) -> None:
        for cell in cells_list:
            self.ws[cell].alignment = Alignment(
                horizontal='center',
                vertical='center',
                wrapText=True
            )

    def set_cells_width(self, colum_list: list[tuple[str, int]]) -> None:
        for cell, _width in colum_list:
            self.ws.column_dimensions[cell].width = _width

    def change_row_height(self, dict_row: list[dict[int, int]]) -> None:
        for item in dict_row:
            index = list(item.items())[0]
            self.ws.row_dimensions[index[0]].height = index[1]

    def merge_headers_cells(self, list_indexes: list[tuple[int, int, int, int]]) -> None:
        for indexes in list_indexes:
            self.ws.merge_cells(
                start_row=indexes[0],
                end_row=indexes[1],
                start_column=indexes[2],
                end_column=indexes[3]
            )

    def format_all_cells(self, index_end_row: int, _fond: str = 'Arial', size: int = 11):
        font = self.create_font(_fond, size)
        for cell in self.ws.iter_rows():
            for i in range(0, index_end_row):
                cell[i].font = font

    def format_headers_cells(
            self,
            list_indexes: list[int],
            _font: str = 'Arial',
            size: int = 11,
            bold: bool = False,
            pattern: str = 'solid',
            fg_color: str = 'cbe4e5'
    ) -> None:
        font = self.create_font(_font, size, bold)
        fill = self.create_fill(pattern, fg_color)
        for cell in self.ws.iter_rows(min_row=1, max_row=list_indexes[0], min_col=1, max_col=list_indexes[1]):
            for i in range(0, list_indexes[1]):
                cell[i].font = font
                cell[i].fill = fill

    @staticmethod
    def create_font(name: str, size: int, bold: bool = False) -> Font:
        return Font(
            name=name,
            size=size,
            bold=bold
        )

    @staticmethod
    def create_fill(pattern: str, fg_color: str) -> PatternFill:
        return PatternFill(
            patternType=pattern,
            fgColor=fg_color
        )

    def save_document(self, filepath: str) -> str:

        self.wb.save(filepath)
        self.wb.close()
        return filepath
