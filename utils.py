import sys
import typing as tp
from pathlib import Path

Data = tp.List[tp.Tuple[int, str]]


def resource(relative_path: str):
    base_path = getattr(sys, "_MEIPASS", Path(__file__).resolve().parent)
    return Path(base_path) / relative_path


# 导出到excel表中
def write_excel(data: Data, path: tp.Union[str, Path]):
    import xlsxwriter

    workbook = xlsxwriter.Workbook(f"{path}/data.xlsx")
    worksheet = workbook.add_worksheet()

    titles = ["ID", "项目名称", "数据类型", "数据路径", "创建时间"]

    for i, title in enumerate(titles):
        worksheet.write(0, i, title)

    for i, each in enumerate(data, 1):
        id, name, data_class, path, createAt = each

        worksheet.write(i, 0, id)
        worksheet.write(i, 1, name)
        worksheet.write(i, 2, data_class)
        worksheet.write(i, 3, path)
        worksheet.write(i, 4, createAt)

    workbook.close()


def read_excel(path: tp.Union[str, Path]):
    from python_calamine import CalamineWorkbook

    workbook = CalamineWorkbook.from_path(path)
    data = workbook.get_sheet_by_name("Sheet1").to_python()

    data.pop(0)
    return data
