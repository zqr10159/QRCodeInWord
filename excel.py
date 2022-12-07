import json
import re

from openpyxl.reader.excel import load_workbook
from openpyxl.workbook import Workbook

def write2excel(file_path, content: list, title=None):
    wb = Workbook()
    ws = wb.active
    if title is not None:
        ws.append(title)

    for row in content:
        ws.append(row)
    try:
        wb.save(file_path)
    except PermissionError:
        input("请确认文件是否关闭，任意键重试")
        wb.save(file_path)


def read_excel(file_path, sheet_num=0, begin_row=0):
    result_value = []
    wb = load_workbook(file_path, read_only=True, data_only=True)
    sheets = wb.sheetnames
    ws = wb[sheets[sheet_num]]
    for row in ws.rows:
        result_value.append([cell.value for cell in row])
    wb.close()
    return result_value[begin_row:]


def str2json(str_cont):
    return [json.loads("{{{}}}".format(i)) for i in re.findall(r"{(.*?)}", str_cont)]
