import os
import re
import qrcode
from docx import Document
from docx.oxml.ns import qn
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.shared import Inches, Cm, Pt

import excel
import get_pic

img_file = r'dst'
filename = r'out.docx'


def get_excel_row_num():
    src = excel.read_excel(excel_file, 0, 0)
    return len(src)


def get_excel_content():
    src = excel.read_excel(excel_file, 0, 0)
    return src


def create_table(document, tuopanhao, ):
    data_count = get_excel_row_num()
    if data_count % 2 != 0:
        data_count = data_count + 1
    row_num = 5 + int(data_count / 2)
    col_num = 4
    table = document.add_table(rows=row_num, cols=col_num, style='Table Grid')
    table = document.tables[-1]
    # 行宽
    for row in table.rows:
        row.cells[0].width = Cm(2)
    for row in table.rows:
        row.cells[1].width = Cm(6)
    for row in table.rows:
        row.cells[2].width = Cm(2)
    for row in table.rows:
        row.cells[3].width = Cm(6)

    # 第一行 物料编码
    table.cell(0, 0).text = '物料编码'
    # table.cell(0,1).text = ''
    # 第二行 数量
    table.cell(1, 0).text = '数量'
    # table.cell(1, 1).text = ''
    # 第三行 托盘号
    table.cell(2, 0).text = '托盘号'
    table.cell(2, 1).text = tuopanhao
    table.rows[2].height = Cm(2.5)
    # 第四行 名称规格
    table.cell(3, 0).text = '名称规格'
    table.cell(3, 1).merge(table.cell(3, 3))
    # table.cell(3, 1).text = ''
    # 第五行 No
    table.cell(4, 0).text = 'NO.'
    table.cell(4, 1).text = 'S.N.'
    table.cell(4, 2).text = 'NO.'
    table.cell(4, 3).text = 'S.N.'
    # 右上角二维码

    qr = qrcode.QRCode(
        version=1,
        error_correction=qrcode.constants.ERROR_CORRECT_H,
        box_size=3,
        border=4
    )
    tuopanfile = img_file + '\\' + tuopanhao + '.jpg'
    qr.add_data(tuopanhao)
    qr.make(fit=True)
    img = qr.make_image()
    img.save(img_file + '\\' + tuopanhao + '.jpg')
    picture = table.cell(0, 2).paragraphs[0].add_run().add_picture(
        tuopanfile)
    picture.height = Cm(2.66)
    picture.width = Cm(2.66)

    table.cell(2, 2).paragraphs[0].add_run().add_text(tuopanhao)
    table.cell(0, 2).merge(table.cell(2, 3))

    # 序号与二维码
    content = get_excel_content()
    for i in range(0, row_num - 5):
        table.cell(5 + i, 0).text = str(i * 2 + 1)

        picture = table.cell(5 + i, 1).paragraphs[0].add_run().add_picture(
            r'dst/' + content[i * 2][0] + '.jpg')
        picture.height = Cm(1.8)
        picture.width = Cm(1.8)

        table.cell(5 + i, 1).paragraphs[0].add_run().add_text('\n'+content[i * 2][0])

        table.cell(5 + i, 2).text = str(i * 2 + 2)

        picture = table.cell(5 + i, 3).paragraphs[0].add_run().add_picture(
            r'dst/' + content[i * 2 + 1][0] + '.jpg')
        picture.height = Cm(1.8)
        picture.width = Cm(1.8)
        table.cell(5 + i, 3).paragraphs[0].add_run().add_text('\n'+content[i * 2 + 1][0])
    # 所有内容居中
    for table in document.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    para.paragraph_format.alignment = WD_TABLE_ALIGNMENT.CENTER


if __name__ == "__main__":
    tuopanhao = input("请输入托盘号:\n")
    excel_file = input("请将excel拖入程序窗口:\n")
    excel_file = re.sub('"','',excel_file)
    print(excel_file)
    src = excel.read_excel(excel_file, 0, 0)
    if (os.path.exists('dst')):
        print('目录存在')
    else:
        os.mkdir('dst')
    get_pic.get(src)
    document = Document()
    document.styles['Normal'].font.name = 'Times New Roman'
    document.styles['Normal']._element.rPr.rFonts.set(qn('w:eastAsia'), u'宋体')
    document.styles['Normal'].font.size = Pt(9)

    create_table(document, tuopanhao)
    document.save('table.docx')
