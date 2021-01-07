import os
import openpyxl
from openpyxl.worksheet.page import PageMargins


def format_excel(left, right, top, bottom):
    if not os.path.exists('./new'):
        os.mkdir('./new')
    xlsx_files = (fn for fn in os.listdir(".") if fn.endswith('.xlsx'))
    for file in xlsx_files:
        print(file)
        wb = openpyxl.load_workbook(file)
        # wb = load_workbook(file, keep_links=False)
        sheets_name = wb.get_sheet_names()
        for i in range(len(sheets_name)):
            sheet = wb[sheets_name[i]]
            sheet.page_setup.fitToHeight = False
            sheet.page_setup.fitToWidth = False
            sheet.page_margins = PageMargins(left=left, right=right, top=top, bottom=bottom)
        wb.save(r'./new/' + file)
        # save_workbook(wb, r'./new/' + file)


if __name__ == '__main__':
    left, right, top, bottom = eval(input("输入页边距格式，默认格式：2.50,1.50,1.50,1.50"
                                                         "表示左2.51，右1.50，上1.50，下1.50: ") or "2.50,1.50,1.50,1.50")
    format_excel(left="%.3f" % (left/2.54), right="%.3f" % (right/2.54), top="%.3f" % (top/2.54),
                bottom="%.3f" % (bottom/2.54))
