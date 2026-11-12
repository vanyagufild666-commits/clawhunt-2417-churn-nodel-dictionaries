"""Generate xlsx deliverable for task 2417 (two dictionaries, xlsx + Markdown)."""
import json

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

data = json.load(open('dictionaries.json', encoding='utf-8'))

HEADERS = ['编码', '名称', '定义', '是否需要填写补充说明', '后续建议动作', '排序值']
WIDTHS = [18, 14, 52, 20, 52, 8]

wb = Workbook()

thin = Side(style='thin', color='D0D0D0')
border = Border(left=thin, right=thin, top=thin, bottom=thin)
head_fill = PatternFill('solid', fgColor='1F4E79')
head_font = Font(color='FFFFFF', bold=True, name='Microsoft YaHei', size=10)
cell_font = Font(name='Microsoft YaHei', size=10)
wrap = Alignment(vertical='top', wrap_text=True)
center = Alignment(vertical='top', horizontal='center', wrap_text=True)

SHEET_TITLES = {'churn': '流失原因字典', 'nodel': '招商未成交原因字典'}

for idx, (key, title) in enumerate(SHEET_TITLES.items(), start=1):
    ws = wb.active if idx == 1 else wb.create_sheet()
    ws.title = title

    for col, (h, w) in enumerate(zip(HEADERS, WIDTHS), start=1):
        c = ws.cell(row=1, column=col, value=h)
        c.font = head_font
        c.fill = head_fill
        c.border = border
        c.alignment = center
        ws.column_dimensions[get_column_letter(col)].width = w

    for row, item in enumerate(data[key], start=2):
        values = [item['code'], item['name'], item['definition'],
                  '是' if item['supplement'] else '否', item['action'], (row - 1) * 10]
        for col, v in enumerate(values, start=1):
            c = ws.cell(row=row, column=col, value=v)
            c.font = cell_font
            c.border = border
            c.alignment = center if col in (4, 6) else wrap
        ws.row_dimensions[row].height = 42

    ws.freeze_panes = 'A2'
    ws.auto_filter.ref = f'A1:F{len(data[key]) + 1}'

wb.save('会员流失与未成交原因字典.xlsx')
print('xlsx written:', len(data['churn']), 'churn +', len(data['nodel']), 'nodel rows')
