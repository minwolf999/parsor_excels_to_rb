import re

from openpyxl.utils import get_column_letter

from src.helpers import Helpers

class ExcelSum:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression

  def exec(self):
    if self.expression.startswith('SUM('):
       return self.normal_sum()

  def normal_sum(self):
    terms = []
    expression = self.expression[3:-1]

    for current_range in expression.split(','):
        cells = current_range.strip().split(':')
        
        end_cell_name = re.sub(r'\W+', '', cells[-1].lower())

        start_reference_sheetname = Helpers.reference_sheetname(self.module, cells[0])
        end_reference_sheetname = start_reference_sheetname.split('(:')[0] + f"(:{end_cell_name})"

        start_cell = Helpers.cell_to_int(start_reference_sheetname.split('(:')[1][:-1])
        end_cell = Helpers.cell_to_int(end_reference_sheetname.split('(:')[1][:-1])

        end_sheet_reference = cells[0].index('!') - 1
        start_sheet_reference = end_sheet_reference - Helpers.end_single_quote(cells[0][:end_sheet_reference][::-1], 0)
        sheetname_referenced = re.sub(r'\W+', '', cells[0][start_sheet_reference:end_sheet_reference]).title().replace(' ', '')

        for col in range(start_cell[0], end_cell[0] + 1, 1):
            for row in range(start_cell[1], end_cell[1] + 1, 1):
                terms.append(f'{self.module}::{sheetname_referenced}.cell(:{get_column_letter(col + 1)}{row})')

    return f'{terms}.sum'
