from openpyxl.utils import get_column_letter

from src.helpers import Helpers
from src.simple_formule import SimpleFormule

class ExcelSum:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression

  def exec(self):
    if self.expression.startswith('SUM('):
      return self.normal_sum(self.expression[3:-1])
    elif self.expression.startswith('SUM.IF('):
      return self.sum_if(self.expression[6:-1])
    elif self.expression.startswith('SUM.IF.ENS('):
      return self.sum_if_ens(self.expression[10:-1])

  def normal_sum(self, expression: str):
    terms = [Helpers.cell_range(current_range) for current_range in expression.split(',')]
    return f'{terms}.sum'

  def sum_if(self, expression: str):
    args = Helpers.split_excel_args(expression)
    plage = Helpers.cell_range(args[0])
    critere = SimpleFormule(self.module, args[1]).exec()

    if len(args) == 3:
      some_plage = Helpers.cell_range(args[2])
    else:
      some_plage = plage

    return f'{some_plage}.each_with_index.map {{ |v, i| ({plage}[i] {critere}) ? v : 0}}.sum'

  def sum_if_ens(self, expression: str):
    args = Helpers.split_excel_args(expression)
    some_plage = Helpers.cell_range(args[0])

    conditions = []

    for i in range(1, len(args), 2):
      plage_critere = Helpers.cell_range(args[i])
      critere = SimpleFormule(self.module, args[i + 1]).exec()
      conditions.append(f'({plage_critere}[i] {critere})')

    condition_globale = ' && '.join(conditions)
    return (
      f'{some_plage}.each_with_index.map {{ |v, i| \n'
      f'  ({condition_globale}) ? v : 0\n'
      f'}}.sum\n'
    )
