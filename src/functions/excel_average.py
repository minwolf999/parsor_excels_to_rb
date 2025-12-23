from src.helpers import Helpers
from src.simple_formule import SimpleFormule

class ExcelAverage:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression

  def exec(self):
    if self.expression.startswith('AVERAGE('):
      return self.normal_average(self.expression[7:-1])
    elif self.expression.startswith('AVERAGEIF('):
      return self.average_if(self.expression[9:-1])
    elif self.expression.startswith('AVERAGEIFS('):
      pass

  def normal_average(self, expression: str):
    cells = [Helpers.cell_range(current_range) for current_range in expression.split(',')]

    return f'{cells}.flatten.sum.to_f / {cells}.flatten.length'

  def average_if(self, expression: str):
    args = Helpers.split_excel_args(expression)
    plage = Helpers.cell_range(args[0])
    critere = SimpleFormule(self.module, args[1]).exec()

    if len(args) == 3:
      some_plage = Helpers.cell_range(args[2])
    else:
      some_plage = plage

    return (
      f'values = {some_plage}.each_with_index.map {{ |v, i| ({plage}[i] {critere}) ? v : nil}}.compact\n'
      'values.sum.to_f / values.length\n'
    )

  def average_ifs(self, expression: str):
    args = Helpers.split_excel_args(expression)
    some_plage = Helpers.cell_range(args[0])

    conditions = []

    for i in range(1, len(args), 2):
      plage_critere = Helpers.cell_range(args[i])
      critere = SimpleFormule(self.module, args[i + 1]).exec()
      conditions.append(f'({plage_critere}[i] {critere})')

    condition_globale = ' && '.join(conditions)
    return (
      f'values = {some_plage}.each_with_index.map {{ |v, i| ({condition_globale}) ? v : nil}}.compact\n'
      'values.sum.to_f / values.length\n'
    )
