from src.helpers import Helpers
from src.simple_formule import SimpleFormule

class ExcelIfs:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression[4:-1]

  def exec(self):
    parts = Helpers.split_excel_args(self.expression)
    expressions = [";".join(parts[i:i+2]) for i in range(0, len(parts), 2)]

    res = ''

    for index, expression in enumerate(expressions):
      args = Helpers.split_excel_args(expression)
      
      condition = SimpleFormule(self.module, args[0]).exec()
      content = SimpleFormule(self.module, args[1]).exec()

      operator = 'if' if index == 0 else 'elsif'

      res += f'{operator} {condition}\n'
      res += f'  {content}\n'

    res += 'else\n'
    res += '  nil\n'
    res += 'end\n'

    return res
