from src.helpers import Helpers
from src.simple_formule import SimpleFormule

class ExcelAndOr:
  def __init__(self, module: str, expression: str, is_or: bool):
    self.module = module
    self.expression = expression[3:-1] if is_or else expression[4:-1]
    self.is_or = is_or

  def exec(self):
    res = []

    for condition in Helpers.split_excel_args(self.expression):
      part1, comp, part2 = Helpers.separate_condition(condition)

      part1 = SimpleFormule(self.module, part1).exec()
      part2 = SimpleFormule(self.module, part2).exec()

      res.append(f'{part1} {comp} {part2}')
  
    if self.is_or:
      return f"({' || '.join(res)})"
    else:
      return f"({' && '.join(res)})"
