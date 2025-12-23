from src.helpers import Helpers
from src.simple_formule import SimpleFormule

class ExcelNot:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression[4:-1]

  def exec(self):
    try:
      part1, comp, part2 = Helpers.separate_condition(self.expression)

      part1 = SimpleFormule(self.module, part1).exec()
      part2 = SimpleFormule(self.module, part2).exec()
      
      return (f'!({part1} {comp} {part2})')
    except ValueError:
      return f'!{SimpleFormule(self.module, self.expression).exec()}'
