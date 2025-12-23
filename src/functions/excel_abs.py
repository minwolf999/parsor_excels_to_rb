from src.simple_formule import SimpleFormule

class ExcelAbs:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression[3:-1]

  def exec(self):
    return f'({SimpleFormule(self.module, self.expression).exec}).abs'
