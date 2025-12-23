from src.helpers import Helpers
from src.simple_formule import SimpleFormule

class ExcelIf:
  def __init__(self, module: str, expression: str):
    self.module = module
    self.expression = expression[3:-1]

  def exec(self):
    args = Helpers.split_excel_args(self.expression)
    if len(args) < 2:
        raise ValueError(f'A excel if statement need at least 2 section a condition, a value if contion is true')

    condition = SimpleFormule(self.module, args[0]).exec()
    if_content = SimpleFormule(self.module, args[1]).exec()
    else_content = SimpleFormule(self.module, args[2]).exec() if len(args) == 3 else 'nil'

    return (
        f'if {condition}\n'
        f'    {if_content}\n'
        f'else\n'
        f'    {else_content}\n'
        f'end\n'
    )
