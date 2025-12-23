from helpers import Helpers

from src.functions.excel_ifs import ExcelIfs
from src.functions.excel_if import ExcelIf
from src.functions.excel_not import ExcelNot
from src.functions.excel_and_or import ExcelAndOr
from src.functions.excel_sum import ExcelSum
from src.functions.excel_false import ExcelFalse

class SimpleFormule:
    def __init__(self, module: str, formule: str):
        self.module = module
        self.formule = formule[1:]

    def exec(self):
        expression = self.formule.strip().replace('_xlfn._xlws', '')

        if expression.startswith('IFS'):
            return ExcelIfs(self.module, expression).exec()
        elif expression.startswith('IF'):
            return ExcelIf(self.module, expression).exec()
        elif expression.startswith('NOT'):
            return ExcelNot(self.module, expression).exec()
        elif expression.startswith('AND'):
            return ExcelAndOr(self.module, expression, False).exec()
        elif expression.startswith('OR'):
            return ExcelAndOr(self.module, expression, True).exec()
        elif expression.startswith('SUM'):
            return ExcelSum(self.module, expression).exec()
        elif expression.startswith('TRUE'):
            return ExcelTrue(self.module, expression).exec()
        elif expression.startswith('FALSE'):
            return ExcelFalse(self.module, expression).exec()
        elif '!' in expression:
            return Helpers.reference_sheetname(self.module, expression)
        elif '&' in expression:
            return ' + '.join([f'{SimpleFormule(self.module, p.strip()).exec()}.to_s' for p in expression.split('&')])
        elif Helpers.is_cell_reference(expression):
            return Helpers.reference_sheetname(self.module, expression)
        else:
            return expression
