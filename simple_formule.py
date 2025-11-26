from helpers import Helpers

class SimpleFormule:
    def __init__(self, formule: str, is_simple: bool):
        self.formule = formule[1:]
        self.is_simple = is_simple

    def exec(self):
        return self.resolve_expression(self.formule)

    def excel_if(self, if_str: str):
        args = Helpers.split_excel_args(if_str)
        if len(args) < 2:
            raise ValueError(f'A excel if statement need at least 2 section a condition, a value if contion is true')

        condition = self.resolve_expression(args[0])
        if_content = self.resolve_expression(args[1])
        else_content = self.resolve_expression(args[2]) if len(args) == 3 else 'nil'

        return (
            f'if {condition}\n'
            f'    {if_content}\n'
            f'else\n'
            f'    {else_content}\n'
            f'end\n'
        )

    def excel_and_or(self, exp_str: str, is_or = False):
        res = []

        for condition in Helpers.split_excel_args(exp_str):
            part1, comp, part2 = Helpers.separate_condition(condition)

            part1 = self.resolve_expression(part1)
            part2 = self.resolve_expression(part2)

            res.append(f'{part1} {comp} {part2}')

        if is_or:
            return f"({' || '.join(res)})"
        else:
            return f"({' && '.join(res)})"

    def excel_not(self, not_str: str):
        try:
            part1, comp, part2 = Helpers.separate_condition(not_str)

            part1 = self.resolve_expression(part1)
            part2 = self.resolve_expression(part2)

            return (f'!({part1} {comp} {part2})')
        except Exception:
            return f'!{self.resolve_expression(not_str)}'

    def resolve_expression(self, expression: str) -> str:
        expression = expression.strip()
        expression = expression.replace('_xlfn._xlws', '')

        if expression.startswith('NOT('):
            return self.excel_not(expression[4:-1])
        elif expression.startswith('AND('):
            return self.excel_and_or(expression[4:-1], False)
        elif expression.startswith('OR('):
            return self.excel_and_or(expression[3:-1], True)
        elif '!' in expression:
            return Helpers.reference_sheetname(expression)
        elif Helpers.is_cell_reference(expression):
            return expression.lower()
        elif '&' in expression:
            return ' + '.join([self.resolve_expression(p.strip()) + '.to_s' for p in expression.split('&')])
        elif expression.startswith('FALSE('):
            return 'false'
        else:
            return expression
