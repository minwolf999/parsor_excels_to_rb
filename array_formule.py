class ArrayFormule:
    def __init__(self, formula: str, is_simple: bool):
        self.formula = formula[1:]
        self.is_simple = is_simple

    def exec(self):
        return ''
