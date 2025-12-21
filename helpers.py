import re

COMPARATORS = ['>=', '<=', '<>', '>', '<', '=']

class Helpers:
    @classmethod
    def is_number(self, s: str) -> bool:
        try:
            float(s)
            return True
        except ValueError:
            return False
        
    @classmethod
    def end_parenthese(self, word: str, start: int) -> int:
        parenthese_found = 0
        list_chars = list(word)

        for index in range(start, len(word)):
            if list_chars[index] == '(':
                parenthese_found += 1
            elif list_chars[index] == ')':
                if parenthese_found == 0:
                    return index
                else:
                    parenthese_found -= 1

        return None
    
    @classmethod
    def end_single_quote(self, word: str, start: int) -> int:
        list_chars = list(word)

        for index in range(start, len(word)):
            if list_chars[index] == "'":
                    return index

        return None

    @classmethod
    def separate_condition(self, condition: str):
        for comparator in COMPARATORS:
            if comparator in condition:
                comp_index = condition.index(comparator)

                part1 = condition[:comp_index]
                part2 = condition[comp_index+len(comparator):]

                if comparator == '=':
                    comparator = '=='

                return part1, comparator, part2

        raise ValueError(f'Unknown comparator in {condition}')
    
    @classmethod
    def reference_sheetname(self, module: str, reference: str) -> str:
        end_sheet_reference = reference.index('!') - 1
        start_sheet_reference = end_sheet_reference - Helpers.end_single_quote(reference[:end_sheet_reference][::-1], 0)

        sheetname_referenced = re.sub(r'\W+', '', reference[start_sheet_reference:end_sheet_reference]).title().replace(' ', '')

        cell = ''
        for i in range(end_sheet_reference + 2, len(reference)):
            char = reference[i]
            if not re.match(r"[\$A-Za-z0-9]", char):
                break

            if not char == '$':
                cell += char
        
        return f'{module}::{sheetname_referenced}.cell(:{cell.lower()})'
    
    @classmethod
    def is_cell_reference(self, reference) -> bool:
        pattern = r'^\$?[A-Z]{1,3}\$?\d+$'
        return bool(re.match(pattern, reference))
    
    @classmethod
    def split_excel_args(self, args_str: str):
        args = []
        paren_level = 0
        current = ''

        for c in args_str:
            if c == ',' and paren_level == 0:
                args.append(current.strip())
                current = ''
            else:
                current += c
                if c == '(':
                    paren_level += 1
                elif c == ')':
                    paren_level -= 1

        if current:
            args.append(current.strip())

        return args

    @classmethod
    def cell_to_int(self, cell: str):
        row = 0
        col = 0

        for c in cell:
            if c.isdigit():
                row = row * 10 + int(c)
            elif c.isalpha():
                col = col * 26 + (ord(c.upper()) - ord('A') + 1)

        return [col, row]
