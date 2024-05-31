import formulas
from formulas.builder import AstBuilder
from typing import Tuple, List


class FormulaParser:

    @staticmethod
    def parse_formula(
        formula: str,
    ) -> Tuple[list, AstBuilder]:
        ast = formulas.Parser().ast(formula)
        return ast
