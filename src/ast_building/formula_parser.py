import xlcalculator


class FormulaParser:

    @staticmethod
    def parse_formula(formula: str) -> xlcalculator.ast_nodes.ASTNode:

        parser = xlcalculator.parser.FormulaParser()

        print(f"formula: {formula}")

        ast = parser.parse(formula=formula, named_ranges={})
        return ast
