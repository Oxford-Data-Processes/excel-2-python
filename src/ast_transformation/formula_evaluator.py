import formulas


class FormulaEvaluator:
    def __init__(self):
        self.formula_parser = formulas.Parser()

    def evaluate_formula(self, formula_string):

        # Ensure the formula starts with an '=' sign for consistent parsing
        if not formula_string.startswith("="):
            formula_string = "=" + formula_string

        # Compile and evaluate the formula
        function = self.formula_parser.ast(formula_string)[1].compile()
        result = function()

        if isinstance(result, formulas.functions.Array):
            if result.shape == ():
                return result.item()
            return result[0][0]
        elif isinstance(result, float):
            return result

        return result
