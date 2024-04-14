class FormulaChecker:
    @staticmethod
    def check_formulas(ast_generator):
        formula_1_ast_new = ast_generator.get_nth_formula(n=1)
        formula_2_ast_new = ast_generator.get_nth_formula(n=2)

        formulas_are_correct = str(formula_1_ast_new) == str(
            ast_generator.formula_1_ast_series
        ) and str(formula_2_ast_new) == str(ast_generator.formula_2_ast_series)

        return formulas_are_correct, formula_1_ast_new, formula_2_ast_new
