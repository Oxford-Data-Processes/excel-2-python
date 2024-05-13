import xlcalculator
import ast
from ast_transformation.formula_generator_old import SeriesIdLoader


class SeriesDependenciesBuilder:

    @staticmethod
    def build_dependencies(formula_1_ast_series_list):

        series_dependencies = {}

        # Iterate through each series_id and its corresponding formula AST
        for series_id, formula_1_ast_series in formula_1_ast_series_list.items():
            # Get the list of dependent series_ids from the AST
            dependencies_series_ids = SeriesDependenciesBuilder.get_series_ids(
                formula_1_ast_series
            )

            # Check if the series_id is already in the DAG adjacency list
            if series_id not in series_dependencies:
                series_dependencies[series_id] = []

            # Add the dependencies to the series_id's list in the DAG
            series_dependencies[series_id].extend(dependencies_series_ids)

        # Optional: Remove duplicates if necessary (depends on the application's needs)
        for key, values in series_dependencies.items():
            # Convert the list to a set and back to a list to remove duplicates
            series_dependencies[key] = list(set(values))

        return series_dependencies

    @staticmethod
    def get_series_ids(formula_1_ast_series):

        series_range_strings = SeriesDependenciesBuilder.extract_range_values(
            formula_1_ast_series
        )

        series_ids = []
        for series_range_string in series_range_strings:

            series_id_strings = ast.literal_eval(str(series_range_string))[0]
            for series_id_string in series_id_strings:
                series_id = SeriesIdLoader.load_series_id_from_string(series_id_string)
                series_ids.append(series_id)
        return series_ids

    @staticmethod
    def extract_range_values(ast):
        """Static method to extract range values from an AST."""
        range_values = []
        SeriesDependenciesBuilder.process_node(ast, range_values)
        return range_values

    @staticmethod
    def process_node(node, range_values):
        """Recursively process each node based on its type, appending range node values."""
        if isinstance(node, xlcalculator.ast_nodes.RangeNode):
            SeriesDependenciesBuilder.handle_range_node(node, range_values)
        elif isinstance(node, xlcalculator.ast_nodes.FunctionNode):
            SeriesDependenciesBuilder.handle_function_node(node, range_values)
        elif isinstance(node, xlcalculator.ast_nodes.OperatorNode):
            SeriesDependenciesBuilder.handle_operator_node(node, range_values)

    @staticmethod
    def handle_range_node(node, range_values):
        """Extract the value from a RangeNode and add it to the list."""
        range_values.append(node.tvalue)

    @staticmethod
    def handle_function_node(node, range_values):
        """Process a FunctionNode by iterating over its arguments."""
        for arg in node.args:
            SeriesDependenciesBuilder.process_node(arg, range_values)

    @staticmethod
    def handle_operator_node(node, range_values):
        """Process an OperatorNode by processing its left and right children."""
        if node.left:
            SeriesDependenciesBuilder.process_node(node.left, range_values)
        if node.right:
            SeriesDependenciesBuilder.process_node(node.right, range_values)
