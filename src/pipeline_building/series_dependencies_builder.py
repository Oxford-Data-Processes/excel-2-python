import xlcalculator


class SeriesDependenciesBuilder:
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
