import graphviz
from xlcalculator import parser, ast_nodes


class FormulaVisualiser:

    def __init__(self):
        self.graph = graphviz.Digraph(comment="AST Visualization", format="png")
        self.node_count = 0
        self.graph.format = "png"

    def add_node(self, parent_name: str, node):
        node_name = str(self.node_count)
        self.node_count += 1

        # Determine the label based on the node type.
        if isinstance(node, ast_nodes.OperandNode):
            label = str(node.token.tvalue)  # Use the operand value
        elif isinstance(node, ast_nodes.FunctionNode):
            label = node.token.tvalue  # Function name
        elif isinstance(node, ast_nodes.OperatorNode):
            label = node.token.tvalue  # Operator symbol
        else:
            label = type(node).__name__  # Fallback to the class name

        # Create a new graph node with the determined label.
        self.graph.node(node_name, label)

        # Connect this node to its parent in the graph.
        if parent_name is not None:
            self.graph.edge(parent_name, node_name)

        # Recurse for child nodes, if any.
        if isinstance(node, ast_nodes.OperatorNode) and hasattr(node, "left"):
            self.add_node(node_name, node.left)
        if isinstance(
            node, (ast_nodes.OperatorNode, ast_nodes.FunctionNode)
        ) and hasattr(node, "right"):
            self.add_node(node_name, node.right)
        if isinstance(node, ast_nodes.FunctionNode) and hasattr(node, "args"):
            for arg in node.args:
                self.add_node(node_name, arg)

    def visualise(self, root):
        self.add_node(None, root)
        return self.graph
