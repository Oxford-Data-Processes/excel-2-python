from typing import Callable
import xlcalculator


class NodeProcessor:
    @staticmethod
    def process_node(
        node: xlcalculator.ast_nodes.ASTNode, handler_func: "NodeHandler"
    ) -> xlcalculator.ast_nodes.ASTNode:
        """General method to process nodes by type and apply a handler function."""
        if isinstance(node, xlcalculator.ast_nodes.RangeNode):
            return handler_func.handle_range_node(node)
        elif isinstance(node, xlcalculator.ast_nodes.FunctionNode):
            return handler_func.handle_function_node(node)
        elif isinstance(node, xlcalculator.ast_nodes.OperatorNode):
            return handler_func.handle_operator_node(node)
        return node

    @staticmethod
    def recursive_process_node(
        node: xlcalculator.ast_nodes.ASTNode,
        process_func: Callable[
            [xlcalculator.ast_nodes.ASTNode], xlcalculator.ast_nodes.ASTNode
        ],
    ) -> xlcalculator.ast_nodes.ASTNode:
        """Recursively process a node by applying a processing function to its children."""
        if hasattr(node, "args"):  # This will typically match FunctionNode
            modified_args = [
                NodeProcessor.recursive_process_node(arg, process_func)
                for arg in node.args
            ]
            return node.__class__(token=node.token, args=modified_args)
        elif hasattr(node, "left") or hasattr(
            node, "right"
        ):  # Typically matches OperatorNode
            modified_left = (
                NodeProcessor.recursive_process_node(node.left, process_func)
                if node.left
                else None
            )
            modified_right = (
                NodeProcessor.recursive_process_node(node.right, process_func)
                if node.right
                else None
            )
            return node.__class__(
                token=node.token, left=modified_left, right=modified_right
            )
        return process_func(node)

    @staticmethod
    def apply_delta_to_node(
        node: xlcalculator.ast_nodes.ASTNode, delta: Any
    ) -> xlcalculator.ast_nodes.ASTNode:
        """Apply a calculated delta to a node, specific logic to be implemented based on node type."""
        # Example pseudocode:
        if isinstance(node, xlcalculator.ast_nodes.RangeNode):
            # Apply some delta calculation logic specific to RangeNodes
            return node  # Return modified node
        return node


class NodeHandler:
    def handle_range_node(
        self, node: xlcalculator.ast_nodes.RangeNode
    ) -> xlcalculator.ast_nodes.ASTNode:
        # Custom logic for RangeNode
        return node

    def handle_function_node(
        self, node: xlcalculator.ast_nodes.FunctionNode
    ) -> xlcalculator.ast_nodes.ASTNode:
        # Custom logic for FunctionNode
        return node

    def handle_operator_node(
        self, node: xlcalculator.ast_nodes.OperatorNode
    ) -> xlcalculator.ast_nodes.ASTNode:
        # Custom logic for OperatorNode
        return node
