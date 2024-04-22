class DAGSorter:
    @staticmethod
    def sort_dag(graph):
        visited = set()
        sorted_nodes = []

        def visit(node):
            if node not in visited:
                visited.add(node)
                if node in graph:
                    for neighbor in graph[node]:
                        visit(neighbor)
                sorted_nodes.append(node)

        for node in graph:
            visit(node)

        return sorted_nodes
