from lark import Lark, Transformer, v_args
from table_generation.geosphere import GeoSphereInfo

grammar = r"""
start: statement+

// -------------------
// STATEMENTS
// -------------------
statement: foreach_stmt
         | output_stmt

foreach_stmt: loop_type "(" expression ")" "as" var "{" statement+ "}"

loop_type: "foreachrow" | "foreachcol"

output_stmt: expression ("|" expression)* NEWLINE?

// -------------------
// EXPRESSIONS
// -------------------
expression: concatenation

concatenation: term ("+" term)*

term: index_access
    | var
    | quoted_string
    | builtin_function

index_access: ("[" expression "]")+

// -------------------
// TERMINALS
// -------------------
quoted_string: ESCAPED_STRING
var: "$" CNAME
builtin_function: "!time_period" | "!influence" | "!variables" | "!newline"

%import common.CNAME
%import common.ESCAPED_STRING
%import common.WS
%import common.NEWLINE
%ignore WS
"""

PARSER = Lark(grammar, parser="lalr", start="start")

class TableExecutor(Transformer):
    def __init__(self, info : GeoSphereInfo):
        self.info = info
        self.vars = {}
        self.line = 0

    def index_access(self, items):
        return list(self.info.get_values(*items))
    
    def var(self, token):
        return self.vars[str(token)]
    
    def quoted_string(self, token):
        return token[1:-1]  # remove quotes
    
    def builtin_function(self, token):
        match str(token):
            case "!time_period":
                return self.info.time_periods
            case "!influence":
                return self.info.influences
            case "!variables":
                return self.info.variables
            case "!newline":
                self.line += 1

    def concatenation(self, items):
        def exec():
            if len(items) == 1:
                return items[0]

            left = items[0]
            right = items[1]

            if isinstance(left, list) and isinstance(right, str):
                raise TypeError("Cannot concatenate list and string")
            return left + right
        return exec

    @v_args(inline=True) # Decorator to take multiple arguments to the function instead of just "items"
    def foreach_stmt(self, loop_type_tree, iterable_tree, var_token, *body):
        # Since loop_type is a rule in the grammar, the left node is the loop type
        loop_type = str(loop_type_tree.children[0])
        # Index target will already be resolved
        iterable = iterable_tree.children[0]
        var = str(var_token)

        results = []
        for val in iterable:
            self.vars[var] = val
            for stmt in body:
                if isinstance(stmt, list) and stmt[0] == 'OUTPUT':
                    results.append(stmt[1:])
    
    def output_statement(self, items):
        print(self.line, items)
