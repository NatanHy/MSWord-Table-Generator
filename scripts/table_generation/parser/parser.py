from lark import Lark, Transformer, Tree, v_args
from table_generation.geosphere import GeoSphereInfo
from .table_state import TableState

GRAMMAR = r"""
start: statement+

// -------------------
// STATEMENTS
// -------------------
statement: foreach_stmt
         | output_stmt

foreach_stmt: loop_type "(" expression ")" "as" var_def "{" statement+ "}"

loop_type: FOREACHROW | FOREACHCOL

FOREACHROW: "foreachrow"
FOREACHCOL: "foreachcol"

output_stmt: expression ("|" expression)*

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
var_def: "$" CNAME
builtin_function: TIME_PERIOD | INFLUENCE | VARIABLES | NEW_LINE

TIME_PERIOD : "!time_period"
INFLUENCE : "!influence"
VARIABLES : "!variables"
NEW_LINE : "!newline"

%import common.CNAME
%import common.ESCAPED_STRING
%import common.WS
%ignore WS
"""

class Parser():
    def __init__(self):
        self.parser = Lark(GRAMMAR, parser="lalr", start="start")
        self.tree = None

    def parse(self, code : str):
        self.tree = self.parser.parse(code)
        print(self.tree.pretty())

    def execute(self, info : GeoSphereInfo):
        executor = TableExecutor(info)

        if self.tree is not None:
            executor.transform(self.tree)
        else:
            raise ValueError("No parse tree found. Perhaps you forgot to parse before executing?")
        
        print(executor.table_state._arr)

class TableExecutor(Transformer):
    def __init__(self, info : GeoSphereInfo):
        self.info = info
        self.vars = {}
        self.table_state = TableState()

    def start(self, items):
        for stmt in items:
            self._resolve(stmt)

    def index_access(self, items):
        def exec():
            index = [self._resolve(item) for item in items]
            return list(self.info.get_values(*index))
        return exec
    
    @v_args(inline=True)
    def var(self, token):
        def exec():
            return self.vars[token.value]
        return exec
    
    @v_args(inline=True)
    def quoted_string(self, token):
        return token.value[1:-1]  # remove quotes
    
    @v_args(inline=True)
    def builtin_function(self, token):
        def exec():
            match token.value:
                case "!time_period":
                    return self.info.time_periods
                case "!influence":
                    return self.info.influences
                case "!variables":
                    return self.info.variables
                case "!newline":
                    self.table_state.next_row()
                    self.table_state.reset_col()
        return exec

    def concatenation(self, items):
        def exec():
            left = self._resolve(items[0])

            if len(items) == 1:
                return left

            right = self._resolve(items[1])

            return left + right
        return exec

    @v_args(inline=True)
    def foreach_stmt(self, loop_type_tree, iterable_tree, var_def_token, *body):
        def exec():
            # Since loop_type is a rule in the grammar, the left node is the loop type
            # loop_type = str(loop_type_tree.children[0])
            loop_type = self._resolve(loop_type_tree).value

            # Index target will already be resolved
            iterable = self._resolve(iterable_tree.children[0])
            var = self._resolve(var_def_token).value

            for val in iterable:
                self.vars[var] = val
                for stmt in body:
                    self._resolve(stmt)

                match loop_type:
                    case "foreachrow":
                        self.table_state.next_row()
                    case "foreachcol":
                        self.table_state.next_col()
        return exec
    
    def output_stmt(self, items):
        def exec():
            for item in items:
                elm = self._resolve(item)
                if elm is not None:
                    self.table_state.set_elm(elm)
                    self.table_state.next_col()
        return exec

    def _resolve(self, stmt):
        if callable(stmt):
            return self._resolve(stmt())
        if type(stmt) == Tree:
            return self._resolve(stmt.children[0])
        return stmt
