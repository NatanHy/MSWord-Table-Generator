from lark import Lark, Transformer, Tree, v_args
from table_generation.geosphere import GeoSphereInfo
from .table_state import TableState
from typing import Dict

GRAMMAR = r"""
start: statement+

// -------------------
// STATEMENTS
// -------------------
statement: foreach_stmt
         | output_stmt

foreach_stmt: "foreach" "(" expression ")" "as" var_def "{" statement+ "}"

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
builtin_function: TIME_PERIOD | INFLUENCE | VARIABLES | NEW_LINE | FORCE_CUTOFF | DESCRIPTION "(" term ")" | SPAN "(" term "," INT ")"

TIME_PERIOD : "!time_period"
INFLUENCE : "!influence"
VARIABLES : "!variables"
NEW_LINE : "!newline"
DESCRIPTION : "!description"
FORCE_CUTOFF : "!force_cutoff"
SPAN : "!span"

%import common.CNAME
%import common.ESCAPED_STRING
%import common.INT
%import common.WS
%ignore WS
"""

class Parser():
    def __init__(self):
        self.parser = Lark(GRAMMAR, parser="lalr", start="start")
        self.tree = None

    def parse(self, code : str):
        self.tree = self.parser.parse(code)

    def execute(self, info : GeoSphereInfo, variable_descriptions : Dict[str, str]) -> TableState:
        executor = TableExecutor(info, variable_descriptions)

        if self.tree is not None:
            executor.transform(self.tree)
        else:
            raise ValueError("No parse tree found. Perhaps you forgot to parse before executing?")
        
        return executor.table_state

class TableExecutor(Transformer):
    def __init__(self, info : GeoSphereInfo, variable_descriptions : Dict[str, str]):
        self.info = info
        self.vars = {}
        self.variable_descriptions = variable_descriptions
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
    def builtin_function(self, token, *args):
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
                case "!force_cutoff":
                    self.table_state.force_cutoff()
                case "!description":
                    arg = self._resolve(args[0])
                    return self.variable_descriptions[arg]
                case "!span":
                    text = self._resolve(args[0])
                    length = int(args[1].value)
                    self.table_state.add_span(text, length)
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
    def foreach_stmt(self, iterable_tree, var_def_token, *body):
        def exec():
            # Index target will already be resolved
            iterable = self._resolve(iterable_tree.children[0])
            var = self._resolve(var_def_token).value

            for val in iterable:
                self.vars[var] = val
                for stmt in body:
                    self._resolve(stmt)
        return exec
    
    def output_stmt(self, items):
        def exec():
            for item in items:
                elm = self._resolve(item)
                if elm is not None:
                    if isinstance(elm, list):
                        self.table_state.set_elm(elm[0])
                    else:
                        self.table_state.set_elm(elm)
                    self.table_state.next_col()     
        return exec

    def _resolve(self, stmt):
        if callable(stmt):
            return self._resolve(stmt())
        if type(stmt) == Tree:
            return self._resolve(stmt.children[0])
        return stmt
