from lark import Lark, Transformer, Tree, Token, v_args
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

foreach_stmt: "foreach" "(" expression ")" "as" "$" CNAME "{" statement+ "}"

output_stmt: expression ("|" expression)*

// -------------------
// EXPRESSIONS
// -------------------

expression: term ("+" term)*

term: index_access
    | var
    | quoted_string
    | builtin_function
    | INT

index_access: ("[" expression "]")+

// -------------------
// TERMINALS
// -------------------
quoted_string: ESCAPED_STRING
var: "$" CNAME
builtin_function: TIME_PERIOD | INFLUENCE | VARIABLES | NEW_LINE | FORCE_CUTOFF | DESCRIPTION "(" term ")" | STYLE "(" term ")" | SPAN "(" term "," INT ")"

TIME_PERIOD : "!time_period"
INFLUENCE : "!influence"
VARIABLES : "!variables"
NEW_LINE : "!newline"
DESCRIPTION : "!description"
FORCE_CUTOFF : "!force_cutoff"
SPAN : "!span"
STYLE : "!style"

%import common.CNAME
%import common.ESCAPED_STRING
%import common.INT
%import common.WS
%ignore WS
"""


class Parser():
    def __init__(self):
        self.parser = Lark(GRAMMAR, parser="lalr")
        self.tree = None
        self._cached = {}

    def parse(self, code : str):
        if code in self._cached:
            self.tree = self._cached[code]
        
        tree = self.parser.parse(code)
        self._cached[code] = tree
        self.tree = tree

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
        self.style = ""
        self.table_state = TableState()

    def start(self, items):
        for stmt in items:
            self._resolve(stmt)

    def index_access(self, items):
        is_static = True
        index = []

        # Resolve items statically
        for item in items:
            idx, static = self._static_resolve(item)
            is_static = is_static and static # If any item cannot be resolved, is_static will be False
            index.append(idx)

        # If all items are statically resolved, return immideatly
        if is_static:
            return self.info.get_value(*index)

        # Else defer for later
        def exec():
            resolved = [self._resolve(elm) for elm in index]
            return self.info.get_value(*resolved)
        return exec
    
    @v_args(inline=True)
    def var(self, token):
        # If the variable is already defined, return the value
        if token.value in self.vars:
            return self.vars[token.value]
        
        # Else defer for later, variable may be defined in an outer scope
        def exec():
            return self.vars[token.value]
        return exec
    
    @v_args(inline=True)
    def INT(self, token):
        return int(token.value)
    
    @v_args(inline=True)
    def quoted_string(self, token):
        return token.value[1:-1]  # remove quotes
    
    @v_args(inline=True)
    def builtin_function(self, token, *args):
        match token.value:
            case "!time_period":
                return self.info.time_periods
            case "!influence":
                return self.info.influences
            case "!variables":
                return self.info.variables
            case "!force_cutoff":
                return lambda : self.table_state.force_cutoff()
            case "!description":
                arg, is_static = self._static_resolve(args[0])
                if is_static:
                    return self.variable_descriptions[arg] #type: ignore
                return lambda: self.variable_descriptions[self._resolve(arg)] #type: ignore
            case "!style":
                def exec():
                    arg = self._resolve(args[0])
                    self.style = arg
                return exec
            case "!newline":
                def exec():
                    self.table_state.next_row()
                    self.table_state.reset_col()
                return exec
            case "!span":
                def exec():
                    text = self._resolve(args[0])
                    length = args[1]
                    self.table_state.set_style(self.style)
                    self.table_state.add_span(text, length)
                return exec

    def expression(self, items):
        # Resolve left and right statically as far as possible
        left, l_is_static = self._static_resolve(items[0])
        if len(items) == 1:
            return left
        
        right, r_is_static = self._static_resolve(items[1])

        if l_is_static and r_is_static:
            # If they can be statically concatenated, do so right away
            return str(left) + str(right)

        # If left or right cannot be statically resolved, defer the concatenation and return a function
        def exec():
            l = self._resolve(left)

            if len(items) == 1:
                return left

            r = self._resolve(right)

            return str(l) + str(r)
        return exec

    @v_args(inline=True)
    def foreach_stmt(self, iterable_tree, var_def_token, *body):
        s_iterable, _ = self._static_resolve(iterable_tree)
        var, _ = self._static_resolve(var_def_token) 
        def exec():
            iterable = self._resolve(s_iterable) 

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
                    self.table_state.set_style(self.style)
                    self.table_state.set_text(elm)
                    self.table_state.next_col()     
        return exec
    
    def _static_resolve(self, stmt):
        if callable(stmt):
            return (stmt, False)
        if isinstance(stmt, Tree):
            return self._static_resolve(stmt.children[0])
        if isinstance(stmt, Token):
            return (stmt.value, True)
        if isinstance(stmt, list):
            l = [self._static_resolve(x) for x in stmt]
            return ([x for x, _ in l], all([b for _, b in l]))

        return (stmt, True)

    def _resolve(self, stmt):
        if callable(stmt):
            return self._resolve(stmt())
        if isinstance(stmt, Tree):
            return self._resolve(stmt.children[0])
        if isinstance(stmt, Token):
            return stmt.value
        if isinstance(stmt, list):
            return [self._resolve(x) for x in stmt]        
        return stmt
