from lark import Lark, Transformer, Tree, Token, v_args
from table_generation.component import ComponentInfo
from .table_state import TableState
from typing import Dict
import ast

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
builtin_function: TIME_PERIOD | INFLUENCE | VARIABLES | NEW_LINE | FORCE_CUTOFF | DESCRIPTION "(" term ")" | STYLE "(" term ")" | FORMAT "(" term ")" | SPAN "(" term "," INT ")"

TIME_PERIOD : "!time_period"
INFLUENCE : "!influence"
VARIABLES : "!variables"
NEW_LINE : "!newline"
DESCRIPTION : "!description"
FORCE_CUTOFF : "!force_cutoff"
SPAN : "!span"
STYLE : "!style"
FORMAT : "!format"

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

    def execute(self, info : ComponentInfo, variable_names : Dict[str, str]) -> TableState:
        executor = TableExecutor(info, variable_names)

        if self.tree is not None:
            executor.transform(self.tree)
        else:
            raise ValueError("No parse tree found. Perhaps you forgot to parse before executing?")
        
        return executor.table_state
    
class TableExecutor(Transformer):
    def __init__(self, info : ComponentInfo, variable_names : Dict[str, str]):
        self.info = info
        self.vars = {}
        self.variable_names = variable_names
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
        # use literal eval to evaluate escape sequences
        return ast.literal_eval(token.value)
    
    @v_args(inline=True)
    def builtin_function(self, token, *args):
        match token.value:
            case "!time_period":
                print(self.info.time_periods)
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
                    return self.variable_names[arg] #type: ignore
                return lambda: self.variable_names[self._resolve(arg)] #type: ignore
            case "!format":
                arg = self._resolve(args[0])
                self.table_state.format = arg #type: ignore
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
        left, _ = self._static_resolve(items[0])
        if len(items) == 1:
            return left
    
        def exec():
            resolved = [self._resolve(item) for item in items]
            stringified = [str(r) for r in resolved]
            return "".join(stringified)
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
