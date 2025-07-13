with open("scripts/table.cfg", "r") as f:
    code = f.read()

tree = PARSER.parse(code)
print(tree.pretty())
