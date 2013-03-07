# -----------------------------------------------------------------------------
# kewl_test.py
#
# A simple calculator with variables.   This is from O'Reilly's
# "Lex and Yacc", p. 63.
# -----------------------------------------------------------------------------

import sys, random
import ply.lex as lex
import ply.yacc as yacc

if sys.version_info[0] >= 3:
    raw_input = input

tokens = (
    'CELL',
    'RELCELL',
    'NAME',
    'EQ',
    'DOLLAR',
    'NUMBER',
    'NUMBER_HEX',
    'INTEGER',
    'FN',
    'ARGSEP',

    'MULT',
    'DIV',
    'SUB',
    'ADD',
    
    'TYPEDEF',
    'STRING',
    'LPAREN',
    'RPAREN',
    'LBRACKET',
    'RBRACKET',
    'LBRACE',
    'RBRACE'
)

# a string needs to match "this is something\nand another line\nand a quote \"welcome\""
t_EQ        = r'='
t_ARGSEP    = r','
t_NAME      = r'\$?[A-Za-z][A-Za-z0-9_]*'
# t_NAME      = r'BACK'

t_MULT      = r'\*'
t_DIV       = r'/'
t_SUB       = r'-'
t_ADD       = r'\+'

t_DOLLAR    = r'\$(list|type)?'
t_LPAREN    = r'\('
t_RPAREN    = r'\)'
t_LBRACKET  = r'\['
t_RBRACKET  = r'\]'
t_LBRACE    = r'\{'
t_RBRACE    = r'\}'

def t_RELCELL(t):
    r'R\[-?[0-9]+\]C'
    return t

def t_FN(t):
    r'IF|AVERAGE|VLOOKUP'
    return t

def t_NUMBER_HEX(t):
    r'0[0-9]+h'
    return t

def t_CELL(t):
    r'[A-Z]+[0-9]+'
    return t

def t_NUMBER(t):
    r'[-]?[0-9]+(\.[0-9]*)?'
    try:
        t.value = int(t.value)
    except ValueError:
        print("Integer value too large $d", t.value)
    return t

def t_INTEGER(t):
    r'[-]?[0-9]+(\.[0-9]*)?'
    try:
        t.value = int(t.value)
    except ValueError:
        print("Integer value too large $d", t.value)
    return t

def t_STRING(t):
    r'"(.|\\[nt\"])*"'
    t.value = t.value[1:-1]
    return t

def t_COMMENT(t):
    r'\#.*'
    t.lexer.lineno += 1
    return None

def t_TYPEDEF(t):
    r'@t_[a-z][a-z0-9]+'
    return t

def t_newline(t):
    r'\n+'
    t.lexer.lineno += t.value.count("\n")

t_ignore    = " \t"
    
def t_error(t):
    print("Illegal character '%s'" % t.value[0])
    t.lexer.skip(1)
    

# dictionary of names
names = { }

def p_formula(p):
    'formula : EQ functions'
    p[0] = ('FORMULA', [p[1]] + p[2])

def p_functions(p):
    'functions : function functions'
    p[0] = [p[1]] + p[2]

def p_functions_empty(p):
    'functions : '
    p[0] = []

def p_function(p):
    '''function : FN LPAREN args RPAREN'''
    p[0] = ('FUNC', p[1], p[3] )     # invoke FUNCtion called p[1] with arguments p[3]

# def p_list_structure(p):
#     'list : LBRACKET terms RBRACKET'
#     p[0] = ('LIST', p[2])

def p_args(p):
    'args : term terms'
    p[0] = ('ARGLIST', [p[1]] + p[2] )

def p_term_function(p):
    'term : function'
    p[0] = p[1]

def p_term_string(p):
    'term : STRING'
    p[0] = ('STRING', p[1])

def p_term_relcell(p):
    'term : RELCELL'
    p[0] = ('RELCELL', p[1])

def p_term_name(p):
    'term : NAME'
    p[0] = ('NAME', p[1])

def p_term_fn(p):
    'term : FN'
    p[0] = ('TERM', p[1])

def p_term_numberhex(p):
    'term : NUMBER_HEX'
    p[0] = ('NUMBER', p[1])

def p_term_number(p):
    'term : NUMBER'
    p[0] = ('NUMBER', p[1])

# def p_term_list(p):
#     'term : list'
#     p[0] = p[1]

# def p_term_map(p):
#     'term : map'
#     p[0] = ('MAP', p[1])

def p_terms(p):
    'terms : ARGSEP term terms'
    p[0] = [p[2]] + p[3]

def p_args_empty(p):
    'terms :'
    p[0] = []

# def p_map(p):
#     'map : LBRACE terms RBRACE'
#     p[0] = ('MAP', p[2])



lexer = lex.lex()
yacc.yacc()


if __name__ == '__main__':
    t = "=VLOOKUP(VLOOKUP(R[-95]C,elements,2,FALSE),sectionList,2,FALSE)"
    lexer.input(t)

    print(t)
    while True:
        tok = lexer.token()
        if not tok: break      # No more input
        print(tok)

    print("DONE")

    tree = yacc.parse(t)
    print(tree)