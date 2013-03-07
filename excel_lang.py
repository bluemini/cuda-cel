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
    'WBREF',
    'CELL',
    'CELLRANGE', 'ROWRANGE', 'COLRANGE',
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
    'GT','LT','GTE','LTE',
    
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
t_GT        = r'>'
t_LT        = r'<'
t_GTE       = r'>='
t_LTE       = r'<='

t_DOLLAR    = r'\$(list|type)?'
t_LPAREN    = r'\('
t_RPAREN    = r'\)'
t_LBRACKET  = r'\['
t_RBRACKET  = r'\]'
t_LBRACE    = r'\{'
t_RBRACE    = r'\}'

def t_WBREF(t):
    r'\[[^\]]+\]'
    raise Exception("External workbook references are not supported")
    return t

def t_RELCELL(t):
    r'R(\[-?[0-9]+\])?C(\[-?[0-9]+\])?'
    return t

def t_FN(t):
    r'AVG|IF|PI|SUM|VLOOKUP'
    return t

def t_NUMBER_HEX(t):
    r'0[0-9]+h'
    return t

def t_CELLRANGE(t):
    r'([A-Za-z0-9]+\!)?\$?[A-Z]+\$?[0-9]+:([A-Za-z0-9]+\!)?\$?[A-Z]+\$?[0-9]+'
    return t

def t_ROWRANGE(t):
    r'([A-Za-z0-9]+\!)?\$?[0-9]+:([A-Za-z0-9]+\!)?\$?[0-9]+'
    return t

def t_COLRANGE(t):
    r'([A-Za-z0-9]+\!)?\$?[A-Z]+:([A-Za-z0-9]+\!)?\$?[A-Z]+'
    return t

def t_CELL(t):
    r'([A-Za-z0-9]+\!)?\$?[A-Z]+\$?[0-9]+'
    return t

def t_NUMBER(t):
    r'[-]?[0-9]+(\.[0-9]*)?'
    try:
        t.value = float(t.value)
    except ValueError as e:
        print("String cannot be converted to float.", t.value, e)
    return t

def t_INTEGER(t):
    r'[-]?[0-9]'
    try:
        t.value = int(t.value)
    except ValueError as e:
        print("String cannot be converted to integer", t.value, e)
    return t

def t_STRING(t):
    r'"([^\"]|\\[nt\"])*"'
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

# formulas are the base cases for an = cell
def p_formula_fn(p):
    'formula : EQ functions'
    p[0] = ('FORMULA', [p[1]] + p[2])

def p_formula_exp(p):
    'formula : EQ expression'
    p[0] = ('FORMULA', [p[2]])

def p_formula_number(p):
    'formula : EQ NUMBER'
    p[0] = ('FORMULA', [p[2]])

def p_formula_cell(p):
    'formula : EQ CELL'
    p[0] = ('FORMULA', [p[2]])

def p_formula_array(p):
    'formula : EQ array'
    p[0] = ('FORMULA', [p[2]])


# array functions
def p_array_function(p):
    'array : LBRACE functions RBRACE'
    p[0] = ["ARRAY_FN", p[2]]


# functions..
def p_functions(p):
    'functions : function functions'
    p[0] = [p[1]] + p[2]

def p_functions_empty(p):
    'functions : '
    p[0] = []

def p_function(p):
    '''function : FN LPAREN terms RPAREN'''
    p[0] = ('FUNC', p[1], p[3] )     # invoke FUNCtion called p[1] with arguments p[3]

# terms..
def p_terms(p):
    'terms : ARGSEP terms'
    p[0] = p[2]

def p_moreterms(p):
    'terms : term terms'
    p[0] = [p[1]] + p[2]

def p_args_empty(p):
    'terms :'
    p[0] = []

# term defs
def p_term_function(p):
    'term : function'
    p[0] = p[1]

def p_term_string(p):
    'term : STRING'
    p[0] = ['STRING', p[1]]

def p_term_relcell(p):
    'term : RELCELL'
    p[0] = ['RELCELL', p[1]]

def p_term_name(p):
    'term : NAME'
    p[0] = ['NAME', p[1]]

def p_term_fn(p):
    'term : FN'
    p[0] = ['TERM', p[1]]

def p_term_numberhex(p):
    'term : NUMBER_HEX'
    p[0] = ['NUMBER', p[1]]

def p_term_number(p):
    'term : NUMBER'
    p[0] = ['NUMBER', p[1]]

def p_term_expression(p):
    'term : expression'
    p[0] = p[1]

def p_term_cell(p):
    'term : CELL'
    p[0] = p[1]

def p_term_cellrange(p):
    'term : CELLRANGE'
    p[0] = p[1]

def p_term_colrange(p):
    'term : COLRANGE'
    p[0] = p[1]

def p_term_rowrange(p):
    'term : ROWRANGE'
    p[0] = p[1]

# expression
def p_expression_sub(p):
    'expression : LPAREN terms RPAREN'
    p[0] = ['BINOP', p[2] ]

# def p_expression(p):
#     'expression : LPAREN expression RPAREN'
#     p[0] = ['BINOP', p[2], [p[1]] + p[3]]

def p_expression_binop(p):
    'expression : term binop terms'
    p[0] = ['BINOP', p[2], [p[1]] + p[3]]

def p_binop(p):
    '''binop : ADD
             | SUB
             | MULT
             | DIV
             | EQ
             | GT
             | LT
             | GTE
             | LTE'''
    p[0] = p[1]



lexer = lex.lex()
yacc.yacc()


if __name__ == '__main__':

    # Test inputs
    inputs = [
                # Simple test formulae
                '=1+3+5',
                '=3 * 4 + 5',
                '=50',
                '=1+1',
                '=$A1',
                '=$B$2',
                '=SUM(B5:B15)',
                '=SUM(B5:B15,D5:D15)',
                '=SUM(B5:B15 A7:D7)',
                '=SUM(sheet1!$A$1:$B$2)',
                        #'=[data.xls]sheet1!$A$1',
                '=SUM((A:A 1:1))',
                '=SUM((A:A,1:1))',
                '=SUM((A:A A1:B1))',
                '=SUM(D9:D11,E9:E11,F9:F11)',
                '=SUM((D9:D11,(E9:E11,F9:F11)))',

                # build the complex IF statement incrementally
                '=IF(P5=8.0,"G")',
                '=IF(P5=7.0,"F",IF(P5=8.0,"G"))',
                '=IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G")))',
                '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
                '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',

                '={SUM(B2:D2*B3:D3)}',
                '=SUM(123 + SUM(456) + (45<6))+456+789',
                '=AVG(((((123 + 4 + AVG(A1:A2))))))',

                # some from a real example
                '=VLOOKUP(VLOOKUP(R[-95]C,elements,2,FALSE),sectionList,2,FALSE)',
                '=IF(R[-277]C>R[-305]C-2*R[-303]C,\'-\',0.58*R[-319]C*PI()*R[-277]C*R[-303]C*R[-3]C/R[-404]C/1000)',
                '=IF(R[-277]C>R[-305]C,\'-\',FALSE)',
                '=IF(R[-277]C,\'-\',FALSE)',

                # E. W. Bachtal's test formulae
                '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
                '=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
                '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
                '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, ' + 
                    'R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] ' +
                    '*IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), ' +
                    'R[45]C[11],R[43]C[11])),0))',
              ]
    for t in inputs:
        print(t)
        dump = False

        try:
            tree = yacc.parse(t)
            print(tree)
            if tree == None:
                dump = True
        except Exception as e:
            print("ERROR:::", e)
            dump = True

        if dump:
            print("DUMP:::")
            lexer.input(t)
            while True:
                tok = lexer.token()
                if not tok: break      # No more input
                print(tok)
            break