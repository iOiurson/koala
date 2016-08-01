# cython: profile=True

import os.path

# import networkx

from excellib import FUNCTION_MAP, IND_FUN
from utils import is_range, split_range, split_address, flatten
from ExcelError import *

class ASTNode(object):
    """A generic node in the AST"""
    
    def __init__(self,token, debug = False):
        super(ASTNode,self).__init__()
        self.token = token
        self.debug = debug
    def __str__(self):
        return self.token.tvalue
    def __getattr__(self,name):
        return getattr(self.token,name)

    def children(self,ast):
        args = ast.predecessors(self)
        args = sorted(args,key=lambda x: ast.node[x]['pos'])
        return args

    def parent(self,ast):
        args = ast.successors(self)
        return args[0] if args else None

    def find_special_function(self, ast):
        found = False
        current = self

        special_functions = ['sumproduct']
        # special_functions = ['sumproduct', 'match']
        break_functions = ['index']

        while current is not None:
            if current.tvalue.lower() in special_functions:
                found = True
                break
            elif current.tvalue.lower() in break_functions:
                break
            else:
                current = current.parent(ast)

        return found

    def has_operator_or_func_parent(self, ast):
        found = False
        current = self

        while current is not None:
            if (current.ttype[:8] == 'operator' or current.ttype == 'function') and current.tvalue.lower() != 'if':
                found = True
                break
            else:
                current = current.parent(ast)

        return found

    def has_ind_func_parent(self, ast):     
      
        if self.parent(ast) is not None and self.parent(ast).tvalue in IND_FUN:       
            return True       
        else:     
            return False      


    def emit(self,ast,context=None):
        """Emit code"""
        self.token.tvalue
    
class OperatorNode(ASTNode):
    def __init__(self, args, ref = None, debug = False):
        super(OperatorNode,self).__init__(args)
        self.ref = ref if ref is not None else 'None' # ref is the address of the reference cell  
        self.debug = debug
        # convert the operator to python equivalents
        self.opmap = {
                 "^":"**",
                 "=":"==",
                 "&":"+",
                 "":"+" #union
                 }

        self.op_range_translator = {
            "*": "multiply",
            "/": "divide",
            "+": "add",
            "-": "substract",
            "==": "is_equal",
            "<>": "is_not_equal",
            ">": "is_strictly_superior",
            "<": "is_strictly_inferior",
            ">=": "is_superior_or_equal",
            "<=": "is_inferior_or_equal"
        }

    def emit(self,ast,context=None):
        xop = self.tvalue
        
        # Get the arguments
        args = self.children(ast)
        codes, names = zip(*[a.emit(ast,context=context) for a in args])
        names = list(flatten(names))
        
        op = self.opmap.get(xop,xop)
        
        parent = self.parent(ast)
        # convert ":" operator to a range function
        if op == ":":
            # print 'TEST: %s' % ':'.join([a.emit(ast,context=context).replace('"', '') for a in args])
            # OFFSET HANDLER, when the first argument of OFFSET is a range i.e "A1:A2"
            range = ':'.join(codes).replace('"', '')

            if (parent is not None and
            (parent.tvalue == 'OFFSET' and 
             parent.children(ast)[0] == self)):
                
                code = '"%s"' % range
            else:
                code = 'self.eval_ref("%s", ref = %s)' % (range, self.ref)

            names.insert(0, range) # updating the named ranges emit() outputs
            return code, names
         
        if self.ttype == "operator-prefix":
            code = "RangeCore.apply_one('minus', %s, None, %s)" % (codes[0], str(self.ref))
            return code, names

        if op in ["+", "-", "*", "/", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            call = 'apply' + ('_all' if is_special else '')
            function = self.op_range_translator.get(op)

            code = "RangeCore." + call + "(%s)" % ','.join(["'"+function+"'", str(codes[0]), str(codes[1]), str(self.ref)])

            return code, names

        parent = self.parent(ast)

        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        if op == "<" or op == "<=":
            aa = codes[0]
            code = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + codes[1]
        elif op == ">" or op == ">=":
            aa = codes[1]
            code =  codes[1] + op + "(" + aa + " if " + aa + " is not None else float('inf'))"
        else:
            code = codes[0] + op + codes[1]
                    

        #avoid needless parentheses
        if parent and not isinstance(parent,FunctionNode):
            code = "("+ code + ")"          

        return code, names

class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)
    def emit(self,ast,context=None):
        t = self.tsubtype
        
        if t == "logical":
            return str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            #if the string contains quotes, escape them
            val = self.tvalue.replace('"','\\"')
            return '"' + val + '"', []
        else:
            return str(self.tvalue), []

class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range, e.g., A5, B3:C20 or INPUT """
    def __init__(self,args, ref = None, debug = False):
        super(RangeNode,self).__init__(args)
        self.ref = ref if ref is not None else 'None' # ref is the address of the reference cell  
        self.debug = debug

    def get_cells(self):
        return resolve_range(self.tvalue)[0]
    
    def emit(self,ast,context=None):
        if isinstance(self.tvalue, ExcelError):
            if self.debug:
                print 'WARNING: Excel Error Code found', self.tvalue
            return self.tvalue, []

        is_a_range = False
        is_a_named_range = self.tsubtype == "named_range"

        if is_a_named_range:
            my_str = "'" + str(self) + "'" 
        else:
            rng = self.tvalue.replace('$','')
            sheet = context + "!" if context else ""

            is_a_range = is_range(rng)

            if is_a_range:
                sh,start,end = split_range(rng)
            else:
                try:
                    sh,col,row = split_address(rng)
                except:
                    if self.debug:
                        print 'WARNING: Unknown address: %s is not a cell/range reference, nor a named range' % str(rng)
                    sh = None

            if sh:
                my_str = '"' + rng + '"'
            else:
                my_str = '"' + sheet + rng + '"'

        to_eval = True
        # exception for formulas which use the address and not it content as ":" or "OFFSET"
        parent = self.parent(ast)
        # for OFFSET, it will also depends on the position in the formula (1st position required)
        if (parent is not None and
            (parent.tvalue == ':' or
            (parent.tvalue == 'OFFSET' and parent.children(ast)[0] == self) or
            (parent.tvalue == 'CHOOSE' and parent.children(ast)[0] != self and self.tsubtype == "named_range"))):
            to_eval = False

        # if parent is None and is_a_named_range: # When a named range is referenced in a cell without any prior operation
        #     return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))
                        
        if to_eval == False:
            return my_str, []

        # OFFSET HANDLER
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref)), []
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref)), []

        # INDEX HANDLER
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[0] == self):

            # return 'self.eval_ref(%s)' % my_str

            # we don't use eval_ref here to avoid empty cells (which are not included in Ranges)
            if is_a_named_range:
                return 'resolve_range(self.named_ranges[%s])' % my_str, []
            else:
                return 'resolve_range(%s)' % my_str, []
        
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref)), []
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref)), []
        # MATCH HANDLER
        elif parent is not None and parent.tvalue == 'MATCH' \
             and (parent.children(ast)[0] == self or len(parent.children(ast)) == 3 and parent.children(ast)[2] == self):
            return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref)), []
        elif self.find_special_function(ast) or self.has_ind_func_parent(ast):
            return 'self.eval_ref(%s)' % my_str, []
        else:
            return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref)), []
    
class FunctionNode(ASTNode):
    """AST node representing a function call"""
    def __init__(self,args, ref = None, debug = False):
        super(FunctionNode,self).__init__(args)
        self.ref = ref if ref is not None else 'None' # ref is the address of the reference cell
        self.debug = False
        # map  excel functions onto their python equivalents
        self.funmap = FUNCTION_MAP
        
    def emit(self,ast,context=None):
        fun = self.tvalue.lower()

        # Get the arguments
        args = self.children(ast)
        codes, names = zip(*[a.emit(ast,context=context) for a in args])
        names = list(flatten(names))

        if fun == "atan2":
            # swap arguments
            return "atan2(%s,%s)" % (codes[1],codes[0]), names
        elif fun == "pi":
            # constant, no parens
            return "pi", []
        elif fun == "if":
            # inline the if

            # check if the 'if' is concerning a Range
            is_range = False
            range = None
            childs = args[0].children(ast)

            for child in childs:
                if ':' in child.tvalue and child.tvalue != ':':
                    is_range = True
                    range = child.tvalue
                    break

            if is_range: # hack to filter Ranges when necessary,for instance situations like {=IF(A1:A3 > 0; A1:A3; 0)}
                return 'RangeCore.filter(self.eval_ref("%s"), %s)' % (range, codes[0]), names
            if len(args) == 2:
                return "%s if %s else 0" %(codes[1],codes[0]), names
            elif len(args) == 3:
                return "(%s if %s else %s)" % (codes[1],codes[0],codes[2]), names
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":
            my_str = '['
            if len(args) == 1:
                # only one row
                my_str += codes[0]
            else:
                # multiple rows
                my_str += ",".join(codes)
                     
            my_str += ']'

            return my_str, names
        elif fun == "arrayrow":
            #simply create a list
            return ",".join(codes), names

        elif fun == "and":
            return "all([" + ",".join(codes) + "])", names
        elif fun == "or":
            return "any([" + ",".join(codes) + "])", names
        elif fun == "index":
            if self.parent(ast) is not None and self.parent(ast).tvalue == ':':
                return 'index(' + ",".join(codes) + ")", names
            else:
                return 'self.eval_ref(index(%s), ref = %s)' % (",".join(codes), self.ref), names
        elif fun == "offset":
            if self.parent(ast) is None or self.parent(ast).tvalue == ':':
                return 'offset(' + ",".join(codes) + ")", names
            else:
                return 'self.eval_ref(offset(%s), ref = %s)' % (",".join(codes), self.ref), names
        else:
            # map to the correct name
            f = self.funmap.get(fun,fun)
            return f + "(" + ",".join(codes) + ")", names
