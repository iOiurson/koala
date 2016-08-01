# cython: profile=True

import os.path
from math import pi
# import networkx

from koala.excellib import FUNCTION_MAP, FUNCTION_MAP_STRING, IND_FUN, index, offset
from koala.utils import is_range, split_range, split_address, resolve_range
from koala.ExcelError import *
from koala.Range import RangeCore

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
    def __init__(self, args, ref, debug = False):
        super(OperatorNode,self).__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell  
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

    def emit(self,ast,context=None, sp = None, mode = 'function'):

        xop = self.tvalue
        
        # Get the arguments
        args = self.children(ast)
        
        op = self.opmap.get(xop,xop)
        
        parent = self.parent(ast)
        # convert ":" operator to a range function
        if op == ":":
            # OFFSET HANDLER, when the first argument of OFFSET is a range i.e "A1:A2"
            if (parent is not None and
            (parent.tvalue == 'OFFSET' and 
             parent.children(ast)[0] == self)):
                if mode == 'function':
                    return lambda: ':'.join([a.emit(ast,context=context, sp = sp, mode = mode)() for a in args])
                else:
                    return '"%s"' % ':'.join([a.emit(ast,context=context, mode = mode).replace('"', '') for a in args])
            else:
                if mode == 'function':
                    return lambda: sp.eval_ref(*tuple([a.emit(ast,context=context, mode = mode, sp = sp)() for a in args]), ref = self.ref)
                else:
                    return "self.eval_ref(%s, ref = %s)" % (','.join([a.emit(ast,context=context, mode = mode) for a in args]), self.ref)

         
        if self.ttype == "operator-prefix":
            if mode == 'function':
                return lambda: RangeCore.apply_one('minus', args[0].emit(ast,context=context, sp = sp, mode = mode)(), None, self.ref)
            else:
                return "RangeCore.apply_one('minus', %s, None, %s)" % (args[0].emit(ast,context=context, mode = mode), str(self.ref))

        if op in ["+", "-", "*", "/", "==", "<>", ">", "<", ">=", "<="]:
            is_special = self.find_special_function(ast)
            function = self.op_range_translator.get(op)

            if mode == 'function':
                if is_special:
                    return lambda: RangeCore.apply_all(function, *tuple([a.emit(ast,context=context, sp = sp, mode = mode)() for a in args]))
                else:
                    return lambda: RangeCore.apply(function, *tuple([a.emit(ast,context=context, sp = sp, mode = mode)() for a in args]))
            else:
                arg1 = args[0]
                arg2 = args[1]
                call = 'apply' + ('_all' if is_special else '')
                return "RangeCore." + call + "(%s)" % ','.join(["'"+function+"'", str(arg1.emit(ast,context=context, mode = mode)), str(arg2.emit(ast,context=context, mode = mode)), str(self.ref)])

        # IS THIS STILL NEEDED ? SEEMS LIKE NOT
        #TODO silly hack to work around the fact that None < 0 is True (happens on blank cells)
        # if op == "<" or op == "<=":
        #     aa = args[0].emit(ast,context=context)
        #     ss = "(" + aa + " if " + aa + " is not None else float('inf'))" + op + args[1].emit(ast,context=context)
        # elif op == ">" or op == ">=":
        #     aa = args[1].emit(ast,context=context)
        #     ss =  args[0].emit(ast,context=context) + op + "(" + aa + " if " + aa + " is not None else float('inf'))"
        # else:

        raise Exception('Operator %s not handled' % op)
        ss = args[0].emit(ast,context=context, sp = sp, mode = mode) + op + args[1].emit(ast,context=context, sp = sp, mode = mode)
        
        # parent = self.parent(ast)         

        # #avoid needless parentheses
        # if parent and not isinstance(parent,FunctionNode):
        #     ss = "("+ ss + ")"          

        return ss

class OperandNode(ASTNode):
    def __init__(self,*args):
        super(OperandNode,self).__init__(*args)
    def emit(self,ast,context=None, sp = None, mode = 'function'):
        t = self.tsubtype
        
        if t == "logical":
            if mode == 'function':
                return lambda: self.tvalue.lower() == "true"
            else:
                return str(self.tvalue.lower() == "true")
        elif t == "text" or t == "error":
            #if the string contains quotes, escape them
            if mode == 'function':
                return lambda: self.tvalue.replace('"','\\"')
            else:
                val = self.tvalue.replace('"','\\"')
                return '"' + val + '"'
        else:
            if mode == 'function':
                return lambda: self.tvalue
            else:
                return str(self.tvalue)

class RangeNode(OperandNode):
    """Represents a spreadsheet cell, range, named_range, e.g., A5, B3:C20 or INPUT """
    def __init__(self,args, ref, debug = False):
        super(RangeNode,self).__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell  
        self.debug = debug

    def get_cells(self):
        return resolve_range(self.tvalue)[0]
    
    def emit(self,ast,context=None, sp = None, mode = 'function'):

        if isinstance(self.tvalue, ExcelError):
            if self.debug:
                print 'WARNING: Excel Error Code found', self.tvalue
            
            if mode == 'function':
                return lambda: self.tvalue
            else:
                return self.tvalue

        is_a_range = False
        is_a_named_range = self.tsubtype == "named_range"

        if is_a_named_range:
            name = str(self)
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
                name = rng
                my_str = '"' + rng + '"'
            else:
                name = sheet + rng
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

                        
        if to_eval == False:
            if mode == 'function':
                return lambda: name
            else:
                return my_str

        # OFFSET HANDLER
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            if mode == 'function':
                return lambda: sp.eval_ref(name, ref = self.ref)
            else:
                return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))
        elif (parent is not None and parent.tvalue == 'OFFSET' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            if mode == 'function':
                return lambda: sp.eval_ref(name, ref = self.ref)
            else:
                return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))

        # INDEX HANDLER
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[0] == self):

            # return 'self.eval_ref(%s)' % my_str

            # we don't use eval_ref here to avoid empty cells (which are not included in Ranges)
            if is_a_named_range:
                if mode == 'function':
                    return lambda: resolve_range(self.named_ranges[name])
                else:
                    return 'resolve_range(self.named_ranges[%s])' % my_str
            else:
                if mode == 'function':
                    return lambda: resolve_range(name)
                else:
                    return 'resolve_range(%s)' % my_str
        
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[1] == self and self.tsubtype == "named_range"):
            if mode == 'function':
                return lambda: sp.eval_ref(name, ref = self.ref)
            else:
                return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))
        elif (parent is not None and parent.tvalue == 'INDEX' and
             parent.children(ast)[2] == self and self.tsubtype == "named_range"):
            if mode == 'function':
                return lambda: sp.eval_ref(name, ref = self.ref)
            else:
                return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))
        # MATCH HANDLER
        elif parent is not None and parent.tvalue == 'MATCH' \
             and (parent.children(ast)[0] == self or len(parent.children(ast)) == 3 and parent.children(ast)[2] == self):
            if mode == 'function':
                return lambda: sp.eval_ref(name, ref = self.ref)
            else:
                return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))
        elif self.find_special_function(ast) or self.has_ind_func_parent(ast):
            if mode == 'function':
                return lambda: sp.eval_ref(name)
            else:
                return 'self.eval_ref(%s)' % my_str
        else:
            if mode == 'function':
                return lambda: sp.eval_ref(name, ref = self.ref)
            else:
                return 'self.eval_ref(%s, ref = %s)' % (my_str, str(self.ref))
    
class FunctionNode(ASTNode):

    """AST node representing a function call"""
    def __init__(self,args, ref, debug = False):
        super(FunctionNode,self).__init__(args)
        self.ref = ref if ref != '' else 'None' # ref is the address of the reference cell
        self.debug = False
        # map  excel functions onto their python equivalents
        self.funmap = FUNCTION_MAP
        self.funmap_string = FUNCTION_MAP_STRING
        
    def emit(self,ast,context=None, sp = None, mode = 'function'):
        fun = self.tvalue.lower()

        # Get the arguments
        args = self.children(ast)

        if fun == "atan2":
            # swap arguments
            if mode == 'function':
                return lambda: atan2(*tuple([args[1].emit(ast,context=context, sp = sp, mode = mode)(), args[0].emit(ast,context=context, sp = sp, mode = mode)()]))
            else:
                return "atan2(%s,%s)" % (args[1].emit(ast,context=context, mode = mode),args[0].emit(ast,context=context, mode = mode))
        elif fun == "pi":
            # constant, no parens
            if mode == 'function':
                return lambda: pi
            else:
                return "pi"
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
                if mode == 'function':
                    return lambda: RangeCore.filter(self.eval_ref(range), args[0].emit(ast,context=context, sp = sp, mode = mode)())
                else:
                    return 'RangeCore.filter(self.eval_ref("%s"), %s)' % (range, args[0].emit(ast,context=context, mode = mode))
            if len(args) == 2:
                if mode == 'function':
                    return lambda: args[1].emit(ast,context=context, sp = sp, mode = mode)() if args[0].emit(ast,context=context, sp = sp, mode = mode)() else 0
                else:
                    return "%s if %s else 0" %(args[1].emit(ast,context=context, mode = mode),args[0].emit(ast,context=context, mode = mode))
            elif len(args) == 3:
                if mode == 'function':
                    return lambda: args[1].emit(ast,context=context, sp = sp, mode = mode)() if args[0].emit(ast,context=context, sp = sp, mode = mode)() else args[2].emit(ast,context=context, sp = sp, mode = mode)
                else:
                    return "(%s if %s else %s)" % (args[1].emit(ast,context=context, mode = mode),args[0].emit(ast,context=context, mode = mode),args[2].emit(ast,context=context, mode = mode))
            else:
                raise Exception("if with %s arguments not supported" % len(args))

        elif fun == "array":

            # careful with multiple rows arrays
            if mode == 'function':
                return lambda: [arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args]
            else:
                my_str = '['
                if len(args) == 1:
                    # only one row
                    my_str += args[0].emit(ast,context=context, mode = mode)
                else:
                    # multiple rows
                    my_str += ",".join(['[' + n.emit(ast,context=context, mode = mode) + ']' for n in args])
                         
                my_str += ']'

                return my_str
        elif fun == "arrayrow":
            #simply create a list

            # might not be correct
            if mode == 'function':
                return lambda: [arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args]
            else:
                return ",".join([n.emit(ast,context=context, mode = mode) for n in args])

        elif fun == "and":
            if mode == 'function':
                return lambda: all([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args])
            else:
                return "all([" + ",".join([n.emit(ast,context=context, mode = mode) for n in args]) + "])"
        elif fun == "or":
            if mode == 'function':
                return lambda: any([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args])
            else:
                return "any([" + ",".join([n.emit(ast,context=context, mode = mode) for n in args]) + "])"
        elif fun == "index":
            if self.parent(ast) is not None and self.parent(ast).tvalue == ':':
                if mode == 'function':
                    return lambda: index(*tuple([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args]))
                else:
                    return 'index(' + ",".join([n.emit(ast,context=context, mode = mode) for n in args]) + ")"
            else:
                if mode == 'function':
                    return lambda: sp.eval_ref(index(*tuple([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args])), ref = self.ref) 
                else:
                    return 'self.eval_ref(index(%s), ref = %s)' % (",".join([n.emit(ast,context=context, mode = mode) for n in args]), self.ref)
        elif fun == "offset":
            if self.parent(ast) is None or self.parent(ast).tvalue == ':':
                if mode == 'function':
                    return lambda: offset(*tuple([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args]))
                else:
                    return 'offset(' + ",".join([n.emit(ast,context=context, mode = mode) for n in args]) + ")"
            else:
                if mode == 'function':
                    return lambda: sp.eval_ref(offset(*tuple([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args])), ref = self.ref)
                else:
                    return 'self.eval_ref(offset(%s), ref = %s)' % (",".join([n.emit(ast,context=context, mode = mode) for n in args]), self.ref)
        else:
            # map to the correct name
            if mode == 'function':
                f = self.funmap.get(fun,fun)
                return lambda: f(*tuple([arg.emit(ast,context=context, sp = sp, mode = mode)() for arg in args]))
            else:
                f = self.funmap_string.get(fun,fun)
                return f + "(" + ",".join([n.emit(ast,context=context, mode = mode) for n in args]) + ")"
