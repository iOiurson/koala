import pyximport; pyximport.install()

import unittest
import os
import sys
import json
from datetime import datetime

dir = os.path.dirname(__file__)
path = os.path.join(dir, '../..')
sys.path.insert(0, path)

from koala.reader import read_archive, read_named_ranges, read_cells
from koala.ExcelCompiler import ExcelCompiler
from koala.Spreadsheet import Spreadsheet
from koala.Cell import Cell
from koala.Range import RangeCore
from koala.excellib import *


file_name = "./tests/ast/big_formula.xlsx"

c = ExcelCompiler(file_name, debug = True)
# c.clean_volatile()
sp = c.gen_graph()

# sp.dump('big_formula.gzip')
# sp = Spreadsheet.load('big_formula.gzip')
sp.mode = 'string'
sp.set_value('Sheet1!A1', 10)

startTime = datetime.now()
print 'String mode', sp.evaluate('Sheet1!G2')
print "___Timing___  Eval done in %s" % (str(datetime.now() - startTime))

sp.mode = 'function'
sp.set_value('Sheet1!A1', 1)
sp.set_value('Sheet1!A1', 10)


startTime = datetime.now()
print 'Function mode', sp.evaluate('Sheet1!G2')
print "___Timing___  Eval done in %s" % (str(datetime.now() - startTime))
