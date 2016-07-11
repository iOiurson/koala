
from datetime import datetime
import sys
import resource

sys.setrecursionlimit(30000)
limit = 67104768 # maximum stack limit on my machine => use 'ulimit -Ha' on a shell terminal
resource.setrlimit(resource.RLIMIT_STACK, (limit, limit))

def sum_r(table):

	if len(table) > 0:
		a = table.pop()

		return a + sum_r(table)
	else:
		return 0


s = compile('sum_e(table)', '<string>','eval')

def sum_e(table):

	if len(table) > 0:
		a = table.pop()

		# return a + eval('sum_e(table)')
		return a + eval(s)
	else:
		return 0


l = range(10000)

start = datetime.now()
value = sum_r(l)
print 'Without eval: %s calculated in %s' % (value, str(datetime.now() - start))

l = range(10000)

start = datetime.now()
value = sum_e(l)
print 'With eval: %s calculated in %s' % (value, str(datetime.now() - start))


"self.eval_ref('Sheet1!A1')"

def f():
	return self.eval_ref('Sheet1!A1')


