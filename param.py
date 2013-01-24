# http://www.artima.com/weblogs/viewpost.jsp?thread=240845

# My topic on this is:
# https://groups.google.com/forum/?hl=en-GB&fromgroups=#!topic/comp.lang.python/S0V7IZjx-kQ
# equivalently: 
# http://is.gd/R28AK0

from functools import wraps

class Param(object):
    def __init__(self, default):
        self.default = default
        #self.func = func

    def __call__(self, func, *args):
        @wraps(func)
        def wrapper(*args):
            #print 'calling hi'
            if len(args) >0:
                self.default = args[0]
            #print len(args),  ' ', args
            return self.default # self.func(*args)
        return wrapper


@Param(42)
def hi(newval):
    pass

def example():
    print hi() # 42
    print hi(12) # 12
    print hi() # 12
    hi(13) # sets it to 13
    print hi() # 13
