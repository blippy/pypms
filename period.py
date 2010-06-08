import datetime, msvcrt, os, sys, traceback

import common

class Period:
    def __init__(self, usePrev = False):        
        self.setToCurrent()        
        if usePrev: self.decMonth()

    def setToCurrent(self):
        t = datetime.datetime.now()
        self.y = t.year
        self.m = t.month
        
    def describe(self):
        print 'YEAR: ', self.y
        print 'MONTH:', self.m
        

    def ask(self, text, default):
        inp = raw_input(text)
        if len(inp) == 0: inp = default
        return inp

    def askNumber(self, fieldDescription, default, low, high):
        text = '%s (%d-%d) [%d]:' % (fieldDescription, low, high, default)
        while 1:
            try:
                value = self.ask(text, default)
                value = int(value)
                if value < low or value > high: raise ValueError
                break
            except ValueError:
                print "Invalid input. Try again"
        return value

    def askYear(self):
        return self.askNumber('Year', datetime.datetime.now().year, 2000, 2020)
            
    def askMonth(self):
        return self.askNumber('Month', datetime.datetime.now().month, 1, 12)

    def decMonth(self):
        self.m -= 1
        if self.m < 1:
            self.m = 12
            self.y -= 1
            
    def inputPeriod(self):
        self.setToCurrent()
        while 1:
            inp = raw_input("Use (r)ecent (c)urrent or (o)ther billing period ([r]/c/o)?")
            if inp == "" or inp =="r":
                self.decMonth()                
                break
            elif inp == "c":
                break
            elif inp == "o":
                self.y = self.askYear()
                self.m = self.askMonth()
                break
            else: 
                print "Invalid input. Try again"    
        


    def outDir(self):
        full = "%s\\%s\\" % (common.outroot, self.yyyymm())
        full = os.path.abspath(full)
        common.makedirs(full)
        return full

    def within(self, d):
        return d.year == self.y and d.month == self.m

    def yyyymm(self):
        'Return the period in the form 2010-01'
        result = "%04d-%02d" % (self.y, self.m)
        return result
    
    def mmmmyyyy(self):
        ' Return the period in the form March 2010'
        d = datetime.date(self.y, self.m, 1)
        result = d.strftime('%B %Y')
        return result
    

