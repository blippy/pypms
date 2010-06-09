# Prepare Mobil work statement

class config:
    def __init__(self):
        fileName = 'M:\Finance\camel\mobil.txt'
        self.fin = file(fileName)
        self.lines = self.fin.readlines()
        self.fin.close()
    
    def __readLine(self):
        line = self.lines[0].strip()
        self.lines = self.lines[1:]
        return line        
    
    def __more(self): return len(self.lines) > 0

    def load(self):
        'Load in the configuration information'
        self.jobs = {}
        while self.__more():
            line = self.__readLine()
            print '*' , line , '*'
            if line == 'job':
                job = {}
                for field in ['job', 'contact', 'title', 'desc', 'po' , 'units']:
                    job[field] = self.__readLine()
                self.jobs[job['job']] = job
        
        

def main():
    conf = config()
    conf.load()                        
    print conf.jobs
    # FIXME

if  __name__ == "__main__":
    main()
    print 'Finished'