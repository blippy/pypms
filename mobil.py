# Prepare Mobil work statement

class config:
    def __init__(self):
        fileName = 'M:\Finance\camel\mobil.txt'
        self.fin = file(fileName)
        self.lines = self.fin.readlines()
        self.fin.close()
    
    def read(self):
        line = self.lines[0].strip()
        self.lines = self.lines[1:]
        return line        
    
    def more(self): return len(self.lines) > 0
        
        

def main():
    jobs = {}
    conf = config()
    while conf.more():
        line = conf.read()
        print '*' , line , '*'
        if line == 'job':
            job = {}
            for field in ['job', 'contact', 'title', 'desc', 'po' , 'units']:
                job[field] = conf.read()
            jobs[job['job']] = job
            
            
    print jobs
    # FIXME

if  __name__ == "__main__":
    main()
    print 'Finished'