import cProfile
import io
import pstats
import string

import excel
import pydra
from common import princ

__raw_profile_file = '../profresults.dat'


def create_raw_profile(exe):
    global __raw_profile_file
    cProfile.run(exe, __raw_profile_file)
    
def create_readable_profile():
    global __raw_profile_file
    text_file = '..\\profile.txt'
    with io.open(text_file, mode = 'wb') as ostream:
        p = pstats.Stats(__raw_profile_file, stream = ostream)
        #p.sort_stats('time')
        p.sort_stats('cumulative')
        p.print_stats()
    #ostream.close()
    
    # rectify the line-endings
    with open(text_file, mode = "rb") as f:
        text = f.read()
        text = string.replace(text, "\n", "\r\n")        
    with open(text_file, mode = "wb") as f:
        f.write(text)

def main():
    cmd = 'excel.load()'
    cmd = 'pydra.main()'
    create_raw_profile(cmd)    
    create_readable_profile()
    
if  __name__ == "__main__":
    main()
    print 'Finished'
