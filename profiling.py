import cProfile
import io
import pstats

import excel
import pydra
from common import princ

__raw_profile_file = '../profresults.dat'


def create_raw_profile(exe):
    global __raw_profile_file
    cProfile.run(exe, __raw_profile_file)
    
def create_readable_profile():
    global __raw_profile_file
    ostream = io.open('..\\profile.txt', mode = 'wb')
    p = pstats.Stats(__raw_profile_file, stream = ostream)
    #p.sort_stats('time')
    p.sort_stats('cumulative')
    p.print_stats()
    ostream.close()

def main():
    cmd = 'excel.import_excel_data()'
    cmd = 'pydra.main()'
    create_raw_profile(cmd)    
    create_readable_profile()
    
if  __name__ == "__main__":
    main()
    print 'Finished'
