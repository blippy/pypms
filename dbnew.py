import copy

import adodbapi

###########################################################################

class TableCache:
    def __init__(self, table_name):
        self.table_name = table_name
        self.db_rows = [] # the rows that are in the database        
        self.populate()
        self.inmemory_rows = copy.deepcopy(self.db_rows) # the rows held in memory
    
    def populate(self):            
        'Return an open connection to the database'
        conStr = r'PROVIDER=Microsoft.Jet.OLEDB.4.0;DATA SOURCE=M:/Finance/camel/camel.mdb;'    
        con = adodbapi.connect(conStr)
        cursor = con.cursor()
        sql = "SELECT * FROM " + self.table_name
        cursor.execute(sql)
        ds = cursor.fetchall()
        self.db_rows = [row for row in ds]
        cursor.close()
        con.close()