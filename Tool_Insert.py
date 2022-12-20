import pyodbc

class InsertTool:
    def __init__(self,Server,Port,DataBase,UserName,PassWord):
        self.cnxn = pyodbc.connect('DRIVER={SQL Server};SERVER='+Server+','+Port+';DATABASE='+DataBase+';UID='+UserName+';PWD='+ PassWord)
        self.cursor = self.cnxn.cursor()

    def load_Datas(self,Table,args,columns=[]):
        self.Table = Table
        self.args=args
        self.Columns = columns

    def Insert(self,Mode='Default'):
        if (Mode == 'Default'):
            self.default_Insert()
            self.commit()
        elif (Mode == 'Specific'):
            self.specific_Insert()
            self.commit()

    def default_Insert(self):
        sql ='INSERT INTO {} VALUES ({})'
        for i in self.args:
            print('Table:',self.Table)
            print('Data:',i)
            self.cursor.execute(sql.format(self.Table,('?,'*len(i))[:-1]), i)

    def specific_Insert(self):
        sql = 'INSERT INTO {} ({}) VALUES ({})'
        for i in self.args:
            print('Table:', self.Table)
            print('Data:', i)
            self.cursor.execute(sql.format(self.Table,str(self.Columns)[1:-1].replace('\'','').replace(' ',''),('?,'*len(i))[:-1]), i)

    def commit(self):
        self.cnxn.commit()
        print('Data have set in MS SQL SERVER!')