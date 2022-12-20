import pyodbc

class CreateTool:
    def __init__(self, Server, Port, DataBase, UserName, PassWord):
        self.cnxn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=' + Server + ',' + Port + ';DATABASE=' + DataBase + ';UID=' + UserName + ';PWD=' + PassWord)
        self.cursor = self.cnxn.cursor()

    def DropTable(self,Table):
        sql='DROP TABLE {}'.format(Table)
        self.execute(sql)
        print('Table {} has benn Drop'.format(Table))
        print()

    def Create(self,Table,columns,Primary_key='',default_Type='nvarchar(max)'):
        try:
            self.DropTable(Table)
        except:
            pass
        tmp=''
        for i in columns:
            tmp+=(i+' '+default_Type+',') if i!=Primary_key else (i+' '+default_Type+' PRIMARY KEY,')
        sql='CREATE TABLE {} ({})'.format(Table,tmp[:-1])
        print('Table:',Table)
        print('Columns:',columns)
        print('PrimaryKey:',Primary_key)
        print('SQL:',sql)
        self.execute(sql)
        self.commit()

    def execute(self,sql):
        self.cursor.execute(sql)

    def commit(self):
        self.cnxn.commit()
        print('Data have set in MS SQL SERVER!')