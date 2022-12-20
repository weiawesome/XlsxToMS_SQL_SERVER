import pyodbc


class UpdateTool:
    def __init__(self, Server, Port, DataBase, UserName, PassWord):
        self.cnxn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=' + Server + ',' + Port + ';DATABASE=' + DataBase + ';UID=' + UserName + ';PWD=' + PassWord)
        self.cursor = self.cnxn.cursor()

    def load_Datas(self, Table,key,columns, args):
        self.Table = Table
        self.args = args
        self.Columns = columns
        self.key=key

    def Update(self):
        Sets=''
        for i in self.Columns:
            Sets+=(i+'=?,')
        sql='UPDATE {} SET {} WHERE {}=?'.format(self.Table,Sets[:-1],self.key)
        for i in self.args:
            print('Table:', self.Table)
            print('Data:', i)
            self.cursor.execute(sql, i)
        self.commit()


    def commit(self):
        self.cnxn.commit()
        print('Data have set in MS SQL SERVER!')