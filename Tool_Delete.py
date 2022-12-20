import pyodbc


class DeleteTool:
    def __init__(self, Server, Port, DataBase, UserName, PassWord):
        self.cnxn = pyodbc.connect(
            'DRIVER={SQL Server};SERVER=' + Server + ',' + Port + ';DATABASE=' + DataBase + ';UID=' + UserName + ';PWD=' + PassWord)
        self.cursor = self.cnxn.cursor()

    def load_Datas(self, Table,key, args):
        self.Table = Table
        self.args = args
        self.key=key

    def Delete(self):
        sql='DELETE FROM {} WHERE {}=?'.format(self.Table,self.key)
        for i in self.args:
            print('Table:', self.Table)
            print('Data:', i)
            self.cursor.execute(sql, i)
        self.commit()

    def commit(self):
        self.cnxn.commit()
        print('Data have set in MS SQL SERVER!')