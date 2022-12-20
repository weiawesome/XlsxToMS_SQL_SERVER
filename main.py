import Tool_Create
import Tool_Insert
import pandas as pd

ExcelFile = './DataBase(Excel)/database.xlsx'
SpecificTable = ''
Create=False
Server = 'localhost'
DataBase = ''
Port = '1433'
UserName = ''
PassWord = ''

CreateTools=Tool_Create.CreateTool(Server = Server,DataBase = DataBase,Port = Port,UserName = UserName,PassWord = PassWord)
InsertTools = Tool_Insert.InsertTool(Server = Server,DataBase = DataBase,Port = Port,UserName = UserName,PassWord = PassWord)

def ReadSpecific(Table):
    df = pd.read_excel(ExcelFile, sheet_name=Table, dtype=str)
    if(Create):
        CreateTools.Create(Table,list(df.columns.values))
    args = list(df.values)
    for index, val in enumerate(args):
        args[index] = list(val)

    InsertTools.load_Datas(Table,args)
    InsertTools.Insert()


def ReadALL():
    xls = pd.ExcelFile(ExcelFile)
    Df = dict()

    for Table in xls.sheet_names:
        Df[Table] = pd.read_excel(ExcelFile, Table, dtype=str)
        if(Create):
            CreateTools.Create(Table,list(Df[Table].columns.values))

    for Table in Df:
        args = list((Df[Table]).values)
        for index, val in enumerate(args):
            args[index] = list(val)
        InsertTools.load_Datas(Table,args)
        InsertTools.Insert()

def main():
    if SpecificTable == '':
        ReadALL()
    else:
        ReadSpecific(SpecificTable)

if __name__ == '__main__':
    main()
