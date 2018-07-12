# -*- coding: utf-8 -*-
import os, datetime
import re
import mysql.connector
import xlrd

def te(str):
    str = str.strip()
    str = str.replace(' ', 'x')
    str = str.lower()
    str = str.replace('(', '')
    str = str.replace(')', '')
    str = str.replace('%', 'P')
    str = str.replace("\"", "")
    str = str.replace("/", "")
    str = str.replace(r"\\", "")
    str = str.replace(".", "")
    str = str.replace("-", "_")
    str = str.replace("[", "")
    str = str.replace("]", "")
    str = str.replace(":", "")
    str = str.replace ("+" , "")
    str = str.replace("\n", "")
    return str
#去除字符串中的特殊字符，避免msql的命名问题
def insert_mysql( table_name2 , cursor , sql , parameter_list , ret ):
    try:
        cursor.execute (sql , parameter_list)
    except mysql.connector.errors.ProgrammingError as e:
        print (e)
        return -1
    except mysql.connector.errors.DataError as e:
        m = re.search ("'(.*)'" , str (e))
        str_need_modify = m.group (0)
        str_need_modify = str_need_modify[1:(len (str_need_modify) - 1)]
        sql0 = "alter table " + table_name2 + " modify column " + str_need_modify + " text"
        cursor.execute (sql0)
        insert_mysql (table_name2 , cursor , sql , parameter_list , ret)

class Excel_Msql:
    def __init__(self,sqlconfig,filepath):
        self.fig=sqlconfig
        self.path=filepath
        self.path_list = []
        self.table_list = []

    def getfile(self):
        for file_name in os.listdir(self.path):
            if os.path.isdir(os.path.join(self.path, file_name)):
                '''get the files in sub folder recursively'''
                son_filepath=Excel_Msql(self.fig,os.path.join(self.path, file_name))
                son_filepath.getfile()
                self.path_list.extend(son_filepath.path_list)
                self.table_list.extend(son_filepath.table_list)
            else:
                self.path_list.append(os.path.join(self.path, file_name))
                '''convert file name to mysql table name'''
                file_name = file_name.split('.')[0]  # remove .xls
                # file_name = file_name.split('from')[0] #remove characters after 'from'
                file_name = file_name.strip()  # remove redundant space at both ends
                file_name = file_name.replace(' ', '_')  # replace ' ' with '_'
                file_name = file_name.replace('-', '_')  # replace ' ' with '_'
                file_name = file_name.lower()  # convert all characters to lowercase
                self.table_list.append(file_name)
    def storeData(self,file_path, table_name, cursor):
        ret = 0
        '''open an excel file'''
        try:
            with xlrd.open_workbook(file_path) as file:
                for sheet_name in file.sheets():
                    sheet = file.sheet_by_name(sheet_name.name)
                    '''get the number of rows and columns'''
                    nrows = sheet.nrows
                    ncols = sheet.ncols
                    '''get column names'''
                    col_names = []
                    table_name2 = table_name + "_" + sheet_name.name
                    table_name2 = te(table_name2)
                    ct = 1
                    a = 1
                    for i in range(0, ncols):
                        title = sheet.cell(0, i).value
                        title = te(title)
                        if title == "" or title == " ":
                            title = "NULL" + str(a)
                            a = a + 1
                        col_names.append(title)
                    # 重复项加角标
                    for z in col_names:
                        if col_names.count(z) > 1:
                            col_names[col_names.index(z)] = col_names[col_names.index(z)] + str(ct)
                            ct = ct + 1
                    '''create table in mysql'''
                    sql = 'create table ' \
                          + table_name2 + ' (' \
                          + 'id int NOT NULL AUTO_INCREMENT PRIMARY KEY, '
                    for i in range(0, ncols):
                        if isinstance(col_names[i], str):
                            sql = sql + str(col_names[i]) + ' varchar(250)'
                        elif isinstance(col_names[i], float):
                            sql = sql + str(col_names[i]) + ' float(250)'
                        if i != ncols - 1:
                            sql += ','
                    sql = sql + ')'
                    try:
                        cursor.execute(sql)
                        sql = 'insert into ' + table_name2 + ' ('
                        for i in range(0, ncols - 1):
                            sql = sql + str(col_names[i]) + ', '
                        sql = sql + str(col_names[ncols - 1])
                        sql += ') values ('
                        sql = sql + '%s,' * (ncols - 1)
                        sql += '%s)'
                        # get parameters
                        parameter_list = []
                        for row in range(1, nrows):
                            for col in range(0, ncols):
                                cell_type = sheet.cell_type(row, col)
                                cell_value = sheet.cell_value(row, col)
                                if cell_type == xlrd.XL_CELL_DATE:
                                    try:
                                        dt_tuple = xlrd.xldate_as_tuple(cell_value, file.datemode)
                                    except xlrd.xldate.XLDateAmbiguous:
                                        BrokenFile_list.append(file_path)
                                    try:
                                        meta_data = str(datetime.datetime(*dt_tuple))
                                    except ValueError as e:
                                        BrokenFile_list.append(file_path)
                                        print(e)
                                else:
                                    meta_data = sheet.cell(row, col).value
                                parameter_list.append(meta_data)
                            # cursor.execute(sql, parameter_list)
                            insert_mysql (table_name2 , cursor , sql , parameter_list , ret)
                            parameter_list = []
                            ret += 1
                    except mysql.connector.errors.ProgrammingError as e:
                        print(e)
                        # return -1
                    '''insert data'''
                    # construct sql statement
        except xlrd.biffh.XLRDError as e:
            print(e, "\n", file_path, "is broken")
        '''get the first sheet'''
        return ret
    def datahelper(self):
        '''import data helper'''
        '''
        Step 0: Validate input database parameters
        '''
        try:
            conn=mysql.connector.connect(**self.fig)
            '''
                    Step 1: Traverse files in datapath, store file paths and corresponding table names in lists
                    lists[0] is the list of files paths
                    lists[1] is the list of table names
                    '''
            self.getfile()
            nfiles = len (self.table_list)
            '''
            Step 2: Store data in mysql via a for-loop
            '''
            cursor = conn.cursor ()
            for file_idx in range (0 , nfiles):
                file_path = self.path_list[file_idx]
                print ("processing file(%d/%d):[ %s ]" % (file_idx + 1 , nfiles , file_path))
                table_name = self.table_list[file_idx]
                num = self.storeData(file_path , table_name , cursor)
                if num >= 0:
                    print ("[ %d ] data have been stored in TABLE:[ %s ]" % (num , table_name))
                conn.commit ()
            cursor.close ()
            conn.close()
        except mysql.connector.errors.ProgrammingError as e:
            print (e)
            return -1

if __name__=='__main__':
    xpath="/home/user/文档/111"
    sql_path={"host":"localhost",
              "user":"root",
              "port":3306,
              "password":"123456",
              "db":"移动考试",
              "charset" : 'utf8'
              }
    a=Excel_Msql(sql_path,xpath)
    a.datahelper()
