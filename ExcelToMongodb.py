# -*- coding: utf-8 -*-
import xlrd
from pymongo import MongoClient


class ExcelToMongodb():

    def __init__(self, path, sheet=0, db='Temp', collection='temp'):
        """
        功能：execl 存入 mongodb  \n
        :param path:  excel存放路径  \n
        :param sheet: 工作表号，从0开始 \n
        :param db:   数据库名 \n
        :param collection:  集合名\n
        """
        self.path = path
        self.sheet = sheet
        self.db = db
        self.collection = collection

    def read_excel(self):

        excel_file = xlrd.open_workbook(self.path)

        try:
            # 根据sheet索引或者名称获取sheet内容
            sheet1 = excel_file.sheet_by_index(self.sheet)

        except Exception as e:
            print("没有此表号!!!")
            print("%s 中有如下所示表:" % self.path)
            # 获取目标EXCEL文件sheet名
            print(excel_file.sheet_names())
            return False

        # 获取整行的值
        title_col = sheet1.row_values(0)

        # 学生信息列表，表中每一个元素为字典
        stu = []

        for i in range(1, sheet1.nrows):
            col_value = sheet1.row_values(i)
            temp = {}
            for j in range(sheet1.ncols):
                temp[title_col[j]] = col_value[j]
            stu.append(temp)

        # self.pretty(stu)
        return stu, sheet1.nrows - 1

        # sheet的名称，行数，列数
        # print(sheet1.name, sheet1.nrows, sheet1.ncols)

    def write_mongodb(self, tup):

        if tup is False:
            return

        conn = MongoClient()

        # 连接mydb数据库，没有则自动创建
        database = conn.get_database(self.db)

        # 使用test_set集合，没有则自动创建
        my_set = database.get_collection(self.collection)
        for dict in tup[0]:
            my_set.insert(dict)

        print("*" * 50)
        print("%d 条数据已存入 %s 数据库, %s 集合中" % (tup[1], self.db, self.collection))
        print("*" * 50)
        conn.close()

    def excel_mongodb(self):
        self.write_mongodb(self.read_excel())


if __name__ == '__main__':
    path = r'd:\cj.xlsx'
    file = ExcelToMongodb(path)
    file.excel_mongodb()
