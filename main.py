import xlrd as xlrd
import math
import xlsxwriter as xlsxwriter

abscissa = []
ordinate = []
PointList = []
# 定义点函数
class Point:
    def __init__(self, x=0, y=0):
        self.x = x
        self.y = y
    # 定义横纵坐标的get方法
    def getX(self):
        return self.x
    def getY(self):
        return self.y
    def __str__(self):
        return '(%s,%s)' %(self.x,self.y)

# 获取数据
def getData():
    # 声明全局变量 abscissa --横坐标  ordinate --纵坐标
    global abscissa
    global ordinate
    worksheet = xlrd.open_workbook("./resource/data.xls")
    sheet_names = worksheet.sheet_names()
    for sheet_name in sheet_names:
        sheet2 = worksheet.sheet_by_name(sheet_name)
        # 横坐标
        abscissa = sheet2.col_values(1)
        # 纵坐标
        ordinate = sheet2.col_values(2)
    # 移除表头 第二列，第三列的列头
    del abscissa[0]
    del ordinate[0]
    # 打印第二列第三列数据
    print(abscissa)

    print(ordinate)


def createPoint():
    for i in range(0,len(abscissa)):
        point = Point(abscissa[i], ordinate[i])
        PointList.append(point)


def createWrite():
    workbook = xlsxwriter.Workbook('e:\dataResult.xlsx')  # 创建一个Excel文件
    worksheet = workbook.add_worksheet()
    title = [U'第一个点', U'第二个点', U'距离']  # 表格title
    worksheet.write_row('A1', title)
    workbook.close()
# 计算两点组合的长度
def getLen():
    # 创建输出文件
    workbook = xlsxwriter.Workbook('e:\dataResult.xlsx')  # 创建一个Excel文件
    worksheet = workbook.add_worksheet()
    title = [U'第一个点', U'第二个点', U'距离']  # 表格title
    worksheet.write_row('A1', title)
    num0 = 1
    for i in range(0, len(abscissa)):
        point = PointList[i]
        for j in range(i+1, len(abscissa)):
            next_point = PointList[j]
            # 计算横坐标之差
            distance_x = point.getX() - next_point.getX()
            # 计算纵坐标之差
            distance_y = point.getY() - next_point.getY()
            # 计算两个点的距离
            distance_point = math.sqrt((distance_x**2)+(distance_y**2))
            data = [point.__str__(), next_point.__str__(), distance_point]
            num = num0 + 1
            row = 'A' + str(num)
            worksheet.write_row(row, data)
            num0=num
            print("frist point %s , second point %s . The distance between these two points is %s" %(point.__str__(),next_point.__str__(),distance_point))
    workbook.close()



if __name__ == '__main__':
    getData()
    createPoint()
    getLen()
