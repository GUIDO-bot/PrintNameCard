from docxtpl import DocxTemplate
import xlrd

# 文件名改成要用的excel
excel = xlrd.open_workbook('2020福布斯中国富豪榜名单列表.xlsx')

sheet = excel.sheets()[0]
print("共", sheet.nrows, "行")

# names 是姓名，后面那个可以不要
names = []
categories = []

for i in range(sheet.nrows):
    names.append(sheet.cell_value(i, 0))
    categories.append(sheet.cell_value(i, 3))

# 设立一个保存路径,创建一个名为out1的文件夹存放结果，这个文件夹命名注意每次都要换，不能重名
filename = '.\out1'
import os
os.mkdir(filename)

for name,category in zip(names, categories):
    # template.docx是模板的名字，记得加后缀
    doc = DocxTemplate("template.docx")
    # context = {'name': name,'category':category}
    context = {'name': name}
    doc.render(context)
    doc.save(filename +'\\'+ name + ".docx")
    print(name + ".docx finished")
