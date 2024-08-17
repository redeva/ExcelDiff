import openpyxl as ol
from openpyxl.styles import PatternFill
from openpyxl.styles import Font

# 填充颜色
yellow_fill = PatternFill(patternType='solid', fgColor='FFFF00')  # 黄
green_fill = PatternFill(patternType='solid', fgColor='90EE90')  # 淡绿色
blue_fill = PatternFill(patternType='solid', fgColor='00CCFF')

# 字体颜色
red_font = Font(color='c00000')  # 红色字体

# 获取工作簿对象实例
book1 = ol.load_workbook('20240724_105656_浅层2.xlsx')
# book2 = ol.load_workbook('2.xlsx')

# 获取工作表对象实例
book1_sheet1 = book1['原始数据查看']
book1_sheet2 = book1['异常值剔除后完整列表']

# 定义两个list ，存放工作表待处理的单元格范围
book1_sheet1_rowlist = range(book1_sheet1.min_row, book1_sheet1.max_row)  # 起始行加1 ，跳过标题行
book1_sheet1_collist = range(book1_sheet1.min_column, book1_sheet1.max_column)

# 简单的循环一下，对比区域内的数值
for j in book1_sheet1_rowlist:
    for k in book1_sheet1_collist:
        value_a = book1_sheet1.cell(row=j, column=k).value
        value_b = book1_sheet2.cell(row=j, column=k).value
        if value_a != value_b:
            print('\n=============================\n')
            print(
                f'出大事了，{value_a}和{value_a}的数值不一样！它们位于第{j}行第{k}列，对应的样品编号是{book1_sheet1.cell(row=j, column=1).value}，对应测试指标是{book1_sheet1.cell(row=1, column=k).value}')
            # 设置字体颜色及填充
            book1_sheet1.cell(row=j, column=k).font = red_font
            book1_sheet1.cell(row=j, column=k).fill = green_fill

# 最后，将book1另存成一个文件
book1.save('a1.xlsx')

'''


raw_excel = 'excel_raw.xlsx'
result_excel = 'result.xlsx'
workbook = load_workbook(raw_excel)
writer = pd.ExcelWriter(result_excel , engine='openpyxl')
writer.book = workbook

df_1.to_excel(writer, sheet_name='test1_sheet')
df_2.to_excel(writer, sheet_name='test2_sheet')

writer.save()
writer.close()





'''
