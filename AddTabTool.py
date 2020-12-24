import xlwings as xw

temp = xw.Book(r'r3template.xlsx')
r3template = wb.sheets['Round 3']

agh = xw.Book(r'testresponses/GHRP_IOM_Indicators_Data_Collection_Template - Afghanistan.xlsx')
new_wb.sheets['Round 3'] = r3template
new_wb(r'new.xlsx')