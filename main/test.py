from data_operations.get_data import GetData
from tool.operation_excel import OperationExcel

excel = OperationExcel(file_name='D:/Lizhi/cps_company.xlsx',sheet_id=0)
data = excel.get_data()
for i in range(excel.get_lines()):
    rowcontent = ''
    for j in range(excel.get_cols()):
        rowcontent = rowcontent+str(excel.get_cell_value(row=i,col=j))+' '
    print(rowcontent)

data = GetData('D:/Lizhi/cps_company.xlsx')
data.opera_excel
print(data.get_case_lines())
