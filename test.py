import xlrd
import xlwt

workbook = xlrd.open_workbook("math.xlsx")
# 获取sheet对象
sheets_object = workbook.sheets()
# 通过index获取第一个sheet对象
sheet1_object = workbook.sheet_by_index(0)
# 通过index判断sheet1是否导入
sheet1_is_load = workbook.sheet_loaded(sheet_name_or_index=0)
print(sheet1_is_load)


# 获取sheet1中的有效行数
nrows = sheet1_object.nrows

counts={}
for i in range(nrows):
    all_row_values = sheet1_object.row_values(rowx=i)
    if all_row_values[0]=='E1':
        m=all_row_values[1]
        if m in counts:
            counts[m]=counts[m]+1
        else:
            counts[m]=1
pairs=list(counts.items())
pair=[]
for i in pairs:
    li=list(i)
    a=li[0]
    li[0]=li[1]
    li[1]=a
    pair.append(li)
        
result=0   
for list1 in pair:
    m=list1[0]
    if m>=0:
        result=result+1
print(result)
result=0
print(counts)

