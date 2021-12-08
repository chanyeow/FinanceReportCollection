import openpyxl

def test():
    wb = openpyxl.load_workbook('cases.xlsx')
    sh = wb['Sheet']
    ce = sh.cell(row = 1,column = 1)   # 读取第一行，第一列的数据
    print(ce.value)
    print(list(sh.rows)[1:])     # 按行读取数据，去掉第一行的表头信息数据
    for cases in list(sh.rows)[1:]:
        case_id =  cases[0].value
        case_excepted = cases[1].value
        case_data = cases[2].value
        print(case_excepted,case_data)
    wb.close()

def create_excel():
    wb = openpyxl.Workbook()
    wb.create_sheet('test_case')
    wb.save('cases.xlsx')

if __name__ == '__main__':
    test()