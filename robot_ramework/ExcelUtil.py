import openpyxl

def get_data(file, sheet_num=0, ignore_header=True, ignore_footer=True):
    workbook = openpyxl.load_workbook(file) 
    sheet = workbook._sheets[sheet_num]

    data = []
    row_num = 1
    for holding in sheet.values:
        holding = list(holding)
        holding.insert(0, row_num)
        data.append(holding)
        row_num += 1
    
    if ignore_header: data = data[1:]
    if ignore_footer: data = data[:-1]

    return data

def set_data(file, rowcol, data, sheet_num=0):
    workbook = openpyxl.load_workbook(file) 
    sheet = workbook._sheets[sheet_num]
    sheet[rowcol] = data
    workbook.save(file)


if __name__ == "__main__":
    file = "D:\\holdings.xlsx"

    holdings = get_data("D:\\holdings.xlsx")
    print(holdings)

    # set_cell_data(file, 'D2', "xxx")
