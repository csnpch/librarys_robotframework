import openpyxl

files = []
autoincrement_id = 0




def increment_id():
    global autoincrement_id
    autoincrement_id += 1

    
def search_file(doc_id):
    for file in files:
        if file['doc_id'] == doc_id:
            return file
    raise Exception('[Fetch Xlsx Library] : doc_id not found, doc_id = ' + str(doc_id))


def Get_List_Files_Xlsx():
    return files


def Fetch_Xlsx_To_List(doc_id, row_start = 1, col_start = 1):
    file = search_file(doc_id)
    return file['xlsx'].fetchXlsxToList(row_start, col_start)


def Fetch_Xlsx_To_Dict(doc_id, header_row = 1, col_start = 1):
    file = search_file(doc_id)
    return file['xlsx'].fetchXlsxToDict(header_row, col_start)


def Write_To_Cell(doc_id, cellName, value, sheet_name=None):
    file = search_file(doc_id)
    return file['xlsx'].write2Cell(row = None, col = None, value=value, sheet_name=sheet_name, how='cell', cellName=cellName)


def Write_To_Row_Col(doc_id, row, col, value, sheet_name=None):
    file = search_file(doc_id)
    return file['xlsx'].write2Cell(row, col, value, sheet_name, 'row_col')


def Swap_Sheet(doc_id, sheet_name):
    file = search_file(doc_id)
    return file['xlsx'].swapSheet(sheet_name)


def Open_Xlsx(file_name, sheet_name = 'Sheet1', doc_id = autoincrement_id):
    global files

    increment_id()
    files.append({
        'doc_id': str(doc_id),
        'xlsx': ManageXlsx(file_name, sheet_name),
    })


def Close_Xlsx(self, doc_id):
    global files

    file = search_file(doc_id)
    files.remove(file)


def Close_All_Xlsx():
    global files

    files = []





class ManageXlsx:


    def __init__(self, file_name, sheet_name):
        self.file_name = file_name
        self.sheet_name = sheet_name
        self.book = openpyxl.load_workbook(self.file_name, data_only=True)
        self.sheet = self.book[self.sheet_name]


    def validateFile(self):
        if not self.file_name:
            raise Exception('[Fetch Xlsx Library] : File name not found')
        if not self.sheet_name:
            raise Exception('[Fetch Xlsx Library] : Sheet name not found')


    def fetchXlsxToList(self, row_start, col_start):
        global files

        self.validateFile()
        try:
            i = 0
            data = []
            rows = list(self.sheet.rows)
            for row in rows[row_start - 1:]:
                data.append([])
                for cell in row[col_start - 1:]:
                    if (not cell.value):
                        cell.value = ''
                    data[i].append(cell.value)
                i += 1
            return data
        except Exception as e:
            raise Exception('[Fetch Xlsx Library] : ' + str(e))


    def fetchXlsxToDict(self, header_row, col_start):
        global files

        self.validateFile()
        try:
            rows = list(self.sheet.rows)
            headerKey = list(rows[header_row - 1])
            for i in range(len(headerKey)):
                headerKey[i] = headerKey[i].value

            i = 0
            data = []
            for row in rows[header_row:]:
                data.append({})
                for cell in row[col_start - 1:]:
                    if (not cell.value):
                        cell.value = ''
                    key = headerKey[cell.column - 1]
                    if not key:
                        key = 'none_key'
                    data[i][key] = cell.value
                i += 1
            return data
        except Exception as e:
            raise Exception('[Fetch Xlsx Library] : ' + str(e))


    def swapSheet(self, sheet_name):
        global files

        self.validateFile()
        try:
            self.sheet = self.book[sheet_name]
            print('[Fetch Xlsx Library] : Swap sheet, sheet_name = ' + str(sheet_name))
            return True
        except Exception as e:
            raise Exception('[Fetch Xlsx Library] : ' + str(e))


    def write2Cell(self, row, col, value, sheet_name, how, cellName):
        global files

        self.validateFile()
        if (sheet_name):
            self.swapSheet(sheet_name)

        try:
            if (how == 'cell'):
                self.sheet[cellName] = value
            elif (how == 'row_col'):
                self.sheet.cell(row=row, column=col).value = value
            self.book.save(self.file_name)
            return True
        except Exception as e:
            raise Exception('[Fetch Xlsx Library] : ' + str(e))



# Open_Xlsx('./test/data/calc2num.xlsx', 'Add', 'calc1')
# Open_Xlsx('./test/data/calc2num.xlsx', 'ColumnTest', 'calc2')

# print('\nGet_List_Files_Xlsx\n', Get_List_Files_Xlsx(), '\n')

# print(
#     Fetch_Xlsx_To_List('calc1', row_start = 2, col_start = 1)
# )
# print()


# Write_To_Row_Col('calc2', 4, 5, 'insertCell')
# Write_To_Cell('calc2', 'E3', 'insertCell')
# Write_To_Cell('calc2', 'E4', 'insertCell')

# print(
#     Fetch_Xlsx_To_Dict('calc2', header_row=2, col_start=2)
# )
