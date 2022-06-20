import openpyxl

files = []
autoincrement_id = 0


def Get_List_Files_Xlsx():
    return files


def increment_id():
    global autoincrement_id
    autoincrement_id += 1

    
def search_file(id):
    for file in files:
        if file['id'] == id:
            return file
    raise Exception('[Fetch Xlsx Library] : Id not found, id = ' + str(id))


def Fetch_Xlsx_To_List(id, row_start = 1, col_start = 1):
    file = search_file(id)
    return file['xlsx'].fetchXlsxToList(row_start, col_start)


def Fetch_Xlsx_To_Dict(id, header_row = 1, col_start = 1):
    file = search_file(id)
    return file['xlsx'].fetchXlsxToDict(header_row, col_start)


def Open_Xlsx(file_name, sheet_name = 'Sheet1', id = autoincrement_id):
    global files

    increment_id()
    files.append({
        'id': str(id),
        'xlsx': ManageXlsx(file_name, sheet_name),
    })


def Close_Xlsx(self, id):
    global files

    file = search_file(id)
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


    def fetchXlsxToDict(self, header_row, col_start):
        global files

        self.validateFile()

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


# Open_Xlsx('./test/data/calc2num.xlsx', 'Add', 'calc1')
# print(
#     Fetch_Xlsx_To_List('calc1', row_start = 2, col_start = 1)
# )
# print()
# Open_Xlsx('./test/data/calc2num.xlsx', 'ColumnTest', 'calc2')
# print(
#     Fetch_Xlsx_To_Dict('calc2', header_row=2, col_start=2)
# )
