import csv
from importlib.machinery import PathFinder
from logging import error

def Fetch_CSV_To_Dict(pathData = None):
    if (not pathData):
        raise Exception('[FetchCSV Library] : Need 1 argument "pathData" eg -> | Fetch Csv ${CURDIR}/data/fie.csv')
    
    data = []
    for row in csv.DictReader(open(pathData)):
        data.append(row)
    return data


def Fetch_CSV_To_List(pathData = None, rowStart = 1):
    if (not pathData):
        raise Exception('[FetchCSV Library] : Need 1 argument "pathData" eg -> | Fetch Csv ${CURDIR}/data/fie.csv')
    
    data = []
    for row in csv.reader(open(pathData)):
        data.append(row)
    return data[rowStart:]


# print('Dict : ',Fetch_CSV_To_Dict('./test/data/calc.csv'))
# print('\n\n')
# print('List : ', Fetch_CSV_To_List('./test/data/calc.csv'))
