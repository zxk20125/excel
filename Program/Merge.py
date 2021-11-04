import xlrd
import xlwt
import os


def getpath():
    path = os.path.abspath(os.path.join(os.getcwd(), "../file"))
    files = []

    for f in os.listdir(path):
        if f.endswith(".xlsx"):
            path = path + '\\'+f
            print(path)
        files.append(path)
    return files


def read_excel():
    for file in getpath():
        workBook = xlrd.open_workbook(file)
        allSheetNames = workBook.sheet_names()
        print(allSheetNames)


if __name__ == '__main__':
    read_excel()
