import openpyxl
import os

files = {"input": "", "output": ""}


def getpath():
    path = os.path.abspath(os.path.join(os.getcwd(), "../file"))
    for f in os.listdir(path):
        if f.endswith("target.xlsx"):
            files["input"] = path + '\\' + f
        elif f.endswith("output.xlsx"):
            files["output"] = path + '\\' + f


def get_subject():
    res = []
    wb = openpyxl.load_workbook(files["input"])
    print(files["input"])
    for name in wb.sheetnames:
        sheet = wb[name]
        for row in sheet.rows:
            for cell in row:
                if not res.count(cell.value) and cell.value is not None:
                    res.append(cell.value)
            break
    return res


def write_subject():
    wb = openpyxl.load_workbook(files["output"])
    sheet = wb["Sheet2"]
    sheet.append(get_subject())
    wb.save(files["output"])


def init_data():
    tmp = {}
    for item in get_subject():
        tmp[item] = []
    return tmp


def get_data():
    subject = get_subject()
    sub = ""
    data = init_data()
    wb = openpyxl.load_workbook(files["input"])
    for sheet in wb.sheetnames:
        if sheet == '报告':
            continue
        ws = wb[sheet]
        length = 0
        for col in ws.columns:
            for cell in col:
                if subject.count(cell.value):
                    sub = cell.value
                elif sub != "":
                    data[sub].append(str(cell.value))
                    if data[sub] != "":
                        length = len(data[sub])
            sub = ""
            for s in subject:
                if data[s] == "":
                    data[s].append([" "]*length)
    print(data['项目'])
    return data


def write_data():
    data = get_data()
    wo = openpyxl.load_workbook(files["output"])
    sheet2 = wo["Sheet1"]
    for subject in get_subject():
        sheet2.append(data[subject])
    wo.save(files["output"])


if __name__ == '__main__':
    getpath()
    # write_data()
    get_data()