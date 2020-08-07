import json

import xlwt


class Excel:
    def __init__(self, path, sheet_names, values):
        self.workbook = xlwt.Workbook()
        self.sheet_names = sheet_names
        self.values = values
        self.path = path
        self.head = ["科目代码", "预算数", "决算数"]
        pass

    def write_excel(self):
        if isinstance(self.values, list):
            for sheet_name, value in zip(self.sheet_names, self.values):
                index = len(value)
                sheet = self.workbook.add_sheet(sheet_name)
                print(value)
                for i in range(0, index + 1):
                    for j in range(0, len(value[0])):
                        if i == 0:
                            print(self.head[j])
                            sheet.write(i, j, self.head[j])
                        else:
                            print(value[i - 1][j])
                            sheet.write(i, j, value[i - 1][j])
            self.workbook.save(self.path)
        elif isinstance(self.values, dict):
            for sheet_name in self.sheet_names:
                sheet = self.workbook.add_sheet(sheet_name)
                index = 1
                for j in range(0, len(self.head)):
                    # print(self.head[j])
                    sheet.write(0, j, self.head[j])
                data = self.values.get(sheet_name)
                for item in data:
                    if isinstance(item, float):
                        sheet.write(index, 0, int(item))
                    else:
                        sheet.write(index, 0, item)
                    sheet.write(index, 1, data.get(item).get("预算数") if data.get(item).get("预算数") else 0)
                    sheet.write(index, 2, data.get(item).get("决算数") if data.get(item).get("决算数") else 0)
                    index += 1

            # shit = load_json('2019.json')
            # shita = shit.get('一般公共预算收支科目')
            # for item in shita:
            #     print(item, shita.get(item).get("预算数"))

            # sheet.write(i, j, value[i - 1][j])
        self.workbook.save(self.path)


if __name__ == '__main__':
    print()
    # 测试问题
    path = r"C:\Users\localhost\Desktop\测试.xls"
    # sheet_name = ["测试1", "测试2"]
    # values = [[[1, 2, 3], [4, 5, 6], [7, 8, 9], [17, 18, 19]], [[7, 8, 9], [17, 18, 19]]]
    # print(len(values))
    with open('2019.json', 'r', encoding='utf-8') as f:
        values = json.load(f)
    sheet_name_list = ["一般公共预算收支科目", "政府性基金预算收支科目", "国有资本经营预算收支科目", "社会保险基金预算收支科目", "支出经济分类科目"]

    excel_writer = Excel(path, sheet_name_list, values)
    excel_writer.write_excel()
