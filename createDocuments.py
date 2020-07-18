from libs.openpyxl import load_workbook
from libs.PyQt5 import QtCore
from part import Part
from operator import attrgetter


class CreateDocuments:
    def __init__(self, name='', orderDate=''):
        self.name = name
        self.orderDate = orderDate

    def closePlan(self, data=None, closeDate=''):
        if not data:
            pass
        for each in data:
            print(each)
        wb = load_workbook("form_copy.xlsx")
        sheets = wb.sheetnames
        Sheet1 = wb[sheets[0]]
        Sheet1.cell(row=7, column=1).value = self.name
        Sheet1.cell(row=7, column=8).value = closeDate
        # Set value to table
        rows = len(data)
        for row in range(rows):
            for column in range(1,3):
                Sheet1.cell(row=7+row, column=3+column).value = data[row][column]
            Sheet1.cell(row=7+row, column=2).value = data[row][0]
            Sheet1.cell(row=7+row, column=12).value = data[row][5]
            Sheet1.cell(row=7+row, column=15).value = data[row][3]
            Sheet1.cell(row=7+row, column=17).value = data[row][6]


        # keep only first Sheet
        for s in sheets:
            if s != 'B01':
                sheet_name = wb.get_sheet_by_name(s)
                wb.remove_sheet(sheet_name)
        return wb

    def embryosPlan(self, data=None):
        if not data:
            pass
        for each in data:
            print(each)
        wb = load_workbook("form_copy.xlsx")
        sheets = wb.sheetnames
        Sheet1 = wb[sheets[1]]
        # Set value to table
        rows = len(data)
        for row in range(rows):
            Sheet1.cell(row=6 + row, column=1).value = row + 1
            for column in range(0, 2):
                Sheet1.cell(row=6 + row, column=2 + column).value = data[row][column+1]
            for column in range(5,7):
                Sheet1.cell(row=6 + row, column=column).value = data[row][column - 2]

        # keep only second Sheet
        for s in sheets:
            if s != 'B02':
                sheet_name = wb.get_sheet_by_name(s)
                wb.remove_sheet(sheet_name)
        return wb

    def inventPlan(self, data=None, dateOfCheck='', closeDate='', closingDay_PXXK=''):
        if not data:
            pass
        for each in data:
            print(each)
        wb = load_workbook("form_copy.xlsx")
        sheets = wb.sheetnames
        Sheet1 = wb[sheets[2]]
        # Set value to table
        rows = len(data)
        Sheet1.cell(row=6, column=6).value = dateOfCheck
        Sheet1.cell(row=6, column=7).value = closeDate
        Sheet1.cell(row=6, column=9).value = closingDay_PXXK
        for each in data:
            print(each)

        for row in range(rows):
            Sheet1.cell(row=6 + row, column=1).value = row + 1
            Sheet1.cell(row=6 + row, column=2).value = data[row][1]
            Sheet1.cell(row=6 + row, column=3).value = data[row][0]
            Sheet1.cell(row=6 + row, column=4).value = data[row][3]
            Sheet1.cell(row=6 + row, column=5).value = data[row][5]



        # keep only second Sheet
        for s in sheets:
            if s != 'B03':
                sheet_name = wb.get_sheet_by_name(s)
                wb.remove_sheet(sheet_name)
        return wb

    def embryosPlanPart(self, data=None, data1=None):
        if not data:
            pass
        if not data1:
            pass
        for each in data:
            print(each)
        wb = load_workbook("form_copy.xlsx")
        sheets = wb.sheetnames
        Sheet1 = wb[sheets[4]]
        # Set value to table
        rows = len(data)
        for row in range(rows):
            Sheet1.cell(row=7 + row, column=2).value = data[row][1]
            Sheet1.cell(row=7 + row, column=3).value = data1[row][1]
            Sheet1.cell(row=7 + row, column=9).value = data[row][4]
            Sheet1.cell(row=7 + row, column=13).value = data[row][6]
            # Sheet1.cell(row=7 + row, column=9).value = data[row][4]

        # keep only four Sheet
        for s in sheets:
            if s != 'B05':
                sheet_name = wb.get_sheet_by_name(s)
                wb.remove_sheet(sheet_name)
        return wb

    def productPlan(self, data=None):
        if not data:
            pass
        for each in data:
            print(each)
        wb = load_workbook("form_copy.xlsx")
        sheets = wb.sheetnames
        Sheet1 = wb[sheets[5]]
        # Set value to table
        rows = len(data)
        for row in range(rows):
            Sheet1.cell(row=9 + row, column=1).value = data[row][0]
            Sheet1.cell(row=9 + row, column=2).value = data[row][1]
            Sheet1.cell(row=9 + row, column=3).value = data[row][9]
            Sheet1.cell(row=9 + row, column=5).value = data[row][7]
            Sheet1.cell(row=9 + row, column=8).value = data[row][5]
            Sheet1.cell(row=9 + row, column=9).value = data[row][6]

            # Sheet1.cell(row=7 + row, column=9).value = data[row][4]

        # keep only four Sheet
        for s in sheets:
            if s != 'B06':
                sheet_name = wb.get_sheet_by_name(s)
                wb.remove_sheet(sheet_name)
        return wb

    def productPlanPart(self, data=None, hPerDay=0, orderday=QtCore.QDate()):
        for i, j in enumerate(data):
            divine = j[8]
            if not divine.isdigit():
                del data[i]
        for i, j in enumerate(data):

            prio = j[9]
            if not prio.isdigit():
                data[i][9] = 1
            else:
                data[i][9] = int(prio)
        for i, j in enumerate(data):
            number = j[4]
            if not number.isdigit():
                data[i][4] = 0
            else:
                data[i][4] = int(number)
        if not data:
            pass

        items = []
        lines = []
        priority = []
        for each in data:
            hour = int(each[4])/int(each[8])
            day = int(hour/hPerDay) + 1
            lines.append(each[7])
            priority.append(each[9])
            item = Part(each[1], each[7], each[9], day)
            items.append(item)

        new_priority = list(set(priority))
        new_priority.sort()
        new_lines = list(set(lines))
        new_items = []
        for i in range(len(items)):
            for j in range(len(new_lines)):
                new_items.append([])
                if items[i].line == new_lines[j]:
                    new_items[j].append(items[i])
        new_items = list(filter(None, new_items))

        for each in new_items:
            each.sort(key=attrgetter('priority'))
            each[0].startDate = orderday
            each[0].endDate = each[0].startDate.addDays(each[0].day)
            for i in range(1, len(each)):
                each[i].startDate = each[i-1].endDate.addDays(1)
                each[i].endDate = each[i].startDate.addDays(each[i].day)


        final_items = []
        for item in new_items:
            for each in item:
                final_items.append(each)
        for item in final_items:
            print(item)
        wb = load_workbook("form_copy.xlsx")
        sheets = wb.sheetnames
        Sheet1 = wb[sheets[6]]
        # Set value to table
        for i in range(len(data)):
            Sheet1.cell(row=6+i, column=1).value = i+1

        for row in range(len(final_items)):
            Sheet1.cell(row=6 + row, column=2).value = final_items[row].name
            Sheet1.cell(row=6 + row, column=3).value = final_items[row].startDate.toString('MMMM d, yyyy')
            Sheet1.cell(row=6 + row, column=4).value = final_items[row].endDate.toString('MMMM d, yyyy')


        # keep only seven Sheet
        for s in sheets:
            if s != 'B07':
                sheet_name = wb.get_sheet_by_name(s)
                wb.remove_sheet(sheet_name)
        return wb

    def save(self, wb, part):
        wb.save(part)

# def main():
#     new = CreateDocuments()
#     new.closurePlan()
#
# main()