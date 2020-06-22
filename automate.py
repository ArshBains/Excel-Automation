import openpyxl
from openpyxl.styles import colors
from openpyxl.styles import Font, Color, Alignment

# set path to the location of the excel file
path = 'd:/progress/data99.xlsx'
wb = openpyxl.load_workbook(path)

# select sheet with name record
sheet = wb['Record']

# get max number of rows in that excel sheet
number_of_students = sheet.max_row
# print(number_of_students)
# print(type(number_of_students))

# set colors and alignment of data to be added
fti = Font(color=colors.GREEN)
ftd = Font(color=colors.RED)
ftn = Font(color=colors.DARKYELLOW)
ftna = Font(color=colors.BLUE)
alignment = Alignment(horizontal='center')

# iterate through all rows except the first one
for i in range(2, number_of_students+1):
    name = sheet['B'+str(i)]
    progress = sheet['D'+str(i)]
    previous = sheet['I'+str(i)]

    # print(name.value)
    # print(progress.value)
    # print(type(previous.value))

    name_val = name.value
    progress_val = str(progress.value)
    message = ""
    if progress.value > 2:
        message = "Hello, "+name_val+". Great job so far on "+progress_val+"%. Keep it up! If you need anything I am here to assist."
    elif progress.value <= 2:
        message = "Hello, "+name_val+". I see you're only at "+progress_val+"%. How can I help you?"

    if type(previous.value) is int or type(previous.value) is float:
        if previous.value < progress.value:
            sheet['F'+str(i)].value = "Increased"
            sheet['F' + str(i)].font = fti
            sheet['F' + str(i)].alignment = alignment
        if previous.value > progress.value:
            sheet['F' + str(i)].value = "Decreased"
            sheet['F' + str(i)].font = ftd
            sheet['F' + str(i)].alignment = alignment
        if previous.value is progress.value:
            sheet['F'+str(i)].value = "No Change"
            sheet['F' + str(i)].font = ftn
            sheet['F' + str(i)].alignment = alignment
    else:
        sheet['F'+str(i)].value = "NA"
        sheet['F' + str(i)].font = ftna
        sheet['F' + str(i)].alignment = alignment

    # add response
    sheet['E'+str(i)].value = message
    sheet['E' + str(i)].alignment = alignment

    # store the results in meta for later use
    sheet['I'+str(i)].value = progress.value
    sheet['I' + str(i)].font = ftd
    sheet['I' + str(i)].alignment = alignment

# save the file and print 'DONE'
wb.save(path)
print("DONE")