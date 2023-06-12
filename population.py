import os,openpyxl,xlrd
cwd = os.getcwd()
files = [file.split(".")[0] for file in os.listdir(cwd +'//population')]
files.sort()
files = [str(file+".xls") for file in files]

population = []
# load files
for file in files:
    wb = xlrd.open_workbook(cwd +'//population' + "//"+file)
    for j in range(1,13):
        sheet = wb.sheet_by_name(str(j))
        population.append(sheet.cell(11,3).value)

# save data

if "population.xlsx" not in os.listdir(cwd):
    wb = openpyxl.Workbook()
    ws = wb.active
    i = 0
    for val in population:
        i +=1
        ws.cell(i,1,value=int(val))
    wb.save("population.xlsx")
else:
    print('Already exsists!')

print('Successfully done!')
