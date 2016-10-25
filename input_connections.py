import xlwt


f = open('file1.txt')
f1 = f.readlines()
f.close()

f = open('file2.txt')
f2 = f.readlines()
f.close()

#f1 = ''.join(f1)
#f2 = ''.join(f2)

font0 = xlwt.Font()
font0.name = 'Calibri'
font0.colour_index = 2
font0.bold = True

style0 = xlwt.XFStyle()
style0.font = font0

wb = xlwt.Workbook()
ws = wb.add_sheet('KMS_Report')

ws.write(0,0, 'KMS_LOGS(FQDN)')
ws.write(0,1, 'INPUT_CONNECTIONS(1688/TCP)')

for ind in range(0, len(f1)):
    ws.write(ind + 1, 0, f1[ind])
for ind in range(0, len(f2)):
    ws.write(ind + 1, 1, f2[ind])

wb.save('example.xls')