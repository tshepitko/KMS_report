import subprocess, smtplib, datetime, xlwt
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

cmd ="wevtutil qe \"Key Management Service\" /rd:true /f:text"

output = subprocess.check_output(cmd, stderr=subprocess.STDOUT)
output = output.split('\n')
res = []
today = datetime.date.today().strftime("%Y-%m")
today_1m = (datetime.date.today() - datetime.timedelta(365/12)).strftime("%Y-%m")

for line in output:
    if line.find('Date:',0,len(line)) != -1:
        if line.find(today,0,len(line)) == -1 and line.find(today_1m,0,len(line)) == -1:
            break
    elif line[:2] == '0x':
        ind1 = line.find(',',0,len(line))
        ind2 = line.find(',',ind1+1,len(line))
        ind3 = line.find(',',ind2+1,len(line))
        res.append(line[ind2+1:ind3])
res = list(set(res))
res.sort()

f = open('C:\\KMS_log\\logs.txt', 'r')
addresses = f.readlines()
f.close()
for ind in range(0, len(addresses)):
    ind1 = addresses[ind].find('       ',0,len(addresses[ind]))
    ind2 = addresses[ind].rfind(':',0,len(addresses[ind]))
    addresses[ind] = addresses[ind][ind1+7:ind2]
addresses = list(set(addresses))
for ind in range(0,len(addresses)):
    if addresses[ind].find('c2-',0,len(addresses[ind])) != -1:
        addresses[ind] = addresses[ind][3:]
        addresses[ind] = addresses[ind].replace('-','.',3)
addresses.sort()        

wb = xlwt.Workbook()
ws = wb.add_sheet('KMS Report')
ws.col(0).width = 256 * 35
ws.col(1).width = 256 * 27

style0 = xlwt.easyxf('font: bold on')

ws.write(0, 0, "KMS LOGS(FQDN)", style0)
ws.write(0, 1, "AUTHORIZATION REQUESTS", style0)

for ind in range(0, len(res)):
    ws.write(ind + 1, 0, res[ind])
for ind in range(0, len(addresses)):
    ws.write(ind + 1, 1, addresses[ind])

wb.save('report.xls')


f = open('report.xls', 'rb')
att = MIMEBase("application", "vnd.ms-excel")
att.set_payload(f.read())
encoders.encode_base64(att)


me = "KMS Server"
you = ["example.com"]

msg = MIMEMultipart()
msg['Subject'] = "KMS Logs"
msg['To'] = "example@example.com" 
msg.add_header('Content-Disposition', 'attachment; filename="report.xls"')
msg.attach(att)

s = smtplib.SMTP('localhost')
s.sendmail(me, you, msg.as_string())
s.quit()
