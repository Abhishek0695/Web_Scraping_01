import requests
import csv
import os
import smtplib
from xlsxwriter import Workbook
from bs4 import BeautifulSoup
from email.message import EmailMessage

email_add = os.environ.get("EMAIL_ADD")
email_password = os.environ.get("EMAIL_PASS")

request = requests.get("https://www.cricbuzz.com/cricket-stats/icc-rankings/men/batting").text
soup = BeautifulSoup(request,'lxml')
all_ranks = soup.find_all('div', class_='cb-col cb-col-100 cb-padding-left0')
dict1={'Test' : 0, 'Odi' : 1, 'T20' : 2}
dict1['Test'] = all_ranks[0]
dict1['Odi'] = all_ranks[1]
dict1['T20'] = all_ranks[2]
a=dict1['Test']
b=dict1['Odi']
c=dict1['T20']

rank_test= a.find_all('div', class_= 'cb-col cb-col-16 cb-rank-tbl cb-font-16')
player_name_test= a.find_all('div', class_='cb-col cb-col-67 cb-rank-plyr')
rating_test= a.find_all('div', class_= 'cb-col cb-col-17 cb-rank-tbl pull-right')

rank_odi= b.find_all('div', class_= 'cb-col cb-col-16 cb-rank-tbl cb-font-16')
player_name_odi= b.find_all('div', class_='cb-col cb-col-67 cb-rank-plyr')
rating_odi= b.find_all('div', class_= 'cb-col cb-col-17 cb-rank-tbl pull-right')

rank_t20= c.find_all('div', class_= 'cb-col cb-col-16 cb-rank-tbl cb-font-16')
player_name_t20= c.find_all('div', class_='cb-col cb-col-67 cb-rank-plyr')
rating_t20= c.find_all('div', class_= 'cb-col cb-col-17 cb-rank-tbl pull-right')

csv_file= open('Cricbuzz_csv.csv', 'w')
csv_writer= csv.writer(csv_file, lineterminator='\n')
csv_writer.writerow(['Rank','Player_name','Country','Rating'])

csv_file.write('Test Rankings \n')

wb = Workbook('Cricbuzz.xlsx')
ws1= wb.add_worksheet('Test')
ws2= wb.add_worksheet('Odi')
ws3= wb.add_worksheet('T20')

ws1.write(0,0,'Rank')
ws1.write(0,1,'PlayerName')
ws1.write(0,2,'Country')
ws1.write(0,3,'Rating')
row = 1
column = 0
for i,j,k in list(zip(rank_test,player_name_test,rating_test))[:10]:
    j1=j.a
    j2=j.find('div', class_='cb-font-12 text-gray')
    #t1.append((i.text,j1.text,j2.text,k.text))
    csv_writer.writerow([i.text,j1.text,j2.text,k.text])
    ws1.write(row,column,i.text)
    ws1.write(row,column+1,j1.text)
    ws1.write(row, column + 2, j2.text)
    ws1.write(row, column + 3, k.text)
    row += 1

csv_file.write('Odi Rankings \n')

ws2.write(0,0,'Rank')
ws2.write(0,1,'PlayerName')
ws2.write(0,2,'Country')
ws2.write(0,3,'Rating')
row = 1
column = 0

for i,j,k in list(zip(rank_odi,player_name_odi,rating_odi))[:10]:
    j1=j.a
    j2=j.find('div', class_='cb-font-12 text-gray')
    #t1.append((i.text,j1.text,j2.text,k.text))
    csv_writer.writerow([i.text,j1.text,j2.text,k.text])
    ws2.write(row,column,i.text)
    ws2.write(row,column+1,j1.text)
    ws2.write(row, column + 2, j2.text)
    ws2.write(row, column + 3, k.text)
    row += 1

csv_file.write('T20 Rankings \n')

ws3.write(0,0,'Rank')
ws3.write(0,1,'PlayerName')
ws3.write(0,2,'Country')
ws3.write(0,3,'Rating')
row = 1
column = 0

for i,j,k in list(zip(rank_t20,player_name_t20,rating_t20))[:10]:
    j1=j.a
    j2=j.find('div', class_='cb-font-12 text-gray')
    #t1.append((i.text,j1.text,j2.text,k.text))
    csv_writer.writerow([i.text,j1.text,j2.text,k.text])
    ws3.write(row, column, i.text)
    ws3.write(row, column + 1, j1.text)
    ws3.write(row, column + 2, j2.text)
    ws3.write(row, column + 3, k.text)
    row += 1

wb.close()

csv_file.close()

msg= EmailMessage()
msg['From']= email_add
msg['TO']= email_add, 'abhishek.pancholi@lntinfotech.com'
msg['Subject']= 'Cricbuzz Rankings'
msg.set_content('Updated  Cricket Rankings - Men. Kindly find the attachment')

with open('Cricbuzz.xlsx', 'rb') as f:
    file_data=f.read()
    file_name=f.name

msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
    smtp.login(email_add,email_password)
    smtp.send_message(msg)




