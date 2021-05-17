import requests
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

def find_ranks(match_type):
    rank= match_type.find_all('div', class_= 'cb-col cb-col-16 cb-rank-tbl cb-font-16')
    player_name= match_type.find_all('div', class_='cb-col cb-col-67 cb-rank-plyr')
    rating= match_type.find_all('div', class_= 'cb-col cb-col-17 cb-rank-tbl pull-right')
    return (rank,player_name,rating)

def write_first_row(worksheet):
    worksheet.write(0,0,'Rank')
    worksheet.write(0,1,'PlayerName')
    worksheet.write(0,2,'Country')
    worksheet.write(0,3,'Rating')

def write_rows_excel(workseet,rank,player_name,rating):
    row = 1
    column = 0
    for i,j,k in list(zip(rank,player_name,rating)):
        j1=j.a
        j2=j.find('div', class_='cb-font-12 text-gray')
        workseet.write(row,column,i.text)
        workseet.write(row,column+1,j1.text)
        workseet.write(row, column + 2, j2.text)
        workseet.write(row, column + 3, k.text)
        row += 1

rank_test,player_name_test,rating_test= find_ranks(a)
rank_odi,player_name_odi,rating_odi= find_ranks(b)
rank_t20,player_name_t20,rating_t20= find_ranks(c)

wb = Workbook('Cricbuzz1.xlsx')
ws1= wb.add_worksheet('Test')
ws2= wb.add_worksheet('Odi')
ws3= wb.add_worksheet('T20')

write_first_row(ws1)
write_first_row(ws2)
write_first_row(ws3)

write_rows_excel(ws1,rank_test,player_name_test,rating_test)
write_rows_excel(ws2,rank_odi,player_name_odi,rating_odi)
write_rows_excel(ws3,rank_t20,player_name_t20,rating_t20)

wb.close()

# msg= EmailMessage()
# msg['From']= email_add
# msg['TO']= email_add, 'abhishek.pancholi@lntinfotech.com'
# msg['Subject']= 'Cricbuzz Rankings'
# msg.set_content('Updated  Cricket Rankings - Men. Kindly find the attachment')

# with open('Cricbuzz.xlsx', 'rb') as f:
#     file_data=f.read()
#     file_name=f.name

# msg.add_attachment(file_data, maintype='application', subtype='octet-stream', filename=file_name)

# with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
#     smtp.login(email_add,email_password)
#     smtp.send_message(msg)
