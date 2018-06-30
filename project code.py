from pyexcel_xls import read_data
from bs4 import BeautifulSoup
import urllib.request as ur
import xlsxwriter
import sqlite3
import re


    # TO CREATE IGNORE TEXT FILE
file=open("E:\\thila\\python\\PROJECT\\ignore.txt","r")
ignoreword=file.read().split()
ignoreset=set(ignoreword)
file.close()
#print("\nSET TYPE FUNCTION:\n",ignoreset)# to avoid repeated words

heading=['TOP LIST','FREQUENCY','DENSITY']

    #TO CREATE DATA BASE
db=sqlite3.connect('data.db')
print("Data Base Created Successfully")
c=db.cursor()
#c.execute('''CREATE TABLE TOPLISTA(KEYWORD TEXT NOT NULL,FREQ INT NOT NULL)''')
#c.execute('''CREATE TABLE TOPCONTENTA(W1 INT NOT NULL,W2 INT NOT NULL,W3 INT NOT NULL,W4 INT NOT NULL,W5 INT NOT NULL)''')
print("Table Created Successfully") 


    #COLLECT THE CONTENT FROM WEB PAGE & CALCULATE TOP 5 WORDS, FREQUENCY & DENSITY 
req=ur.Request("https://www.myklassroom.com/Engineering-branches/28/Industrial-Engineering",data=None,headers={'user_Agent':'Mozilla 5.0(Macintosh;Indel MAc Os X 10-9-3)applewebkit/537.36(KHTML,like Gecko) chrome/35.0.1916.47 safari/537.36'})
f=ur.urlopen(req)
#print(f.read())#.decode('utf-8'))

soup=BeautifulSoup(f,'html.parser')
head=[soup.title.string]

for script in soup(['script','style','[document]','head','title']):
    script.extract()
    text=soup.get_text().lower()
    fill=filter(None,re.split('\W|\d',text))
    d={}
    wordcount=len(text)
for word in fill:
     word=word.lower()
     if word not in ignoreset:
         if word not in d:d[word]=1 
         else:d[word]+=1 

a=sorted(d.items(),key=lambda x:x[1],reverse=True,)[:5]

density=[]
for ke,va in a:
    key=len(ke)
    dens=((key/wordcount*100))
    density.append(dens)

print("Calculations are over")

    #TO INSERT THE DATABASE
steps=[(k,v)for k,v in a]
#c.executemany("INSERT INTO TOPLISTA(KEYWORD,FREQ) VALUES(?,?)",(steps))
#c.execute("INSERT INTO TOPCONTENTA(W1,W2,W3,W4,W5) VALUES(?,?,?,?,?)",(density))
db.commit()
c.execute("SELECT * FROM TOPLISTA")

word_collection=[]
for mat in c.fetchall():
    word_collection.append(mat)
    
top_word=[]
top_frequency=[]
for k1,v1 in word_collection:
    top_word.append(k1)
    top_frequency.append(v1)
    
c.execute("SELECT * FROM TOPCONTENTA")
value=[]
for res in c.fetchall():
    for val in res:
        value.append(val)

    
    #TO APPLY THE RESULT TO XLSX WRITER
workbook=xlsxwriter.Workbook("E:\\thila\\python\\PROJECT\\analysis.xlsx")
worksheet=workbook.add_worksheet()
bold = workbook.add_format({'bold': True})
data=[top_word,top_frequency,value]
worksheet.set_column('A:A',20)
worksheet.set_column('B:B',15)
worksheet.set_column('C:C',20)

worksheet.write_row('A1',heading,bold)
worksheet.write_column('A2',data[0])
worksheet.write_column('B2',data[1])
worksheet.write_column('C2',data[2])
chart=workbook.add_chart({'type':'pie'})

chart.add_series({
        'name':       '=Sheet1!$A$2:$A$6',
        'categories': '=Sheet1!$B$2:$B$6',
        'values':     '=Sheet1!$C$2:$C$6',
      })

worksheet.write('B8',"THANK YOU",bold)

    #DRAW A CHART
chart.set_title ({'name': 'MY CALCULATION'})
worksheet.insert_chart('E2', chart)
workbook.close()

print("Check The Result")
print("Find The Result in Excel Sheet")
























    


















    
