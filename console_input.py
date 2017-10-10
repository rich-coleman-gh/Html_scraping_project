#import the library used to query a website
import urllib.request  as urllib2

# INPUT VARIABLES SPECIFIED BY THE USER
Excel_column_index = 8
Table_from_web_index = 8
Table_from_web_column_index = 2
Excel_sheet_index = 0
filename = "Verisk Model_Send_Excel_2.xlsx"
path = "~/Documents/Git/Html_scraping_project/"

#specify the url
#url = "http://www.verisk.com/press-releases/2016/november/verisk-third-quarter-2016-results.html"
url = "http://www.verisk.com/press-releases/2017/february/verisk-analytics-inc-reports-fourth-quarter-2016-financial-results.html"

user_agent = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7'
headers={'User-Agent':user_agent,}

request=urllib2.Request(url,None,headers) #The assembled request
page = urllib2.urlopen(request)
data = page.read() # The data u need

#import the Beautiful soup functions to parse the data returned from the website
from bs4 import BeautifulSoup

#Parse the html in the 'page' variable, and store it in Beautiful Soup format
soup = BeautifulSoup(data,'html.parser')

#####gather all tables from page into array 
all_tables = soup.find_all("table", class_="table-blue")

parsed_tables = []
for i in range(len(all_tables)-1):
    table_body = all_tables[i].find('tbody')

    rows = table_body.find_all('tr')

    df_temp = []
    for row in rows:
        cols =row.find_all('td')
        cols = [ele.text.strip() for ele in cols]
        df_temp.append([ele for ele in cols]) #get rid of empty values
    parsed_tables.append(df_temp)


import pandas as pd #import pandas to convert list to data frame

df_from_web = pd.DataFrame(parsed_tables[Table_from_web_index])


def loadExcelDoc(sheet_index):
    # Open up Faton Excel file
    xl = pd.ExcelFile(path + filename)
    sheets = xl.sheet_names


    ## open up the first sheet and print our data
    df = xl.parse(sheets[sheet_index],header=None)
    #row_index = df.iloc[0][0]

    #df = xl.parse(sheets[sheet_index])
    #df = df.set_index(row_index)

    return df



df = loadExcelDoc(Excel_sheet_index)


# Lets try to match the row index names from the web to the excel doc
excel_labels = df.loc[:,0] # we assume that the labels are always in the first column of both dataframes
web_labels = df_from_web.loc[:,0]

for i,excel_label in enumerate(excel_labels):
    for j,web_label in enumerate(web_labels):
        if excel_label == web_label:
            #modify the Excel_column_index to equal the Table_from_web_column_index value
            #print("excel label: ",i,excel_label)
            #print("web label: ",j,web_label)
            #print("web table value: ", df_from_web.loc[j,Table_from_web_column_index])
            df.loc[i,Excel_column_index] = df_from_web.loc[j,Table_from_web_column_index]
          

          