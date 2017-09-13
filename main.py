import urllib.request  as urllib2 #import the library used to query a website
from bs4 import BeautifulSoup #import the Beautiful soup functions to parse the data returned from the website
import pandas as pd #import pandas to convert list to data frame

def loadExcelDoc(sheet_index):
    # Open up Faton Excel file
    filename = "Verisk Model_Send_Excel_2.xlsx"
    path = "~/Documents/Git/Html_scraping_project/"
    xl = pd.ExcelFile(path + filename)
    sheets = xl.sheet_names


    ## open up the first sheet and print our data
    df = xl.parse(sheets[sheet_index],header=None)
    row_index = df.iloc[0][0]

    df = xl.parse(sheets[sheet_index])
    df = df.set_index(row_index)

    return df

def main():
	#specify the url
	url = "http://www.verisk.com/press-releases/2017/february/verisk-analytics-inc-reports-fourth-quarter-2016-financial-results.html"

	user_agent = 'Mozilla/5.0 (Windows; U; Windows NT 5.1; en-US; rv:1.9.0.7) Gecko/2009021910 Firefox/3.0.7'
	headers={'User-Agent':user_agent,} 

	request=urllib2.Request(url,None,headers) #The assembled request
	page = urllib2.urlopen(request)
	data = page.read() # The data u need


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

    df = loadExcelDoc(0)



