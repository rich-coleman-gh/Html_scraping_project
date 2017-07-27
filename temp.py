# coding: utf-8
rows
for row in rows:
    print(row)
    
for row in rows:
    print("$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$
    )
    
    print(row)
    
for row in rows:
    print("#########################")
    print(row)
    
temp = []
for row in rows:
    cols =row.find_all('td')
    cols = [ele.text.strip() for ele in cols]
    temp.append([ele for ele in cols if ele]) #get rid of empty values
    
temp
