

# import requests 
# for i in range(17194,17208):
#     try:        
#         file_url = "https://gunsan.mof.go.kr/ko/calendar/calendarDownloadFile.do?calKey=" + str(i)
#         r = requests.get(file_url)
#         rawlocation = "C:/Users/Dave/Documents/정석윤/9_macro_project/gunsan2/"
#         finallocation = rawlocation+str(i)+'.xls'
#         # print (finallocation)
#         with open(finallocation,"wb") as f: 
#             f.write(r.content)
#     except:
#         pass


import pandas as pd
import glob
import numpy as np
import xlrd


all_data = pd.DataFrame()
path = 'C:/Users/Dave/Documents/정석윤/9_macro_project/gunsan2/*.xls'
for f in glob.glob(path):
    df = pd.read_excel(f,skiprows=2)
    all_data = all_data.append(df,ignore_index=True)

woodpellet = all_data[all_data['화 물']=='우드펠렛']
woodpellet.to_csv('final5.csv')
