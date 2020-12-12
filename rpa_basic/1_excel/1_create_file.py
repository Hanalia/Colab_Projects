# from openpyxl import Workbook
# wb = Workbook() # 새 워크북 생성
# ws = wb.active
# ws.title = "NadoShee"
# ws.sheet_properties.tabColor = "ff66ff"
# for x in range (1,11):
#     c= ws.cell(row = x, column = 1, value = 4)
#     print (c.value)
# print(ws.max_row)

# wb.save("sample.xlsx")
# wb.close()

import glob
filelocation = input("위치를적어주세요")
filetype = input("파일유형 적어주세요")
myList = glob.glob(filelocation + "\*." + filetype)
if not myList:
    fileresult = input("잘못되었습니다")
else:
    print (*myList, sep = "\n")
    fileresult = input("여기있습니다") 

# print (os.path.abspath("17194.xls"))

# 어떤 폴더에 어떤파일들이 있는지 알려주는 프로그램

#C:\Users\Dave\Documents\정석윤\9. 매크로 프로젝트\gunsan\*.xls

#my testbed for codingxls
