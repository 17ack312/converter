import re,sys
import win32com.client,os

w=win32com.client.Dispatch("Word.application")
w.Visible=0

try:
	file=sys.argv[1]
except:
	file=input("enter pdf file name with absolute path and extension :")

file=re.sub(r'\\',r"\\\\",file)


'''
print("0. editable '.doc' ")
print("1. editable '.dot' ")
print("2/3/4/5/7. editable '.txt' ")
print("6. editable '.rtf' ")
print("8. editable '.htm'  with source folder")
print("9. editable '.mht' ")
print("10. editable '.html' ")
print("11/19/20/21/22. editable '.xml' ")
print("12/16/24. editable '.docx' ")
print("13. editable '.docm' ")
print("14. editable '.dotx' ")
print("15. editable '.dotm' ")
print("17. editable '.pdf' ")
print("18. editable '.xps' ")
print("23. editable '.odt' ")
print("28. open as editable word document")
'''

print("[=] You are using ",file.split('.')[-1],"file")

print("""
	[0] -> '.doc'        [1] -> '.dot'      [2] -> '.txt'        
	[3] -> '.txt'        [4] -> '.txt'      [5] -> '.txt' 
	[6] -> '.rtf'        [7] -> '.txt'   	[8] -> '.html'       
	[9] -> '.mht'        [10]-> '.html'     [11]-> '.xml' 
	[12]-> '.docx'       [13]-> '.docm'     [14]-> '.dotx'
	[15]-> '.dotm'       [16]-> '.docx'     [17]-> '.pdf'
	[18]-> '.xps'        [19]-> '.xml'      [20]-> '.xml'
	[21]-> '.xml'        [22]-> '.xml'      [23]-> '.odt'
	                     [24]-> '.docx'       

	                   [28]. open with Word

""")




#word=os.path.abspath(file+"".format())
i=int(input("choice ="))
wb = w.Documents.Open(file)
word=os.path.abspath(file[0:-4]+"".format())
wb.SaveAs2(word,FileFormat=i)
#print("done")
print("[=] File saved as",word)

wb.Close()
w.Quit()