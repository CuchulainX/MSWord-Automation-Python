#Basic includes to connect with Microsoft Word and Creating Small GUI
import win32com.client as win32
import os
def CountTableColumn12(Name):
    #Open the document
    words = win32.gencache.EnsureDispatch('Word.Application')
    path_to_script = os.path.dirname(os.path.abspath(__file__))
    f = open(path_to_script + '\Report.txt','a')
    doc = words.Documents.Open(Name)
    words.Visible = True
    #Gives the all table in document
    TableCount = doc.Tables.Count
    Count = 0
    #logic
    f.write('\n' + "************************* Find Table with Column 12 **********************************" + '\n') 
    for i in range(1,TableCount):
        table = doc.Tables(i)
        Col = table.Columns.Count        
        if Col == 12:
            Count = Count + 1
    f.write(str(Count))

