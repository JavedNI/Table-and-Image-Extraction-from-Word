# import relevant modules 
from tkinter import *
from PIL import ImageTk, Image
from tkinter import filedialog as fd
from tkinter import messagebox
import pandas as pd
import xlsxwriter
import docx2txt as d2t
from docx.api import Document



root = Tk()
root.title('file finder')


# function to extract tables from a word docx and write them to an .xlsx or .csv file 
def extract_tables_from_docx(file_path,excel_path): 
    
    # file path for the document being read
    document = Document(file_path)  
    data = []
    
    # iterates through the tables in the word doc by collating both variables together
    for table in document.tables:  
        # column data stored in keys and values, then stores data in a list 
        keys = (cell.text for cell in table.columns[0].cells) 
        values = (cell.text for cell in table.columns[1].cells)
        data.append(dict(zip(keys,values))) 
        
    # the dataframe is converted into a .csv or .xlsx file                                       
    df1 = pd.DataFrame(data) 
    if ('.xlsx') in excel_path: 
        df1.to_excel(excel_path, engine = 'xlsxwriter', index = False)
    else:
        df1.to_csv(excel_path ,index=False)
    
    
    
# function to find the document path for the extract_tables_from_docx function
def find_document_path(): 
    
    # file_path stores the file location as a global string to be used by the extract_tables_from_docx function
    global file_path
    file_path = fd.askopenfilename(initialdir="/", title='File search', filetypes=(('docx files', '*.docx'), ('all files', '*.*')))
    
    # checks if file_path conatins a docx file 
    if ('.docx') in file_path: 
        # checks if the user wants to use the selected file path 
        check = messagebox.askyesno('File Path', 'Use this file path? ' + file_path)
        if check == 1:
            # confirms the file path being used, then changes the state of the buttons on the root label 
            path_info = messagebox.showinfo('File path found', 'Using: '+ file_path)
            open_btn['state'] = 'disabled'
            save_btn['state'] = 'normal'
        else:
            return
    else:
        doc_err = messagebox.showerror('Docx error', 'File path does not contain a .docx file')
    
    
# function to save a .xlsx or .csv file 
def save_file():
    
    # excel_path stores the string used to save the excel file 
    excel_path = fd.asksaveasfilename(defaultextension=".*", initialdir="/", title="Save File", filetypes=(('xlsx files', '*.xlsx'), ('cvs files','.csv'))) #
    # checks to see if a path was selected 
    if excel_path:
        extract_tables_from_docx(file_path, excel_path)
        # confirms the excel file was saved and closes the program 
        saved = messagebox.showinfo('File saved', 'File saved')
        root.after(0, root.destroy)

        
# creates and puts buttons on the root label, when the buttons are pressed they call functions 
open_btn = Button(root, text='open file', command=find_document_path)
save_btn = Button(root, text='save file', state  = DISABLED, command=save_file) 
open_btn.pack()
save_btn.pack()
root.mainloop()



        
        
        
            

        
       
    
    
    

