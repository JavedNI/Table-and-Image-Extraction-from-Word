import cv2
import numpy as np
import shutil
import matplotlib.pyplot as plt
import os
from docx.api import Document
import pandas as pd
from re import search
from functools import reduce
import xlsxwriter

print("TABLE EXTRACTOR")

def extract_tables_from_docx(path,output_path,xlsx_name):
    document = Document(path)
                                    #opens the document
    data = []
                                    #creates an empty list datatype
    for table in document.tables:
        keys = (cell.text for cell in table.columns[0].cells)
        values = (cell.text for cell in table.columns[1].cells)
        data.append(dict(zip(keys,values)))
        
    """
    Iterating every table in that word doc and creating a dictionary variable by collating the two variable together
    Before that each column data is being retrived and stored in the keys and values variable that are collated as dictionaries
    which is then converted into a dataframe only which makes it easier to save to the relevant excel file and for further data
    manipulation

    """

    df1 = pd.DataFrame(data)
    df1.to_csv(str(output_path)+"/"+str(xlsx_name)+" cs_version.csv",index=False)
    df1.to_excel(str(output_path)+"/"+str(xlsx_name)+".xlsx",engine = "xlsxwriter",index=False)
    print(df1)
                 

#create a function for tranposing the data
#Creating a function in which the word document is extracted and defined from the second line of code
#Using pandas to create an empty excel canvas that ensures that the file name can be the users choice
#for loop being set up and having dummy parameters setup for the table from word to be extracted and be
#laid out in different sheets
#Using pandas to append all the row data together from a dict datatype and creating a dataframe
#converting the dataframe to an excel file with x amounts of sheets which is reflected by the amounts of table
#existing



user_input = input("Enter the path of you word docx file w/extension: ")
while (".docx") not in user_input:
    print("Please enter a file directory that includes the .docx extension") 
    user_input = input("Enter the path of you word docx file w/extension: ")
else:
    path = str(user_input)
    while os.path.isfile(user_input) == False:
        print("This document does not exist")
        user_input = input("Enter the path of your word docx file w/extension: ")
    else:
        while True:
            folder_location = input("Where would you like to place content? ")
            while os.path.exists(folder_location) == False:
                print("Please place file in an existing folder")
                folder_location = input("Where would you like to place content")
            else:
                output_path = str(folder_location)
                break

        os.chdir(output_path)
        xlsx_name = input("Enter name of xlsx file: ")
        xlsx_name = str(xlsx_name)
        while os.path.isfile(xlsx_name + ".xlsx") == True:
            print("File already exists with that name, please enter a different name")
            xlsx_name = input("Enter name of xlsx file: ")

        else:
            extract_tables_from_docx(path,output_path,xlsx_name)
            print("XLSX and CSV files created!")


image_permission = input("Image Extraction, Y/N? ")
if (image_permission == "Y"):
    import IE
else:
    print("done")

                   
 #option for image extraction

                


            
    
#Prompts user to input the directory to the file, using a quick check the user specifically entered a .docx file
#then double checking with os.path.isfile to ensure that it's actually a file otherwise asking the user again for
#for the docx file
#Then the same check happens for where the user wants to save his files, a check if the file directory actually exists
#Once that is confirmed, the working directory is changed to ensure it's working at the right directory since directory changes
#User inputs name of xlsx file
#Once all the condition is met, function above is called
    
#Since I didn't check the add path box when installing python, it's limited as to where I can save my files to
#However, when that is fixed it should work



        
        
        
            

        
       
    
    
    

