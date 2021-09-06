import cv2
import numpy as np
import shutil
import os
import docx2txt as d2t
from docx.api import Document
import pandas as pd
import xlsxwriter
import sys
from tkinter import*
from tkinter import filedialog
from tkinter import messagebox


def extract_images_from_docx(path_to_file,images_folder,get_text=False):
    text = d2t.process(path_to_file,images_folder)
    if(get_text):
        return text

def browse_file():
    broButton["state"] = DISABLED
    docx = filedialog.askopenfilename(filetypes = [("Docx files","*.docx")])
    print(docx)
    while os.path.isfile(docx) == False:
        messagebox.showerror("Uh oh","This document does not exists")
        docx = filedialog.askopenfilename(filetypes = [("Docx files","*.docx")])
    else:
        path_to_file = str(docx)
        messagebox.showinfo("Where to save your content","Find a location where you want to save your images")
        folder_location = filedialog.askdirectory()
        images_folder = str(folder_location)
        text_input = messagebox.askyesno("Extract Texts","Do you want to extract texts too?")
        if text_input == 1:
            data = extract_images_from_docx(path_to_file,images_folder,get_text = True)
            count = 1
            path = os.chdir(images_folder)
            age_window = Toplevel()
            age_window.title("Age")
            age_label = Label(age_window).pack()
            age_prompt = age_window.Entry(age_window) ##continue from here
            age_prompt.grid()
            new.focus()
            age_prompt.grid()

        else:
            print("okay")

            
            
        
        
        
    
        

root = Tk()
root.wm_title("Image Extractor!")
broButton = Button(root,text = "Begin!", width = 6, command = browse_file)
broButton.grid(row = 0,column = 1)



root.mainloop()


#-----------------------------------Image Extractor-----------------------------------------------------------------------------#


#Creating the function by using docx2txt library to extract images to folder and have the ability to just copy text

user_input = input("Enter the path of your word docx file w/ extension: ")
while os.path.isfile(user_input)== False:
    print("This document does not exist")
    user_input = input("Enter the path of your word docx file w/ extension: ")
else:
    path_to_file = str(user_input)
    while True:
        folder_location = input("Where would you like to place content? ")
        while os.path.exists(folder_location) == False:
            print("Please place file in an existing folder")
            folder_location = input("Where would you like to place content? ")
        else:
            images_folder = str(folder_location)
            break
     
        
#User prompted to specify the file location and the filename and extension and the location of where the user wants the image
#located. Double checks that the user entered a word doc otherwise will ask the user to enter the correct file. 
   
text_input = input("Do you want to extract texts too? ")
#Giving the user the option to extract texts
while text_input not in ("Y","N","Yes","No"):
    print("Either Yes/No or Y/N")
    text_input = input("Do you want to extract texts too? ")
if text_input == "Y":
    #If the user says yes, then the extraction of the images and text begins, specifying iterations and parameters before renaming the image file
    #according to age
    data = extract_images_from_docx(path_to_file,images_folder,get_text =True)
    count = 1
    path = os.chdir(images_folder)
    age_prompt = input("Enter the Age of the Dataset: ")
    while age_prompt.isdigit() == False:
        print("Please enter an integer Age value")
        age_prompt = input("Enter the Dataset number: ")
    else:
        for filename in os.listdir(path):
            extension = [".jpg",".png",".bmp",".gif",".tiff",".psd",".pdf",".eps",".ai",".indd",".raw"]
            if filename.endswith(tuple(extension)):
                new_name = ("IMG " + str(age_prompt) + "-{}.jpg").format(str(count).zfill(3))
                try:
                    os.rename(filename,new_name)
                    count = count + 1
                except FileExistsError as error:
                    print("Images already exists with that name")
                    if filename != new_name:
                        os.remove(filename)
            else:
                print("No images detected")
                
        print("Done!")
                    
        
    #User will be prompted to enter the age value so that the image rename will include the ages to prevent any images from being copied if you wish to store the groups of images
    #in the same file location and ensures that the value inputted is an integer only otherwise will continue to ask the user to input the age value

    print(data)    
else:
    data = extract_images_from_docx(path_to_file,images_folder,get_text =False)
    count = 1
    path = os.chdir(images_folder)
    age_prompt = input("Enter the Age of the Dataset: ")
    while age_prompt.isdigit() == False:
        print("Please enter an integer Age value")
        age_prompt = input("Enter the data set number: ")
    else:
        for filename in os.listdir(path):
            extension = [".jpg",".png",".bmp",".gif",".tiff",".psd",".pdf",".eps",".ai",".indd",".raw"]
            if filename.endswith(tuple(extension)):
                new_name = ("IMG " + str(age_prompt) + "-{}.jpg").format(str(count).zfill(3))
                try:
                    os.rename(filename,new_name)
                    count = count + 1
                except FileExistsError as error:
                    print("Images already exists with that name")
                    if filename != new_name:
                        os.remove(filename)
            else:
                print("No images detected")

        print("Done!")


        
#same process happens here when the user chooses to not include texts
        
#----------------------------------------------------------------------------------------------------------------#


