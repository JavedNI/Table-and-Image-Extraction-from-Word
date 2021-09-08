import numpy as np
import os
import docx2txt as d2t
from docx.api import Document
import xlsxwriter
from tkinter import*
from tkinter import filedialog
from tkinter import messagebox
from ttkthemes import ThemedTk
from tkinter import ttk


def window(main):
    main.title("Extraction Program")
    main.update_idletasks()
    width = main.winfo_width()
    height = main.winfo_height()
    x = (main.winfo_screenwidth()//2) - (width //2)
    y = (main.winfo_screenheight()//2) - (height//2)
    main.geometry("{}x{}+{}+{}".format(width,height,x,y))


def extract_images_from_docx(path_to_file,images_folder,get_text=False):
    text = d2t.process(path_to_file,images_folder)
    if(get_text):
        return text

def quit():
    root.destroy()
               
def extract_images():
    beginButton["state"] = DISABLED
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

        def quit():
            root.destroy()

        def clear_label():
            response["text"] = ""

        def age_rename(count = 1,path = os.chdir(images_folder)):
            age_button["state"] = DISABLED
            try:
                x = int(age_prompt.get())
                for filename in os.listdir(path):
                    extensions = [".jpg",".png",".bmp",".gif",".tiff",".psd",".pdf",".eps",".ai",".indd",".raw"]
                    if filename.endswith(tuple(extensions)):
                        new_name = ("IMG " + str(x) + "-{}.jpg").format(str(count).zfill(3))
                        try:
                            os.rename(filename,new_name)
                            count = count + 1
                            response.config(text = "Done!")
                            response.after(2000,clear_label)
                        except FileExistsError as error:
                            if filename != new_name:
                                os.remove(filename)
                                print("Identical Images which is: " + str(filename))

            except ValueError:
                response.config(text = "Please enter an integer age value")
                age_window.after(2000,clear_label)
                age_button["state"] = ACTIVE
                    
        
        if text_input == 1:
            decision = text_input
            data = extract_images_from_docx(path_to_file,images_folder,get_text =True)
            age_window = Toplevel()
            window(age_window)
            age_window.title("Age of Dataset")
            age_label = ttk.Label(age_window, text = "Enter the age of the Dataset").grid(row = 2, column = 2)
            age_prompt = ttk.Entry(age_window, width = 20)
            age_prompt.grid(row = 0, column = 50) 
            age_button = ttk.Button(age_window, text = "Proceed",command = age_rename)
            age_button.grid(row = 0, column = 40,pady = 5)
            quit_button = ttk.Button(age_window, text = "Close", command = quit)
            quit_button.grid(row = 2,column = 40)
            response = ttk.Label(age_window, text = "")
            response.grid(row = 1, column = 40)

            text_window = Toplevel()
            window(age_window)
            text_window.title("Extracted Text")

            # Create A Main Frame

            main_frame = Frame(text_window)
            main_frame.pack(fill = BOTH, expand = 1)
            

            # Create a Canvas

            my_canvas = Canvas(main_frame)
            my_canvas.pack(side = LEFT, fill = BOTH, expand = 1)

            # Add A Scrollbar To The Canvas

            my_scrollbar = ttk.Scrollbar(main_frame, orient = VERTICAL, command = my_canvas.yview)
            my_scrollbar.pack(side = RIGHT, fill = Y)

            # Configure The Canvas

            my_canvas.configure(yscrollcommand = my_scrollbar.set)
            my_canvas.bind("<Configure>", lambda e: my_canvas.configure(scrollregion = my_canvas.bbox("all")))

            # Create ANOTHER Frame INSIDE the Canvas
            second_frame = Frame(my_canvas)

            # Add that New Frame To a Window In The Canvas
            my_canvas.create_window((0,0),window = second_frame, anchor = "nw")

            data_output = Label(second_frame, text = data)
            data_output.grid(row = 0, column = 0)
            

                                                           
        else:
            decision = text_input
            data = extract_images_from_docx(path_to_file,images_folder,get_text =False)
            age_window = Toplevel()
            window(age_window)
            age_window.title("Age of Dataset")
            #age_label = Label(age_window, text = "Enter the age of the Dataset").grid(row = 2, column = 2)
            age_prompt = ttk.Entry(age_window, width = 20)
            age_prompt.grid(row = 0, column = 50) 
            age_button = ttk.Button(age_window, text = "Proceed",command = age_rename)
            age_button.grid(row = 0, column = 40,pady = 5)
            response = ttk.Label(age_window, text = "")
            response.grid(row = 1, column = 40)
            quit_button = ttk.Button(age_window, text = "Quit", command = quit)
            quit_button.grid(row = 1, column = 40)

root = ThemedTk(theme = "black")
window(root)            

root.wm_title("Image Extractor!")
beginButton = ttk.Button(root,text = "Begin!", width = 6, command = extract_images)
beginButton.grid(row = 0,column = 1)

#### FOR TOMORROW, CHANGE UP THE SIZES OF THE WIDGETS AND TEXTBOXES
###Improve on the widgets, small functions such as checking for images in file or saving function
### Showing message on a label isntead of terminal if there exists duplicate copies

quit_button = ttk.Button(root, text = "Quit", command = quit)
quit_button.grid(row = 0,column = 2)            
root.mainloop()





