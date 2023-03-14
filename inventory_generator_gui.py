import tkinter as tk
import tkinter.font as font
import fitz 
from tkinter import messagebox
import easygui
import tkinter.font as font
import os
from os import DirEntry, curdir, getcwd, chdir, rename
from glob import glob as glob
import win32com.client as win32

root = tk.Tk();
root.title("ROE-Inventory Records")

def create_files():
    destination_folder=easygui.diropenbox(msg="Please Select The Location Where You Want to Save")
    if destination_folder is None:
        messagebox.showinfo(title="ERROR",message="Select the location where you want to save the records!")
        return(root)
       # exit()
    else:
        destination_folder
        working_directory=os.getcwd()

    source_name=easygui.fileopenbox(msg="Please Select Your XLSX Database:",filetypes=["*.xlsx"])
    if source_name is None:
        messagebox.showinfo(title="ERROR",message="Select Your XLSX Database!")
        return(root)
       # source_name
    else:
        source_name

    #source_name=filedialog.askopenfile(title="Select your XLSX Database:",filetypes=[("Excel files", "*.xlsx")])

    wordApp=win32.Dispatch('Word.Application')
    wordApp.Visible=False

    sourceDoc=wordApp.Documents.Open(easygui.fileopenbox(msg="Please Select Your Word Template",filetypes=["*.docx"]))
    if sourceDoc is None:
        messagebox.showinfo(title="ERROR",message="Selec your word template!")
        return(root)
    else:
        sourceDoc

    mail_merge=sourceDoc.MailMerge
    mail_merge.OpenDataSource(
        Name:=os.path.join(working_directory,source_name),
        sqlstatement:="SELECT * FROM [extractGLPI$]")
    record_count= mail_merge.DataSource.RecordCount

    for i in range(1,record_count + 1):
        mail_merge.DataSource.ActiveRecord = i
        mail_merge.DataSource.FirstRecord = i
        mail_merge.DataSource.LastRecord = i

        mail_merge.Destination = 0
        mail_merge.Execute(True)
    
        base_name=mail_merge.DataSource.DataFields(('Name'.replace(' ','_'))).Value
        targetDoc=wordApp.ActiveDocument
        

        #targetDoc.SaveAs2(os.path.join(destination_folder,base_name + '.docx'),16)
        targetDoc.ExportAsFixedFormat(os.path.join(destination_folder,base_name),exportformat:=17)

        targetDoc.Close(False)
        targetDoc = None
    sourceDoc.MailMerge.MainDocumentType= -1

    directory=destination_folder
    chdir(directory)

    pdf_list=glob('*.pdf')
    pdf_list_actualized = pdf_list
    for pdf in pdf_list:
        with fitz.open(pdf) as pdf_obj:
            text=pdf_obj[0].get_text()
        new_file_name=text.split("\n",1)[0].strip()
        nr_files_same = 1
        f_new_pdf = new_file_name+ '_'+str(nr_files_same) + '.pdf'
        while f_new_pdf in pdf_list_actualized:
            nr_files_same+=1
            f_new_pdf = new_file_name+ '_'+str(nr_files_same) + '.pdf'
        else:
            rename(pdf,f_new_pdf)
        pdf_list_actualized = glob('*.pdf')
    messagebox.showinfo(title="Completed",message=destination_folder)
    
    
    
    
canvas=tk.Canvas(root, height=300,width=500)
canvas.pack()


frame=tk.Frame(root,bg="#263D42")
frame.place(relwidth=.8,relheight=.2,relx=.1,rely=.1)

#button=tk.Button(frame,text="Create Inventory Records",height=1,width=30,bg="#ff2200",fg="#fff",pady=20,command=create_files,)
button=tk.Button(frame,text="Create Inventory Records",height=1,width=30,bg="#ff2200",fg="#fff",pady=20,command=create_files)
button_font=font.Font(size=20)
button["font"]=button_font
button.pack()




root.mainloop()