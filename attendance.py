from tkinter import *
from tkinter import filedialog, messagebox, ttk
import os
import pandas as pd
from io import StringIO
import re


root = Tk()
root.geometry("900x750")
root.title('Convert the txt to csv')
#root.iconbitmap()
root.pack_propagate(False)
root.resizable(0, 0)

path = os.getcwd()


#the load text file view
data_frame=LabelFrame(root, text="Text Data")
data_frame.place(height=600, width=240, rely=0.01, relx=0)

#the window to show the excel from session before 
execfile_frame = LabelFrame(root, text="Excel Data")
execfile_frame.place(height=600, width=630, rely=0.01, relx=0.28)

#frame of the buttons
file_frame = LabelFrame(root, text = "Open File")
file_frame.place(height=125, width=900, rely=0.82, relx=0)



#label
label_file = Label(file_frame, text = "No File Has Been Selected")
label_file.place(rely=0,relx=0)

#open button
Browse_button = Button(file_frame, text = "Browse The File", command=lambda: open_file())
Browse_button.place(rely= 0.65, relx=0.10)

#load button
open_button = Button(file_frame, text = "Load The txt File", command=lambda: load_txt())
open_button.place(rely= 0.65, relx=0.30)

#excel button openner
open_excel_button = Button(file_frame, text = "Load The Excel", command=lambda: load_excel())
open_excel_button.place(rely=0.65, relx=0.50)

#save button
save_button = Button(file_frame, text = "Save", command=lambda: save_txt())
save_button.place(rely= 0.65, relx=0.70)

#count the attendance of students
save_buttom = Button(file_frame, text = "Count", command=lambda: count_of_attendance())
save_buttom.place(rely= 0.65, relx=0.80)
#tree view
tv1 = ttk.Treeview(data_frame)
tv1.place(relheight=1, relwidth=1)

#make scrollbar for the text window
treescrolly = Scrollbar(data_frame, orient="vertical", command= tv1.yview)
treescrollx = Scrollbar(data_frame, orient="horizontal", command= tv1.xview)
tv1.configure(xscrollcommand=treescrollx.set, yscrollcommand=treescrolly.set)
treescrollx.pack(side="bottom", fill='x')
treescrolly.pack(side="right", fill="y")

def open_file():
    filename = filedialog.askopenfilename(initialdir="path", title="Open Text File", filetypes=(("Text Files", "*.txt"), ))
    label_file["text"] = filename
    return None



def load_txt():
    file_path = label_file["text"]
    global df

    #have the excel file loaded?
    global isSet 
    isSet = False
    
    #infrom the function that the var is global
    

    try:
        tex_filename = r"{}".format(file_path)
        #print("Here: ** ",tex_filename[-4:])
        if tex_filename[-4:]==".txt":
            df = pd.read_csv(tex_filename)
             
        else:
            print("*********do something*********")
            

    except ValueError:
        messagebox.showerror("Information", "The file you have chosen is invalid")
        return None
    except FileNotFoundError:
        messagebox.showerror("Information", "No such file as {file_path}")
        return None

    label.config(text="Text File is imported successfully to the Frame.")
    #manipulating the df
    #for  index, row in df.iterrows():
        #print(row)

    #remove the first row
    column_names = list(df.columns.values)
    column_names = column_names[0]
    m = re.search('(?<=at )(.*)', column_names)
    if not m:
        m = re.search('(?<=vom )(.*)', column_names)
    col_name = m.groups(0)[0]
    df.columns.values[0] = col_name
    df.drop([0,0], axis=0, inplace=True)

   
    
    
    clear_data() 
    tv1["column"] = list(df.columns)
    tv1["show"] = "headings"
    for column in tv1["columns"]:
        tv1.heading(column, text=column) # let the column heading = column name
        

    df_rows = df.to_numpy().tolist() # turns the dataframe into a list of lists
    for row in df_rows:
        #print(row[0])
        if row[0] == "Gerrit Burmester" or row[0] == "Sarmad Rezayat" or row[0] == "Sven Hartmann":
            pass
        elif row[0] == "Sorted by last name:" or row[0]=="Sortiert nach Nachname:":
            break
        else:
            tv1.insert("", "end", values=row) # inserts each list into the treeview. For parameters see https://docs.python.org/3/library/tkinter.ttk.html#tkinter.ttk.Treeview.insert

#to delete rows that we don't need
def clear_rows():
    global df4
    #remove row that sorted the names with lastname
    column_names = list(df.columns.values)
    a = column_names[0]
    print('here: ', a)
    for i in df.index:    
        if df.at[i, a] == "Sorted by last name:" or  df.at[i, a]=="Sortiert nach Nachname:":
            soreted_i = i - 1
            break

    
    df2 = df.iloc[:soreted_i].copy()

    #list of prof. lectror and tutors
    indx = []
    df3 = df2.copy()
    for i in df3.index:    
        if df2.at[i, a] == "Gerrit Burmester" or df3.at[i, a] == "Sarmad Rezayat" or df3.at[i, a] == "Sven Hartmann":
            #add rows to a list
            indx.append(i)

    for num in indx:
        df3.drop(num, axis=0, inplace=True)
        print(num)


    #drop all duplication
    df3 = df3.drop_duplicates(subset=None, keep='first', inplace=False, ignore_index=False)
    #display(df4)

    #reset the index   
    df4 = df3.reset_index(drop=True)   
    #print(df)
    #df4.index += 1

def clear_data():
    tv1.delete(*tv1.get_children())
    return None
    #create a dataframe from the string
    
def load_excel():
    
    filename2 = filedialog.askopenfilename(initialdir="path", title="Open Text File", filetypes=((("xlxs files", ".*xlsx"),)))
    if filename2:
      try:
         filename2 = r"{}".format(filename2)
         df = pd.read_excel(filename2)
      except ValueError:
         label.config(text="File could not be opened")
      except FileNotFoundError:
         label.config(text="File Not Found")

    # Clear all the previous data in tree
    clear_treeview()

    # Add new data in Treeview widget
    tree["column"] = list(df.columns)
    tree["show"] = "headings"

    # For Headings iterate over the columns
    for col in tree["column"]:
       tree.heading(col, text=col)

    # Put Data in Rows
    df_rows = df.to_numpy().tolist()
    for row in df_rows:
         tree.insert("", "end", values=row)

    tree.pack()

    
    

    #load the excel file to datafram 
    global df_excel 
    df_excel = pd.read_excel(filename2)
    global isSet
    isSet = True
    #print(df_excel)

# Clear the Treeview Widget
def clear_treeview():
   tree.delete(*tree.get_children())

# Create a Treeview widget
tree = ttk.Treeview(execfile_frame)


# Add a Label widget to display the file content
label = Label(file_frame, text='')
label.pack(pady=35)

def save_txt():
    clear_rows()
    print(isSet)
    if not isSet:
        df4.to_excel("listOfAttendances.xlsx")
        label.config(text="DataFrame is exported successfully to 'listOfAttendances.xlsx' Excel File.")
    else:
        #add the new list to the dataframe
        column_names = list(df4.columns.values)
        col_df4 = df4[column_names[0]]
        global export_DF
        export_DF = df_excel.join(col_df4)

        #delete the index column
        ex_col_names = list(export_DF.columns.values)
        a  = ex_col_names[0]
        export_DF.drop(a, axis=1, inplace=True)

        #count the attendance of each student
        last_df = export_DF.melt(id_vars=None).value.dropna().value_counts()
        df6 = pd.DataFrame({'all_attendances': last_df.index, 'Count': last_df.values})
        #df6
        result = export_DF.join(df6)
        #display(result)

        #save the result dataframe to the excel file
        result.to_excel("listOfAttendances.xlsx")
        label.config(text="DataFrame is added successfully to 'sample.xlsx' Excel File.")


def count_of_attendance():
    
    return None 

root.mainloop() 