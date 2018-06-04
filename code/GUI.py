from tkinter import *
from tkinter import filedialog
import functions

import os



window=Tk()

def s_add():
    some=filedialog.askdirectory()
    st.delete('1.0', END)
    return st.insert(END, some)
def s_add2():
    some2=filedialog.askdirectory()
    dt.delete('1.0', END)
    return dt.insert(END, some2)
def s_add3():
    some3=filedialog.askopenfilename()
    incident_table.delete('1.0', END)
    return incident_table.insert(END, some3)

b1=Button(window,command=s_add,text="Browse")
b1.grid(row=1,column=5)
b2=Button(window,command=s_add2,text="Browse")
b2.grid(row=2,column=5)
dim_source=Button(window,command=s_add3,text="Browse")
dim_source.grid(row=5,column=5)
dim_source.grid_remove()
window.minsize(500,200)
source=Label(window,text="Select the source file",width=15)
source.grid(row=1,column=3)
destination=Label(window,text="Select where you want to save the file")
destination.grid(row=2,column=3)

st=Text(window,width=20, height=1)
st.grid(row=1,column=4)
dt=Text(window,width=20, height=1)
dt.grid(row=2,column=4)

def exec():
    '''st=source text, dt = destination text.., the text feild in tkinter adds up '\n' and this function deletes that'''
    new1=dt.get("1.0",END)
    new=new1.rstrip("\n")
    final = r"/combined.xls"
    newfinal = new + final
    path1 = st.get("1.0",END)
    path=path1.rstrip("\n")

    main_Funtion(path,newfinal)
    return

def main_Funtion(path,newfinal):
    '''we are calling the functions that were written'''
    for filenameList in os.listdir(path):
        if os.path.isdir(os.path.join(path, filenameList)):
            path = os.path.join(path, filenameList)
    functions.combing_files(path, newfinal)
    functions.top_row(newfinal)
    functions.drop_duplicates(newfinal, "last")
    functions.end_time(newfinal)
    functions.new_resolved_cols(newfinal)
    functions.res_row(newfinal)
    functions.correct_dept(newfinal)
    functions.end_state(newfinal)
    functions.format_date(newfinal)
    functions.drop_duplicates(newfinal, "last")
    functions.time_diff(path, newfinal)
    functions.aggregate(newfinal)
    functions.timetaken(newfinal)
    functions.database_operations()
    functions.SLA_Counter(newfinal)
    #functions.testing(path, newfinal)

    return

b3=Button(window,text="Execute",command=exec)
b3.grid(row=3,column=4)

def copy_files():
    incident_table.grid()
    dim_source.grid()
    add.grid()
    return
def upload_dim():
    functions.dim_table(incident_table.get("1.0", END).rstrip("\n"))
    return

incident_table=Text(window,width=20, height=1)
b4=Button(window, text="Add to Dim table", width=15,command=copy_files)
b4.grid(row=5, column=3)

add=Button(window, text="Add", width=15,command=upload_dim)
add.grid(row=6, column=4)
add.grid_remove()
incident_table.grid(row=5,column=4)
incident_table.grid_remove()
window.mainloop()

