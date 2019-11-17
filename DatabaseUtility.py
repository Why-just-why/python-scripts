import mysql.connector as MySQLdb
import xlrd
from tkinter import ttk
from tkinter import filedialog,Checkbutton
from tkinter import Tk as ThemedTk
from tkinter import StringVar,IntVar
from tkinter import END
from tkinter import Toplevel
# from ttkthemes import ThemedTk, THEMES

import datetime

def MsgBox(Message): 
    child = Toplevel(root) 
    child.resizable(False,False)
    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()
    positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
    positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
    child.geometry("+{}+{}".format(positionRight, positionDown))
    ttk.Label(child,text = Message,anchor='center').grid(row=1, sticky='nwes')
    def close_win():
        child.destroy()
    ttk.Button(child,text='OK',width=25,command=close_win).grid(row=2,sticky='nwes')
    child.protocol("WM_DELETE_WINDOW", close_win)


def customConfigureRows(): 
    child = Toplevel(root) 
    child.resizable(False,False)
    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()
    positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
    positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
    child.geometry("+{}+{}".format(positionRight, positionDown))
    ttk.Label(child,text = "Configure Row range" ,anchor='center',width=25).grid(row=0,columnspan=2, sticky='nwes')
    ttk.Label(child,text = "Enter Starting Range",width=20).grid(row=1, sticky='nwes')
    start=ttk.Entry(child,width=5)
    start.grid(row=1,column=1,sticky='nwes')
    ttk.Label(child,text = "Enter Ending Range",width=20).grid(row=2, sticky='nwes')
    end=ttk.Entry(child,width=5)
    end.grid(row=2,column=1,sticky='nwes')
    def getvalues():
        global rowstart,rowend
        rowstart=start.get()
        rowend=end.get()
        child.destroy()
    ttk.Button(child,text='OK',width=25,command=getvalues).grid(row=3,columnspan=2,sticky='nwes')
    child.protocol("WM_DELETE_WINDOW", child.destroy)

def customConfigureColumns(): 
    child = Toplevel(root) 
    child.resizable(False,False)
    windowWidth = root.winfo_reqwidth()
    windowHeight = root.winfo_reqheight()
    positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
    positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
    child.geometry("+{}+{}".format(positionRight, positionDown))
    ttk.Label(child,text = "Configure Column range" ,anchor='center',width=25).grid(row=0,columnspan=2, sticky='nwes')
    ttk.Label(child,text = "Enter Starting Range",width=20).grid(row=1, sticky='nwes')
    start=ttk.Entry(child,width=5)
    start.grid(row=1,column=1,sticky='nwes')
    ttk.Label(child,text = "Enter Ending Range",width=20).grid(row=2, sticky='nwes')
    end=ttk.Entry(child,width=5)
    end.grid(row=2,column=1,sticky='nwes')
    def getvalues():
        global Columnstart,Columnend
        Columnstart=start.get()
        Columnend=end.get()
        child.destroy()
    ttk.Button(child,text='OK',width=25,command=getvalues).grid(row=3,columnspan=2,sticky='nwes')
    child.protocol("WM_DELETE_WINDOW", child.destroy)    



def Execute():
    try:
        def dateConv(cellvalue):
            return datetime.datetime(*xlrd.xldate_as_tuple(cellvalue, book.datemode))
        # Open the workbook and define the worksheet
        book = xlrd.open_workbook(displayfile.get())
        sheet = book.sheet_by_name(combo.get())

        # Establish a MySQL connection
        database = MySQLdb.connect (host=str(getHost.get()), user = str(getUsername.get()), passwd = str(getPassword.get()), db = str(getDatabaseName.get()))
        # Get the cursor, which is used to traverse the database, line by line
        cursor = database.cursor(buffered=True)

        cursor.execute('select * from ' + getTableName.get() +'')
        b = ''
        names = tuple(map(lambda x: x[0], cursor.description))

        if var1.get():
            rows = range(1,sheet.nrows)
        elif not var1.get():
            rows = range(int(rowstart),int(rowend)+1)

        if var2.get():
            clms = range(0,sheet.ncols)
            for _i in names:
                b=",".join(names)
            c = '%s,' * len(names)
            c = c[0:len(c)-1]
        elif not var2.get():
            clms = range(int(Columnstart),int(Columnend)+1)
            customnames = tuple(str(names[i]) for i in range(int(Columnstart),int(Columnend)+1))
            for _i in customnames:
                b=",".join(customnames)
            c = '%s,' * len(customnames)
            c = c[0:len(c)-1]


        query = 'INSERT INTO {} ({}) VALUES ({})'.format(getTableName.get(), b,c)

        # Create a For loop to iterate through each row in the XLS file, starting at row 2 to skip the headers
        for r in rows:
            values = []
            for c in clms:
                if(sheet.cell(r,c).ctype==xlrd.XL_CELL_DATE):
                    a=str(dateConv(sheet.cell(r,c).value))
                    values.append(a)
                elif(sheet.cell(r,c).ctype==xlrd.XL_CELL_NUMBER):
                    a=int(sheet.cell(r,c).value)
                    values.append(a)
                else:
                    values.append(sheet.cell(r,c).value)
            # Assign values from each row
            values = tuple(values)
            # Execute sql Query
            cursor.execute(query, values)

        # Close the cursor
        cursor.close()

        child = Toplevel(root) 
        child.resizable(False,False)
        windowWidth = root.winfo_reqwidth()
        windowHeight = root.winfo_reqheight()
        positionRight = int(root.winfo_screenwidth()/2 - windowWidth/2)
        positionDown = int(root.winfo_screenheight()/2 - windowHeight/2)
        child.geometry("+{}+{}".format(positionRight, positionDown))
        ttk.Label(child,text = 'Would you like to commit the Data',anchor='center').grid(row=0,columnspan=2,sticky='nwes')
        def close_win():
            # Close the database connection
            database.close()
            child.destroy()
            MsgBox('Action Aborted')
            
        def commit():
            # Commit the transaction
            database.commit()
            # Close the database connection
            database.close()
            child.destroy()
            MsgBox('Successfully Completed')

        ttk.Button(child,text='Yes',width=15,command=commit).grid(row=1,sticky='nwes')
        ttk.Button(child,text='No',width=15,command=close_win).grid(row=1,column=1,sticky='nwes')
        child.protocol("WM_DELETE_WINDOW", close_win)

        
    except FileNotFoundError:
        MsgBox('No File Found')
    except xlrd.biffh.XLRDError:
        MsgBox('No Sheet Selected')

def settings():
    combo.configure(state='normal')
    file = xlrd.open_workbook(displayfile.get(), on_demand=True)
    sheets = file.sheet_names()
    getsheet.set(sheets)

def openfile():
    root.filename =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("excel files","*.xlsx"),("all files","*.*")))
    getfile.set(root.filename)

def configurerows():
    if var1.get():
        configureRows.configure(state='disabled')
    elif not var1.get():
        configureRows.configure(state='normal')

def configurecolumns():
    if var2.get():
        configureColumns.configure(state='disabled')
    elif not var2.get():
        configureColumns.configure(state='normal')


if __name__ == "__main__":
    root = ThemedTk()
    root.title("")
    root.resizable(False,False)
    #row0
    selectPath = ttk.Button(root,width=50, text='Select File',command=openfile)
    selectPath.grid(row=0,columnspan=2,sticky='nwes')
    #row1
    getfile = StringVar()
    displayfile = ttk.Entry(root,width=50,textvariable=getfile)
    displayfile.grid(row=1,columnspan=2,sticky='nwes')
    displayfile.config(state='readonly')
    #row2
    selectsheet = ttk.Button(root,width=25,text = 'Select Sheet',command=settings)
    selectsheet.grid(row=2,sticky='nwes')
    #retrive data dummy method
    getsheet = StringVar()
    displaysheet = ttk.Entry(root,width=25,textvariable=getsheet)
    displaysheet.grid(row=2,column=1,sticky='nwes')
    displaysheet.config(state='readonly')

    temp = StringVar()
    temp.set("Select")

    combo = ttk.Combobox(root,width=25,textvariable=temp,values=displaysheet.get(),postcommand=lambda: combo.config(values=displaysheet.get(),state='readonly'))
    combo.configure(state='disabled')
    combo.grid(row=2,column=1,sticky='nwes')

    #row3
    var2 = IntVar()
    loadColumns = ttk.Checkbutton(root,width=25,text="Load All Columns",variable=var2,command=configurecolumns)
    loadColumns.grid(row=3, sticky='nwes')

    configureColumns = ttk.Button(root,width=25,text = 'Configure Columns',command=customConfigureColumns)
    configureColumns.grid(row=3,column=1,sticky='nwes')

    #row4
    var1 = IntVar()
    loadRows = ttk.Checkbutton(root,width=25,text="Load All Rows",variable=var1,command=configurerows)
    loadRows.grid(row=4, sticky='nwes')

    configureRows = ttk.Button(root,width=25,text = 'Configure Rows',command=customConfigureRows)
    configureRows.grid(row=4,column=1,sticky='nwes')

    #row5
    selectDB = ttk.Label(root,text = 'Select Database',anchor='center',width=25)
    selectDB.grid(row=5,sticky='nwes')

    default = StringVar()
    default.set("Select")
    DBcombo = ttk.Combobox(root,textvariable=default,width=25,values=["MySQLDb"])
    DBcombo.grid(row=5,column=1,sticky='nwes')
    DBcombo.config(state='readonly')

    #row6
    Host = ttk.Label(root,text = 'Host Name',anchor='center',width=25)
    Host.grid(row=6,sticky='nwes')
    getHost = ttk.Entry(root,width=25)
    getHost.grid(row=6,column=1,sticky='nwes')

    #row7
    Username = ttk.Label(root,text = 'User Name',anchor='center',width=25)
    Username.grid(row=7,sticky='nwes')
    getUsername = ttk.Entry(root,width=25)
    getUsername.grid(row=7,column=1,sticky='nwes')

    #row8
    Password = ttk.Label(root,text = 'Password',anchor='center',width=25)
    Password.grid(row=8,sticky='nwes')
    getPassword = ttk.Entry(root,width=25)
    getPassword.grid(row=8,column=1,sticky='nwes')

    #row9
    DatabaseName = ttk.Label(root,text = 'Database Name',anchor='center',width=25)
    DatabaseName.grid(row=9,sticky='nwes')
    getDatabaseName = ttk.Entry(root,width=25)
    getDatabaseName.grid(row=9,column=1,sticky='nwes')

    #row10
    TableName = ttk.Label(root,text = 'Table Name',anchor='center',width=25)
    TableName.grid(row=10,sticky='nwes')
    getTableName = ttk.Entry(root,width=25)
    getTableName.grid(row=10,column=1,sticky='nwes')

    #row11
    Execute = ttk.Button(root,width=25,text = 'Execute',command=Execute)
    Execute.grid(row=11,columnspan=2,sticky='nwes')

    root.mainloop()