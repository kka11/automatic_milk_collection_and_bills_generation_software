from tkinter import *
import tkinter as tk
from tkinter.filedialog import askopenfilename
import sys, subprocess
import csv
import os
import pypyodbc
import xlrd
import openpyxl 
#from xlsxwriter.workbook import Workbook
from tkinter import messagebox
from datetime import datetime, timedelta,date
from openpyxl import Workbook


config=open('config.cfg','r')
path=config.read()



'''for i in range(nrows-1):
    print (member[i][1])'''

currentdate=datetime.date(datetime.now())

root = Tk()
root.title("MDBCS Software - Harshit ")
imgicon = PhotoImage(file=os.path.join('2.gif'))
root.tk.call('wm', 'iconphoto', root._w, imgicon) 

config=open('config.cfg','r')
path=config.read()




w = 600 # width for the Tk root
h = 400 # height for the Tk root

# get screen width and height
ws = root.winfo_screenwidth() # width of the screen
hs = root.winfo_screenheight() # height of the screen

# calculate x and y coordinates for the Tk root window
x = (ws/2) - (w/2)
y = (hs/2) - (h/2)

# set the dimensions of the screen 
# and where it is placed
root.geometry('%dx%d+%d+%d' % (w, h, x, y))





        


def centerCodeValidate(cCode):
    if(cCode!=""):
        if os.path.exists(path+cCode):
            #messagebox.showinfo("Info","path exists")
            return True
        else:
            messagebox.showerror("Error","This Center does not exist")
            return False
    else:
        messagebox.showerror("Error","Please Enter Center Code")
        return False
    


def validate(date_te):
	date_text = date_te
	try :
		if date_text != datetime.strptime(date_text, "%Y-%m-%d").strftime('%Y-%m-%d'):
			raise ValueError
		#print("Date is Correct")
		return True

	except ValueError:
		#print("Value Error")
		return False






def dateValidate():
	#print("At Step 1")
	date_1 = dat.get()
	date_2 = dat_2.get()
	val_1 = validate(date_1)
	val_2 = validate(date_2)
	if(val_1 == 0 and val_2 == 0):
		#print("Starting and End date has incorrect input")
		messagebox.showerror("Error" , "Start date and End date has incorrect input")
		return 0
		
	elif(val_1 == 0):
		#print("Starting date has incorrect input")
		messagebox.showerror("Error" , "Start date has incorrect input")
		return 0

	elif(val_2 == 0):
		#print("End date has incorrect input")
		messagebox.showerror("Error" , "End date has incorrect input")
		return 0

	elif(date_1 > date_2):
		#print("Starting date Should be before last date")
		messagebox.showerror("Error" , "Start date should be before End date ")
		return 0


	elif(val_1 == 1 and val_2 == 1):
		#print("Both Input are in correct Format")
		return 1

	#print("YES")










def convert(cy):

    member=[]
    if(cy==0):
        cv=centerCodeValidate(dat_3.get())
        if(cv==True):
            dv=dateValidate()
            if(dv!=0):
                c = dat_3.get()
                #label_c = Label()
                #label_c.pack()
                s_date = dat.get()
                #label_s_date = Label()
                #label_s_date.pack()
                
                e_date = dat_2.get()
                #label_e_date = Label()
                #label_e_date.pack()
                

                '''s_date="2017-08-04"
                e_date="2017-09-02"
                c="1"'''

                        
                workbook = xlrd.open_workbook(path+c+"/member.xlsx")
                sheet = workbook.sheet_by_index(0)
                nrows=sheet.nrows
                
                #print (nrows)
                for rowx in range(1,sheet.nrows):
                    cols = sheet.row_values(rowx)
                    t=[int(cols[0]),cols[1]]
                    member.append(t)
            else:
                return
        else:
            return
        
                        
        
    elif(cy==1):
        if not centerCodeValidate(dat_4.get()):
            return 
        c = dat_4.get()
        #label_c = Label()
        #label_c.pack()

        
        workbook = xlrd.open_workbook(path+c+"/member.xlsx")
        sheet = workbook.sheet_by_index(0)
        nrows=sheet.nrows
        
        #print (nrows)
        for rowx in range(1,sheet.nrows):
            cols = sheet.row_values(rowx)
            t=[int(cols[0]),cols[1]]
            member.append(t)

        if(currentdate.day >= 1 and currentdate.day <= 10):
            startDay = 21
            endDay = 31
            if( currentdate.month == 1 ):
                    ourMonth = 12
                    ourYear = currentdate.year -1 
            else :
                    ourMonth = currentdate.month - 1
                    ourYear = currentdate.year
        elif(currentdate.day >= 11 and currentdate.day <= 20):
                startDay = 1
                endDay = 10
                ourMonth = currentdate.month
                ourYear = currentdate.year
        elif(currentdate.day >= 21 and currentdate.day <= 31):
                startDay = 11
                endDay = 20
                ourMonth = currentdate.month
                ourYear = currentdate.year
                ini_date = date(ourYear , ourMonth , startDay)
        s_date = ini_date.strftime('%Y-%m-%d')
        fin_date = date(ourYear , ourMonth , endDay)
        e_date = fin_date.strftime('%Y-%m-%d')
    else:
        exit()

    

    #DATABASE = askopenfilename()
    # MS ACCESS DB CONNECTION
    con=pypyodbc.connect("DRIVER={Microsoft Access Driver (*.mdb)};UID=admin;UserCommitSync=Yes;Threads=3;SafeTransactions=0;PageTimeout=5;MaxScanRows=8;MaxBufferSize=2048;FIL={MS Access};DriverId=281;DefaultDir="+path+c+"/;DBQ="+path+c+"/Reilsams.mdb")

    # OPEN CURSOR AND EXECUTE SQL
    cur = con.cursor()
    fields=[]
    for col in cur.columns(table='YearlySumm'):
            fields.append(col[3])
    #print (fields)
    cur.execute("SELECT * FROM YearlySumm")
    with open(path+c+'/temp/YearlySumm.csv', 'w',newline='') as f:
        writer = csv.writer(f)
        writer.writerow(fields)
        # OPEN CSV AND ITERATE THROUGH RESULTS
        for row in cur.fetchall() :
            writer.writerow(row)

    cur.close()
    con.close()





    r = csv.reader(open(path+c+'/temp/YearlySumm.csv'))
    lines = [l for l in r]

    newrows = len(lines)
    newcolumns = len(lines[0])


    for rowindex in range(len(lines)):
            for columnindex in range (len(lines[0])):
                    if rowindex > 0:
                            #print ("Row index = " + str(rowindex))
                            if (columnindex == 1 or columnindex == 2 or columnindex == 3 or columnindex == 8 or columnindex == 9 or columnindex ==10):
                                    a = lines[rowindex][columnindex]
                                    aINfloat = float(a)
                                    b = "%.2f" % aINfloat
                                    lines[rowindex][columnindex] = b
                            elif (columnindex == 4):
                                    string1 = lines[rowindex][columnindex]
                                    string2 = ""
                                    for k in string1:
                                            if(k == " "):
                                                    break;
                                            else:
                                                    string2 = string2 + k;
                                    lines[rowindex][columnindex] = string2;


    file = (open(path+c+'/temp/corr_output.csv','w',newline=''))
    writer=csv.writer(file)
    writer.writerows(lines)
    file.close()

    reader = csv.reader(open(path+c+'/temp/corr_output.csv','r'))
    #messagebox.showinfo("Information" , "CSV File in Correct Foramt is generated")
    with open(path+c+'/temp/Final Date-wise Sheet.csv', 'w',newline='')as f:
        writer = csv.writer(f)
        
        fields=(next(reader))
        fields.append('')
        #print(fields)
        nfields=0
        for col in fields:
            nfields=nfields+1
        #print(nfields)
        tf=fields
        for i in reversed(range(1,nfields)):
            fields[i]=fields[i-1]
        fields[1]='Name'
        
        
        
        fields=fields+['Deduction','Fodder','Description','','']
        #print (fields)
        writer.writerow(fields)
        curr_v_date= (row[0][4] for row in reader)
    #print(curr_v_date) 


        PrevDateExistance = 0


        for row in reader:
            row.append('')
            if(row[4]==curr_v_date):
                if(row[4]>=s_date and row[4]<=e_date):
                    PrevDateExistance = 1
                    for i in reversed(range(1,nfields)):
                        row[i]=row[i-1]
                    for i in member:
                        if(str(row[1])==str(i[0])):
                            row[1]=i[1]
                    
                    ##### first check Mcod here Add like fields+[name](Do nothing Right now)
                    writer.writerow(row)

                    
            else:    
                
                if(PrevDateExistance == 1):
                    writer.writerow([])
                    writer.writerow(['','','','Total','','','','','','','','','','','','Amul Payment']) # Keep a Variable for unneccessity
                    writer.writerow([])
                    writer.writerow(fields)
                    PrevDateExistance = 0
                curr_v_date=row[4]
        writer.writerow([])
        writer.writerow(['','','','TOTAL','','','','','','','','','','','','Amul Payment'])

        #writer.writerow(['','','TOTAL','','','','','','','','','','','','Amul Payment'])
    #os.chdir(path+c+'/temp/')
    f.close()
    tpath=path+c+'/temp/'
    
    npath=path+c+'/'
    
    
            
    
    wb = Workbook()
    ws = wb.active
    with open(tpath+'Final Date-wise Sheet.csv', 'r') as f:
        for row in csv.reader(f):
            ws.append(row)
    wb.save(npath+'Final Sheet.xlsx')
    messagebox.showinfo("Information" , "Your file is Saved in \n"+path+c)








def raise_frame(frame):
    frame.tkraise()


f1 = Frame(root)
f2 = Frame(root)
f3 = Frame(root)
for frame in (f1 , f2 , f3):
    frame.grid(row = 0 , column = 0 , sticky = 'news')
img=tk.PhotoImage(file='2.gif')
label_entry_10 = Label(f1 ,image=img)
label_entry_10.pack(side=tk.LEFT)

label_entry_10 = Label(f2 ,image=img)
label_entry_10.pack(side=tk.LEFT)

label_entry_10 = Label(f3 ,image=img)
label_entry_10.pack(side=tk.LEFT)



 
Button(f1 , text = "By Date" ,width=15, height=2,font=(16), command = lambda:raise_frame(f2)).pack(side=tk.LEFT, padx=5, pady=5)
Button(f1 , text = "By Payment Cycle" ,width=15, height=2,font=(16), command = lambda:raise_frame(f3)).pack(side=tk.RIGHT, padx=5, pady=5)

label_entry_3 = Label(f2 , text = "Enter Center Code.",font=(16))
label_entry_3.pack(side=tk.TOP, padx=5, pady=5)
Lf=('',20)
dat_3 = StringVar()

dateEntry_3 = Entry(f2 , width=4,textvariable = dat_3,font=Lf)
dateEntry_3.pack(side=tk.TOP, padx=3, pady=3)

label_entry_1 = Label(f2 , text = "Enter Starting Date.",font=(16))
label_entry_1.pack(side=tk.TOP, padx=3, pady=3)
dat = StringVar()
dateEntry = Entry(f2 , textvariable = dat,width=10,font=Lf)
dateEntry.pack(side=tk.TOP, padx=3, pady=3)

label_entry_2 = Label(f2 , text = "Enter last Date.",font=(16))
label_entry_2.pack(side=tk.TOP, padx=3, pady=3)
dat_2 = StringVar()
dateEntry_2 = Entry(f2 ,width=10, textvariable = dat_2,font=Lf)
dateEntry_2.pack(side=tk.TOP, padx=3, pady=3)







label_entry_4 = Label(f3 , text = "Enter Center Code.",font=(16))
label_entry_4.pack(side=tk.TOP, padx=3, pady=3)
dat_4 = StringVar()
dateEntry_4 = Entry(f3 ,width=4, textvariable = dat_4,font=Lf)
dateEntry_4.pack(side=tk.TOP, padx=3, pady=3)

button_1 = Button(f2 , text = "Convert",font=(16) ,width=15, height=2, command = lambda : convert(0))
button_1.pack(side=tk.TOP, padx=3, pady=3)
Button(f2 , text = "By Payment Cycle" ,font=(16), width=15, height=2,command = lambda:raise_frame(f3)).pack(side=tk.TOP, padx=3, pady=3)

button_4 = Button(f3 , text = "Payment Cycle ",font=(16) ,width=15, height=2, command = lambda : convert(1)).pack(side=tk.TOP, padx=3, pady=3)
#button_4.pack()


Button(f3 , text = "By Date",font=(16) ,width=15, height=2, command = lambda:raise_frame(f2)).pack(side=tk.TOP, padx=3, pady=3)


raise_frame(f1)
    







root.mainloop()
