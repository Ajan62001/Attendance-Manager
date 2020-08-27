from tkinter import *
import datetime
from tkinter import messagebox as tmsg
import xlrd

def calculate():
    workbook = xlrd.open_workbook("Attendance Report.xlsx")
    worksheet = workbook.sheet_by_index(0)

    present_sum = 0
    absent_sum = 0

    n = int(worksheet.nrows)
    print(n)
    for i in range(2, n):
        present_sum = present_sum + (worksheet.cell(i, 4).value)

    for i in range(2, n):
        absent_sum = absent_sum + (worksheet.cell(i, 6).value)
    global totalclass,bunkclass,currentAttendence
    total = eval(totalclass.get())
    k=0
    Attendance = ((total)*eval(currentAttendence.get()))//100
    if eval(currentAttendence.get())>100:
        k=1
        tmsg.showinfo("Error","Attendance is not greater than 100")
    if eval(bunkclass.get())<0 or eval(attainclass.get())<0:
        k=1
        tmsg.showinfo("Error", "Classes will be  A Positive Integer Value")
    if ( not bunkclass.get().isdigit()):
        tmsg.showinfo("Error", "Classes will be  A Integer Value")
    if (not attainclass.get().isdigit()):
        k=1
        tmsg.showinfo("Error", "Classes will be  A Integer Value")
    if k==0:
       try:
            updateAttendance = ((Attendance+eval(attainclass.get()))/(total+eval(bunkclass.get())+eval(attainclass.get())))*100
            tmsg.showinfo("Updated Attendance",f"Updated attendance will be {updateAttendance} .")
       except ZeroDivisionError:
            tmsg.showinfo("Error","Enter the details")

def save():
    file1 = open("Attendance.txt","a+")
    file1.write(str(eval(totalclass.get())+eval(attainclass.get())+eval(bunkclass.get())))
    file1.write(f" date {datetime.datetime.now()}\n ")
    file1.close()
    tmsg.showinfo("save","Total classes saved")

root = Tk()
root.geometry("280x200")
root.maxsize(280,220)
root.minsize(280,100)
root.title("AttendanceManager")


file1 = open("Attendance.txt","r")
linelist = file1.readlines()
file1.close()
lastline = linelist[-2]
l = [str(x) for x in lastline.split(" ")]


totalclass = StringVar()
totalclass.set(l[1])

bunkclass = StringVar()
bunkclass.set("0")

currentAttendence = StringVar()
currentAttendence.set("0")


attainclass = StringVar()
attainclass.set("0")

variableList=[totalclass,currentAttendence,bunkclass,attainclass]
labelList=["Total classes :","current attendance:","Bunk class :","Attain class "]
for i in range(4):
    Label(root,text=labelList[i],pady=10).grid(column=0)
    screen1 = Entry(root,textvariable=variableList[i],font="lucida 10 bold",relief=SUNKEN,borderwidth=3)
    screen1.grid(row = i,column=1)


frame = Frame(root)
Button(frame,text="Submit",command=calculate,padx=10,bg="skyblue").grid(row=4,column=1)
Button(frame,text="Save",command=save,padx=10,bg="skyblue").grid(row=4,column=2)
frame.grid(column=1)

def Help():
    tmsg.showinfo("Help","you can do arithmetic operation here")
    tmsg.showinfo("Help","you can save the total classes for your convenience")

def about():
    tmsg.showinfo("about","this application is created by AJAN")

def info():
    tmsg.showinfo("info",f"Total classes updated on {l[3]}")

mainmenu = Menu(root)
mainmenu.add_command(label="info",command=info)
mainmenu.add_command(label="Help",command=Help)
mainmenu.add_command(label="About",command=about)
root.config(menu=mainmenu)

root.mainloop()
