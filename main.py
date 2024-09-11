from tkinter import *
from PIL import ImageTk
from tkinter import messagebox, filedialog
import os
import pandas as pd 
import email_function
import time
class BULK_EMAIL:
    def __init__(self, root):
        self.root=root
        self.root.title("BULK EMAIL APPLICATION")
        self.root.geometry("1000x550+200+50")
        self.root.resizable(False, False)
        self.root.config(bg = "white")
        #----------Icons-------------
        self.email_icon = ImageTk.PhotoImage(file="Images/email1.png") 
        self.setting_icon = ImageTk.PhotoImage(file="Images/setting1.png")
        #----------Title-------------
        title = Label(self.root, text="Mailfoundr",image = self.email_icon, padx=15, pady=8, compound=LEFT, font=("Goudy Old Style", 50, "bold"), bg="#222A35", fg="white", anchor="w").place(x=0, y=0, relwidth=1)
        desc = Label(self.root, text="Send bulk emails with one click using an Excel file. Email column must be 'Email' v1.0", font=("Calibri (Body)", 14, "bold"), bg="#01B98D", fg="white").place(x=0, y=80, relwidth=1)

        btn_setting = Button(self.root, image=self.setting_icon, bg="#222A35", bd=0, cursor="hand2", command=self.setting_window).place(x=900, y=5)


        self.var_choice=StringVar()
        single=Radiobutton(self.root, text="Single", value="single",variable=self.var_choice, font=("times new roman", 30, "bold"), bg= "white", fg="#222A35", command=self.check_single_or_bulk).place(x=50, y=150)
        bulk=Radiobutton(self.root, text="Bulk", value="bulk", variable=self.var_choice, font=("times new roman", 30, "bold"), bg= "white", fg="#222A35", command=self.check_single_or_bulk).place(x=250, y=150)
        self.var_choice.set("single")

        to=Label(self.root, text="To (Email Address)", font=("times new roman", 18), bg="white", fg="#222A35").place(x=50, y=250)
        subj=Label(self.root, text="Subject", font=("times new roman", 18), bg="white", fg="#222A35").place(x=50, y=300)
        msg=Label(self.root, text="Message", font=("times new roman", 18), bg="white", fg="#222A35").place(x=50, y=350)

        self.txt_to=Entry(self.root, font=("times new roman", 18),bg="white", fg="#2B2B2B", insertbackground="#2B2B2B")
        self.txt_to.place(x=260, y=250, width=350, height=30)

        self.btn_browse = Button(self.root, command=self.browse_file, text="BROWSE", font=("times new roman", 18, "bold"), highlightbackground="#2B2B2B", fg="#2B2B2B", cursor="hand2", state=DISABLED)
        self.btn_browse.place(x=620, y=250, width=120)

        self.txt_subj=Entry(self.root, font=("times new roman", 18),bg="white", fg="#2B2B2B", insertbackground="#2B2B2B")
        self.txt_subj.place(x=260, y=300, width=450, height=30)

        self.txt_msg=Text(self.root, font=("times new roman", 18),bg="white", fg="#2B2B2B", insertbackground="#2B2B2B")
        self.txt_msg.place(x=260, y=350, width=650, height=120)

        #-----------Status--------------------
        self.lbl_total=Label(self.root, font=("times new roman", 18), bg="white", fg="#2B2B2B")
        self.lbl_total.place(x=50, y=490)

        self.lbl_sent=Label(self.root, font=("times new roman", 18), bg="white", fg="green")
        self.lbl_sent.place(x=280, y=490)

        self.lbl_left=Label(self.root, font=("times new roman", 18), bg="white", fg="orange")
        self.lbl_left.place(x=400, y=490)

        self.lbl_failed=Label(self.root, font=("times new roman", 18), bg="white", fg="red")
        self.lbl_failed.place(x=520, y=490)
        

        btn_clear = Button(self.root, text="CLEAR", command=self.clear1, font=("times new roman", 18, "bold"), highlightbackground="#2B2B2B", fg="#2B2B2B", cursor="hand2").place(x=700, y=490, width=120)

        btn_send = Button(self.root, command=self.send_email, text="SEND", font=("times new roman", 18, "bold"), highlightbackground="#2B2B2B", fg="#2B2B2B", cursor="hand2").place(x=830, y=490, width=120)

        self.check_file_exist()
    
    def browse_file(self):
        op = filedialog.askopenfile(initialdir='/', title="Select Excel File for Emails", filetypes=(("All files", "*.*"), ("Excel files", ".xlsx")))
        if op!=None:
            data=pd.read_excel(op.name)
            if 'Email' in data.columns:
                self.emails = list(data['Email'])
                c=[]
                for i in self.emails:
                    if pd.isnull(i)==False:
                        c.append(i)
                self.emails = c
                if len(self.emails)>0:
                    self.txt_to.config(state=NORMAL)
                    self.txt_to.delete(0, END)
                    self.txt_to.insert(0, str(op.name.split("/")[-1]))
                    self.txt_to.config(state="readonly")
                    self.lbl_total.config(text="TOTAL: "+ str(len(self.emails)))
                    self.lbl_sent.config(text="SENT: ")
                    self.lbl_left.config(text="LEFT: ")
                    self.lbl_failed.config(text="FAILED: ")
                else:
                    messagebox.showerror("Error", "This file doesnot have any emails", parent=self.root)
            else:
                messagebox.showerror("Error", "Please select file which have email columns", parent=self.root)


    def send_email(self):
        x=len(self.txt_msg.get('1.0', END))
        if self.txt_to.get()=="" or self.txt_subj.get()=="" or x==1:
              op=messagebox.showerror("Error", "All Fields are required", parent = self.root)
        else:
            if self.var_choice.get()=="single":
                status = email_function.email_send_function(self.txt_to.get(), self.txt_subj.get(), self.txt_msg.get('1.0', END), self.from_, self.pass_)
                if status=="s":
                    messagebox.showinfo("Success", "Email has been sent", parent = self.root)
                if status=="f":
                    messagebox.showinfo("Success", "Email has not been sent, Try Again", parent = self.root)
            if self.var_choice.get()=="bulk":
                self.failed=[]
                self.s_count=0
                self.f_count=0
                for x in self.emails:
                    status=email_function.email_send_function(x, self.txt_subj.get(), self.txt_msg.get('1.0', END), self.from_, self.pass_)
                    if status=="s":
                        self.s_count+=1
                    if status=="f":
                        self.f_count+=1
                    self.status_bar()
                messagebox.showinfo("Success", "Email has been sent, Please Check Status", parent = self.root)

    def status_bar(self):
        self.lbl_total.config(text="STATUS: "+ str(len(self.emails))+ "=>>")
        self.lbl_sent.config(text="SENT: "+ str(self.s_count))
        self.lbl_left.config(text="LEFT: "+ str(len(self.emails)-(self.s_count+self.f_count)))
        self.lbl_failed.config(text="FAILED: "+ str(self.f_count))
        self.lbl_total.update()
        self.lbl_sent.update()
        self.lbl_left.update()
        self.lbl_failed.update()
        
    def check_single_or_bulk(self):
        if self.var_choice.get()=="single":
            self.btn_browse.config(state=DISABLED)
            self.txt_to.config(state=NORMAL)
            self.txt_to.delete(0, END)
            self.clear1()
        if self.var_choice.get()=="bulk":
            self.btn_browse.config(state=NORMAL)
            self.txt_to.delete(0, END)
            self.txt_to.config(state="readonly")

    def clear1(self):
        self.txt_to.config(state=NORMAL)
        self.txt_to.delete(0, END)
        self.txt_subj.delete(0, END)
        self.txt_msg.delete('1.0', END)
        self.var_choice.set("single")
        self.btn_browse.config(state=DISABLED)
        self.lbl_total.config(text="")
        self.lbl_sent.config(text="")
        self.lbl_left.config(text="")
        self.lbl_failed.config(text="")

    def setting_window(self):
        self.check_file_exist()
        self.root2 = Toplevel()
        self.root2.title("Setting - Bulk Email Application")
        self.root2.geometry("700x350+350+90")
        self.root2.focus_force()
        self.root2.grab_set()
        self.root2.config(bg="white")
        title2 = Label(self.root2, text="Credentials Setting",image = self.setting_icon, padx=15, pady=8, compound=LEFT, font=("Goudy Old Style", 50, "bold"), bg="#222A35", fg="white", anchor="w").place(x=0, y=0, relwidth=1)
        desc2 = Label(self.root2, text="Please enter the email address and password that will be used to send all emails", font=("Calibri (Body)", 14, "bold"), bg="#01B98D", fg="white").place(x=0, y=80, relwidth=1)
        from_=Label(self.root2, text="Email Address", font=("times new roman", 18), bg="white", fg="#222A35").place(x=50, y=150)
        pass_=Label(self.root2, text="Password", font=("times new roman", 18), bg="white", fg="#222A35").place(x=50, y=200)

        self.txt_from=Entry(self.root2, font=("times new roman", 18),bg="white", fg="#2B2B2B", insertbackground="#2B2B2B")
        self.txt_from.place(x=230, y=150, width=350, height=30)
        self.txt_pass=Entry(self.root2, font=("times new roman", 18),bg="white", fg="#2B2B2B", insertbackground="#2B2B2B", show="*")
        self.txt_pass.place(x=230, y=200, width=350, height=30)

        btn_clear2 = Button(self.root2, command=self.clear2, text="CLEAR", font=("times new roman", 18, "bold"), highlightbackground="#FE0919", fg="#FE0919", cursor="hand2").place(x=300, y=260, width=120)

        btn_save = Button(self.root2, command=self.save_setting, text="SAVE", font=("times new roman", 18, "bold"), highlightbackground="#01B98D", fg="#01B98D", cursor="hand2").place(x=430, y=260, width=120)

        self.txt_from.insert(0, self.from_)
        self.txt_pass.insert(0, self.pass_)
    
    def clear2(self):
        self.txt_from.delete(0, END)
        self.txt_pass.delete(0, END)
    
    def check_file_exist(self):
        if os.path.exists("important.txt")==False:
            f=open("important.txt", "w")
            f.write(",")
            f.close()
        f2 = open("important.txt", "r")
        self.credentials=[]
        for i in f2:
            self.credentials.append([i.split(",")[0], i.split(",")[1]])
        print(self.credentials)
        self.from_=self.credentials[0][0]
        self.pass_=self.credentials[0][1]

    def save_setting(self):
        if self.txt_from.get()=="" or self.txt_pass.get()=="":
            messagebox.showerror("Error", "All Fields are required", parent = self.root2)
        else: 
            f=open("important.txt", "w")
            f.write(self.txt_from.get()+","+ self.txt_pass.get())
            f.close()
            messagebox.showinfo("Success", "Saved Successfully", parent = self.root2)
            self.check_file_exist()
root=Tk()
obj=BULK_EMAIL(root)
root.mainloop()
