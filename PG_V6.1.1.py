#Importing necessary modules
from tkinter import *
from tkinter import messagebox, filedialog, ttk
from PIL import Image, ImageTk
from tkcalendar import Calendar
from cryptography.fernet import Fernet
from email.message import EmailMessage
import os, datetime, csv, pickle, webbrowser, win32com.client, smtplib, imghdr, random, string
current_version="6.1.1"
OneTimeUse="0"
bdw=1
cwd=os.getcwd()
def encrypt_message(strin,key):
    strin=str(strin)
    encoded_message = strin.encode()
    f = Fernet(key)
    return f.encrypt(encoded_message)
def decrypt_message(datakey,key):
    f = Fernet(key)
    decrypted_message = f.decrypt(datakey)
    return decrypted_message.decode()
# def troller(tkwin):
#     try:
#         troll=open(f"{cwd}\\files\\troll.dat","rb")
#         x=pickle.load(troll)
#         troll.close()
#     except:
#         x="DisablE"
#     if x.lower()=="disable":
#         pass
#     else:
#         curse=["arrow","circle","clock","cross","dotbox","exchange","fleur","heart","man","mouse","pirate","plus", "shuttle","sizing","spider","spraycan","star","target","tcross","trek","watch"]
#         tkwin.config(cursor=random.choice(curse))
def main():
    themeindex=0

    mainbg="#00a5db"
    subbg="#81e0ff"

    today=datetime.datetime.now().strftime("%d-%m-%Y")
    root=Tk() #Initialising GUI Application
    root.resizable(0,0)
    root.config(bg=mainbg)
    root.iconbitmap(cwd+"\\Template_&_Images\\icon.ico")
    root.title(f"Prescription Generator V{current_version}")
    root.state("zoomed")
    # troller(root)
    #Adding Scroll bar


    second_frame = Frame(root, borderwidth=0, bg=mainbg).pack()
    header=LabelFrame(second_frame,bg=mainbg, borderwidth=0)
    header.pack()
    greet=Label(header,font="Arial 20",text=f"Prescription Generator V{current_version}", bg=mainbg, fg="black").pack(pady=10)

    utilitiesF=LabelFrame(header,bg=mainbg, borderwidth=0)
    utilitiesF.pack()

    settingLF=LabelFrame(utilitiesF,bg=mainbg, borderwidth=bdw)

    megaF=LabelFrame(second_frame,bg=mainbg, borderwidth=0)
    megaF.pack()

    dateF=LabelFrame(utilitiesF, padx=10, pady=10, bg=mainbg, borderwidth=bdw)
    dateF.grid(row=0, column=1)

    date=Label(dateF,font="Arial 14",text="Date:", bg=mainbg).grid(row=0,column=0, sticky=E)
    edate=Entry(dateF, width=9,borderwidth=0, font="Arial 14")
    edate.insert(0,today)
    edate.grid(row=0, column=1, padx=5)

    def treeviewT_fun():
        global treeviewTK
        treeviewTK=Toplevel()
        # troller(treeviewTK)
        search_by=StringVar()
        search_value=StringVar()
        search_by_index=IntVar()
        treeviewTK.config(bg=subbg)
        treeviewTK.iconbitmap(cwd+"\\Template_&_Images\\icon.ico")
        #treeviewTK.grab_set()
    #Inserting Treeview Here
        style = ttk.Style()
        #Pick a theme
        style.theme_use()#$$$$$$$$$$$$ want to sprcify a theme??$$$$$$$$$$$$$$$
        # Configure our treeview colors
        style.configure("Treeview", 
            background="#D3D3D3",
            foreground="black",
            rowheight=25,
            fieldbackground="#D3D3D3"
            )
        # Change selected color
        style.map('Treeview', background=[('selected', 'blue')])

        search_frame=Frame(treeviewTK, bg=subbg, borderwidth=0, pady=10)
        search_frame.pack()

        # Create Treeview Frame
        tree_frame = Frame(treeviewTK, bg=subbg)
        tree_frame.pack(pady=5)
        # Treeview Scrollbar
        tree_scroll = Scrollbar(tree_frame)
        tree_scroll.pack(side=RIGHT, fill=Y)
        # Create Treeview
        my_tree = ttk.Treeview(tree_frame, yscrollcommand=tree_scroll.set, selectmode="extended")
        # Pack to the screen
        my_tree.pack(pady=5)
        #Configure the scrollbar
        tree_scroll.config(command=my_tree.yview)
        # Define Our Columns

        cols = ("Sno", "Name", "Patient ID","Age","Date","Gender","Contact","eMail")
        my_tree['columns']=cols
        # Formate Our Columns
        my_tree.column("#0", width=0, stretch=NO)
        my_tree.column("Sno", anchor=CENTER, width=40)
        my_tree.column("Name", anchor=W, width=120)
        my_tree.column("Patient ID", anchor=CENTER, width=80)
        my_tree.column("Age", anchor=CENTER, width=40)
        my_tree.column("Date", anchor=CENTER, width=75)
        my_tree.column("Gender", anchor=CENTER, width=50)
        my_tree.column("Contact", anchor=CENTER, width=100)
        my_tree.column("eMail", anchor=W, width=200)
        # Create Headings 
        my_tree.heading("#0", text="", anchor=CENTER)
        my_tree.heading("Sno", text="Sno", anchor=CENTER)
        my_tree.heading("Name", text="Name", anchor=CENTER)
        my_tree.heading("Patient ID", text="Patient ID",anchor=CENTER)
        my_tree.heading("Age", text="Age" ,anchor=CENTER)
        my_tree.heading("Date",text="Date" ,anchor=CENTER)
        my_tree.heading("Gender",text="Gender" ,anchor=CENTER)
        my_tree.heading("Contact",text="Contact" ,anchor=CENTER)
        my_tree.heading("eMail",text="eMail" ,anchor=CENTER)
        # Add Data
        file=open(f"{os.getcwd()}\\_data\\Patient_Data.csv","r", newline="")
        reader=csv.reader(file, delimiter=",")
        data=[]
        for i in reader:
            x=[i[0],i[1],i[2],i[3],i[4],i[5],i[6],i[7],i[8],i[9],i[10],i[11],i[12]]
            data.append(x)
        file.close()

        # Create striped row tags
        my_tree.tag_configure('oddrow', background="white")
        my_tree.tag_configure('evenrow', background="lightblue")
        global count
        count=0
        for record in data:
            if count % 2 == 0:
                my_tree.insert(parent='', index='end', iid=count, text="", values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],record[8], record[9], record[10],record[11], record[12]), tags=('evenrow',))
            else:
                my_tree.insert(parent='', index='end', iid=count, text="", values=(record[0], record[1], record[2], record[3], record[4], record[5], record[6], record[7],record[8], record[9], record[10],record[11], record[12]), tags=('oddrow',))
            count += 1


        def select_record():
            eb=[eage, egender, ename, ePID, econtact, eMail]
            tb=[ecomp, eallergy, ehistory, eDiag, eR]
            for i in eb:
                i.delete(0,END)
            for i in tb:
                i.delete(1.0,END)
            # Grab record number
            selected = my_tree.focus()
            # Grab record values
            values = my_tree.item(selected, 'values')
            #temp_label.config(text=values[0])
            slightyellow="#fffdd3"
            # output to entry boxes
            ename.insert(0, values[1])
            ePID.insert(0, values[2])
            eage.insert(0, values[3])
            egender.insert(0, values[5])
            econtact.insert(0, values[6])
            eMail.insert(0, values[7])
            ecomp.insert(1.0, values[8])
            eallergy.insert(1.0, values[9])
            ehistory.insert(1.0, values[10])
            eDiag.insert(1.0, values[11])
            eR.insert(1.0, values[12])
            all_ent_tb=[ename, ePID,eage, egender, econtact, eMail, ecomp, eallergy, ehistory, eDiag, eR]
            for i in all_ent_tb:
                i.config(bg=slightyellow)
        def clicker(e):
            select_record()
        my_tree.bind("<ButtonRelease-1>", clicker)
        treeviewTK.mainloop()
    databaseimg=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\database{themeindex}.png"))
    databasebutton=Button(dateF, image=databaseimg,bg=mainbg, borderwidth=0, command=treeviewT_fun)
    databasebutton.grid(row=0, column=3)

    settingIMG=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\Settings{themeindex}.png"))
    ResettIMG=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\reset{themeindex}.png"))

    def Settings():
        if settingLF.grid_info()=={}:
            settingLF.grid(row=0, column=0)
            settingLF.grid(row=0, column=0)
            def emailLoginTK():
                def openGoogleAppPass():
                    webbrowser.open(r"https://accounts.google.com/signin/v2/challenge/pwd?continue=https%3A%2F%2Fmyaccount.google.com%2Fapppasswords&service=accountsettings&osid=1&rart=ANgoxce3yqp489PQSC0hhItmUhQZuVOZTDXaFGz676-jaoY0fn532uN5Fk3zuj-N4HhNNMG1_e8XCPtlMNx-wL_Ocq0CQltkpg&TL=AM3QAYYPuf9De6BFjPVR2C_uLoGnXGMcqEvWYRxAOxvkAKqUZ5HaNnV9LVffcjiU&flowName=GlifWebSignIn&cid=1&flowEntry=ServiceLogin")
                def SaveNewD(x,y):
                    file2=open(cwd+f"\\files\\login.dat","wb")
                    details=f"{x}__$#@__{y}"
                    Ecryped_msg=encrypt_message(details, MainKey)
                    pickle.dump(Ecryped_msg,file2)
                    file2.close()
                LoginTK=Toplevel()
                LoginTK.resizable(0,0)
                # troller(LoginTK)
                LoginTK.config(bg=mainbg)
                LoginTK.iconbitmap(cwd+"\\Template_&_Images\\icon.ico")
                LoginTK.title(f"Prescription Generator V{current_version}")
                Centre(LoginTK,500,250)

                LoginTkMF=LabelFrame(LoginTK,text="Doctor's Email Details",font="Arial 14",borderwidth=0, bg=subbg, padx=10, pady=10)
                LoginTkMF.pack(pady=25)
                file=open(cwd+f"\\files\\login.dat","rb")
                login_cred=pickle.load(file)
                msg=decrypt_message(login_cred,MainKey)
                login_cred=msg.split("__$#@__")

                file.close()

                emaal=Label(LoginTkMF,font="Arial 12",text="Email:", bg=subbg).grid(row=0,column=0, sticky=E)
                eA=Entry(LoginTkMF, width=25,font="Arial 12",borderwidth=0)
                eA.insert(0,login_cred[0])
                eA.grid(row=0, column=1)

                passwd=Label(LoginTkMF,font="Arial 12",text="Password:", bg=subbg).grid(row=1,column=0, sticky=E, pady=10)
                eP=Entry(LoginTkMF, width=25,font="Arial 12", show="*",borderwidth=0)
                eP.insert(0,login_cred[1])
                eP.grid(row=1, column=1, sticky=W)

                AppPass=Button(LoginTkMF,font="Arial 10",text="Set Up App Password", bg="#49ff9a", command=openGoogleAppPass, borderwidth=0).grid(row=2, column=1)
                saveD=Button(LoginTkMF,font="Arial 10",text="Save Details", bg="yellow", command=lambda:[SaveNewD(eA.get(),eP.get()),LoginTK.destroy(),messagebox.showinfo(f"Prescription Generator V{current_version}","Login credentials saved successfully!")], borderwidth=0).grid(row=3, column=1, pady=10)

                LoginTK.mainloop()

            def save4everfun():
                global OneTimeUse
                fille=open(cwd+f"\\files\\email_data.dat","wb")
                sujeet,bodi=eA2.get(),actualbody.get(1.0,END)

                dat=[sujeet,bodi]
                pickle.dump(dat,fille)
                fille.close()
                OneTimeUse="0"
                eSubnBodTK.destroy()
                messagebox.showinfo(f"Prescription Generator V{current_version}","Email Data has been updated successfully!")

            def use_once_fn():
                global OneTimeUse, mail_inform
                OneTimeUse="1"

                mail_inform=[eA2.get(),actualbody.get(1.0,END)]
                eSubnBodTK.destroy()
                messagebox.showinfo(f"Prescription Generator V{current_version}","Current Email Data has been set for One time use only for the duration of this session!")
            def emailSubnBodyTK():
                global eA2, actualbody, eSubnBodTK
                eSubnBodTK=Toplevel()
                eSubnBodTK.resizable(0,0)
                # troller(eSubnBodTK)
                eSubnBodTK.config(bg=mainbg)
                eSubnBodTK.iconbitmap(cwd+"\\Template_&_Images\\icon.ico")
                eSubnBodTK.title(f"Prescription Generator V{current_version}")
                Centre(eSubnBodTK,650,615)
                try:
                    file99=open(cwd+f"\\files\\email_data.dat","rb")
                    dat=pickle.load(file99)
                    file99.close()
                except:
                    file01=open(cwd+f"\\files\\email_data.dat","wb")
                    dat=['Prescription of Consultation with doctor',"Dear {Mname},\nHope all your queries were resolved in your recent consultation with (Insert Doctor's Name) on: {Mdate}\nYour prescription is attached herewith. Wishing you a speedy recovery!\n\nThank You\n\nRegards\n(insert Doctor's Name)\n+91 (Add Phone Number)"]
                    pickle.dump(dat,file01)
                    file01.close()                 

                encompasser1=Frame(eSubnBodTK,borderwidth=0, bg=subbg, padx=10, pady=10)
                encompasser1.place(rely=0.05, relx=0.5, anchor=N)

                LoginTk2MF=Frame(encompasser1,borderwidth=0, bg=subbg, padx=10)

                greeting=Label(encompasser1,text="Email Subject and Body",font="Arial 20",bg=subbg,borderwidth=0).pack(pady=15)
                LoginTk2MF.pack()
                
                subl=Label(LoginTk2MF,font="Arial 12",text="Subject:", bg=subbg).grid(row=0,column=0, sticky=W)
                eA2=Entry(LoginTk2MF, width=54,font="Arial 12",borderwidth=0)
                eA2.insert(0,dat[0])
                eA2.grid(row=0, column=1, pady=10, sticky=W)

                bodl=Label(LoginTk2MF,font="Arial 12",text="Body:", bg=subbg).grid(row=1, column=0, sticky=W)
                actualbody=Text(encompasser1, width=60, height=20,font="Arial 12", borderwidth=0, wrap=WORD)
                actualbody.pack()
                actualbody.insert(1.0,dat[1])

                buttonframe=Frame(encompasser1,borderwidth=0, bg=subbg, padx=10)
                buttonframe.pack()

                use_once=Button(buttonframe,font="Arial 12",text="Just Once", bg="cyan", command=use_once_fn, borderwidth=0).grid(row=0, column=0, pady=10, sticky=E)
                save_permanent=Button(buttonframe,font="Arial 12",text="Always", bg="orange", command=save4everfun, borderwidth=0).grid(row=0, column=1,padx=10, sticky=E)
                
                eSubnBodTK.mainloop()
            def contDevTK():
                pass
            BemailAccount=Button(settingLF,font="Arial 12",text="Your Email Account", bg="#49ff9a", command=emailLoginTK, borderwidth=0).grid(row=0, column=0,padx=10)
            BemailData=Button(settingLF,font="Arial 12",text="Edit Email Data", bg="#49ff9a", command=emailSubnBodyTK, borderwidth=0).grid(row=0, column=1)
            Bcontactdev=Button(settingLF,font="Arial 12",text="Change Colour Theme", bg="orange", command=contDevTK, borderwidth=0).grid(row=0, column=2,padx=10)
        else:
            settingLF.grid_forget()


    SetButt=Button(dateF, image=settingIMG, command=Settings, borderwidth=0,bg=mainbg).grid(row=0, column=2,padx=10, sticky=E)

    my_tree_frame=LabelFrame(header, padx=10, pady=10, bg=subbg, borderwidth=0)
    my_tree_frame.pack()

    mega2F=Frame(megaF,bg=subbg)
    mega2F.pack()

    mainL=Frame(mega2F,borderwidth=bdw, padx=10, pady=10, bg=subbg)
    mainL.grid(row=0, column=0)

    calFrame=Frame(mainL,borderwidth=bdw, bg=subbg)

    cal=Calendar(calFrame,
    date_pattern="dd-mm-y",
    selectmode="day",
    firstweekday="monday",
    background="#0f8bff",
    foreground="black",
    disabledforeground="#a3bec7",
    bordercolor="grey",
    normalbackground="#d0f4ff",
    weekendbackground="#8ffeff",
    weekendforeground ="black" ,
    disabledbackground="99b3bc")

    calFrame.grid(row=2, column=0,pady=20, sticky=NSEW) #old pady=35
    cal.pack(pady=10,padx=90, fill="both", expand=True)
    def getCustomDate():
        edate.delete(0,END)
        edate.insert(0,cal.get_date())

    gdimg=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\GTD{themeindex}.png"))
    getdat=Button(calFrame,image=gdimg,bg=subbg, command=getCustomDate, borderwidth=0)
    getdat.pack()

    main2L=LabelFrame(mega2F,borderwidth=0, padx=10, pady=10, bg=subbg)
    main2L.grid(row=0, column=1,pady=20, padx=30)

    PDlf=LabelFrame(mainL, padx=10, pady=13, bg=subbg, borderwidth=bdw)
    PDlf.grid(row=0, column=0, sticky=NSEW)

    MaleFrame=LabelFrame(mainL, padx=50, pady=10, bg=subbg, borderwidth=bdw)
    MaleFrame.grid(row=1, column=0, sticky=NSEW)
    Mail=Label(MaleFrame, text="Email:", bg=subbg, font="Arial 14").grid(row=0,column=0, sticky=E)
    eMail=Entry(MaleFrame,font="Arial 14", width=25,borderwidth=0)
    eMail.grid(row=0,column=1)

    sno=Label(PDlf,font="Arial 14",text="Sno:", bg=subbg).grid(row=0,column=0, sticky=E)
    eSNO=Entry(PDlf, width=3,font="Arial 14",borderwidth=0)
    ##########################
    eSNO.grid(row=0, column=1)

    age=Label(PDlf,font="Arial 14",text="Age:", bg=subbg).grid(row=1,column=0, sticky=E)
    eage=Entry(PDlf, width=3,font="Arial 14",borderwidth=0)
    eage.grid(row=1, column=1, pady=10)

    gender=Label(PDlf,font="Arial 14",text="Gender:", bg=subbg).grid(row=2,column=0, sticky=E)
    egender=Entry(PDlf, width=3,font="Arial 14",borderwidth=0)
    egender.grid(row=2, column=1)

    name=Label(PDlf,font="Arial 14",text="Name:", bg=subbg).grid(row=0,column=3, sticky=E)
    ename=Entry(PDlf, width=15,font="Arial 14",borderwidth=0)
    ename.grid(row=0, column=4)

    Pid=Label(PDlf,font="Arial 14",text="PatientID:", bg=subbg).grid(row=1,column=3, sticky=E)
    ePID=Entry(PDlf, width=15,font="Arial 14",borderwidth=0)
    ePID.grid(row=1, column=4, pady=10)

    contact=Label(PDlf,font="Arial 14",text="Contact:", bg=subbg).grid(row=2,column=3, sticky=E)
    econtact=Entry(PDlf, width=15,font="Arial 14",borderwidth=0)
    econtact.grid(row=2, column=4)

    blank=Label(PDlf,font="Arial 14",text="         ", bg=subbg).grid(row=1,column=2)

    contentLF=LabelFrame(main2L, padx=10, bg=subbg, borderwidth=bdw) #, pady=10
    contentLF.grid(row=0, column=2, sticky=SE)

    blank2=Label(PDlf,font="Arial 14",text="         ", bg=subbg).grid(row=0,column=2)

    complaint=Label(contentLF,font="Arial 14",text="Complaint:", bg=subbg).grid(row=0,column=1, sticky=E)
    ecomp=Text(contentLF, font="Arial 10", width=78, height=3,borderwidth=0)
    ecomp.grid(row=0, column=2, pady=10)

    allergy=Label(contentLF,font="Arial 14",text="Allergy:", bg=subbg).grid(row=1,column=1, sticky=E)
    eallergy=Text(contentLF, font="Arial 10", width=69, height=1,borderwidth=0)
    eallergy.grid(row=1, column=2)

    history=Label(contentLF,font="Arial 14",text="History:", bg=subbg).grid(row=2,column=1, sticky=E)
    ehistory=Text(contentLF, font="Arial 10", width=73, height=3,borderwidth=0)
    ehistory.grid(row=2, column=2, pady=10)

    Diag=Label(contentLF,font="Arial 14",text="Diagnosis:", bg=subbg).grid(row=3,column=1, sticky=E)
    eDiag=Text(contentLF, font="Arial 10", width=70, height=1,borderwidth=0)
    eDiag.grid(row=3, column=2)

    R=Label(contentLF,font="Arial 20",text="Rx:", bg=subbg).grid(row=4,column=1, sticky=E)
    eR=Text(contentLF, font="Arial 10", width=80, height=17,borderwidth=0)
    eR.grid(row=4, column=2, pady=10)

    buttF=LabelFrame(main2L, bg=subbg, borderwidth=bdw)
    buttF.grid(row=4, column=2, sticky=E)

    ###############################$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$$treeview was here
    genPDF= IntVar()
    mailSender= IntVar()
    WAsender= IntVar()

    chkmail=Checkbutton(MaleFrame,font="Arial 12",text="Send Email", bg=subbg, variable=mailSender)
    chkmail.grid(row=1, column=1, pady=5)
    chkpdf=Checkbutton(buttF,font="Arial 12",text="Export PDF", bg=subbg, variable=genPDF)
    chkpdf.grid(row=0, column=1)
    chkmail.select()
    chkpdf.select()
    chkWA=Checkbutton(PDlf,font="Arial 12",text="Send WhatsApp", bg=subbg, variable=WAsender)
    chkWA.grid(row=3, column=4)
    chkWA.deselect()

    def Reload():
        eb=[edate,eage, egender, ename, ePID, econtact, eMail, eSNO]
        tb=[ecomp, eallergy, ehistory, eDiag, eR]
        for i in eb:
            i.delete(0,END)
            i.config(bg="white")
        for i in tb:
            i.delete(1.0,END)
            i.config(bg="white")
        eSNO.insert(0,SnoWriter())
        edate.insert(0,today)
        try:
            button_proceed.grid_forget()
        except:
            pass
        try:
            settingLF.grid_forget()
        except:
            pass
        try:
            notif_WA.grid_forget()
        except:
            pass
        try:
            notif_PDF.grid_forget()
        except:
            pass
        try:
            notif_mail.grid_forget()
        except:
            pass
        try:
            treeviewTK.destroy()
        except:
            pass
        # troller(root)

    ResetButt=Button(dateF,image=ResettIMG, command=Reload, borderwidth=0, bg=mainbg).grid(row=0, column=4, padx=10)

    try:
        keyfile=open(f"{cwd}\\files\\##00..)(..00##.key","rb")
        MainKey=pickle.load(keyfile)

    except:
        keyfile=open(f"{cwd}\\files\\##00..)(..00##.key","wb")
        MainKey=Fernet.generate_key()
        initialdata="Your_Email@gmail.com__$#@__"

        deven=encrypt_message(initialdata,MainKey)

        loginF=open(f"{cwd}\\files\\login.dat","wb")
        pickle.dump(deven,loginF)
        loginF.close()
        pickle.dump(MainKey,keyfile)

    keyfile.close()
    def SnoWriter():
        sno_reader=open(f"{cwd}\\_data\\Patient_Data.csv","r", newline="")
        reader=csv.reader(sno_reader, delimiter=",")
        SnoL=[i[0] for i in reader]
        sno_reader.close()
        try:
            snodata=str(1+(int(SnoL[-1])))
        except ValueError:
            snodata=1
        return snodata
    eSNO.insert(0,SnoWriter())

    def Centre(NameOfTkinterWindow,Width,Height):#This function Centres the Tkinter window on the screen
        scrwdth = NameOfTkinterWindow.winfo_screenwidth()
        scrhgt = NameOfTkinterWindow.winfo_screenheight()
        xLeft = (int(scrwdth)//2 - (int(Width)//2))
        yTop = (int(scrhgt)//2 - (int(Height)//2))
        NameOfTkinterWindow.geometry(str(Width) + "x" + str(Height) + "+" + str(xLeft) + "+" + str(yTop))
    
    buttproceedimg=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\buttproced{themeindex}.png"))
    button_proceed=Button(dateF, image=buttproceedimg, borderwidth=0, bg=mainbg)

    #Defining All Functions used in this project
    def AddData2CSV():
        e=edate.get()
        a=eSNO.get()
        d=eage.get()
        f=egender.get()
        b=ename.get()
        c=ePID.get()
        g=econtact.get()
        h=ecomp.get(1.0,END)
        i=eallergy.get(1.0,END)
        j=ehistory.get(1.0,END)
        k=eR.get(1.0,END)
        m=eMail.get()
        diag=eDiag.get(1.0, END)
        data=[a,b,c,d,e,f,g,m,h,i,j,diag,k]

        file=open(f"{cwd}\\_data\\Patient_Data.csv","a", newline="")
        writer=csv.writer(file, delimiter=",")
        writer.writerow(data)
        file.close()

        savefileas=f"{c}-{b}_({e})".replace(" ","")

        messagebox.showinfo(f"Prescription Generator V{current_version}", "Data has been saved to the database successfully!")
        
        objShell = win32com.client.Dispatch("WScript.Shell")
        UserDocs = objShell.SpecialFolders("MyDocuments")
        ExpName=filedialog.asksaveasfilename(initialdir=UserDocs,initialfile=str(savefileas+".png"), title="Save Prescription As", defaultextension=".png",filetypes=[("PNG Files","*.png")])
        ExpName=ExpName.replace("/","\\")

        psApp = win32com.client.Dispatch("Photoshop.Application")
        psApp.Open(f"{cwd}\\Template_&_Images\\Presc_Template.psd")
        doc = psApp.Application.ActiveDocument

        lf1 = doc.ArtLayers["name"]
        tol1 = lf1.TextItem
        tol1.contents = ename.get()

        lf2 = doc.ArtLayers["age"]
        tol2 = lf2.TextItem
        tol2.contents = eage.get()

        lf3 = doc.ArtLayers["gender"]
        tol3 = lf3.TextItem
        tol3.contents = egender.get()

        lf4 = doc.ArtLayers["pid"]
        tol4 = lf4.TextItem
        tol4.contents = ePID.get()

        lf4 = doc.ArtLayers["date"]
        tol4 = lf4.TextItem
        tol4.contents = edate.get()

        lf5 = doc.ArtLayers["contact"]
        tol5 = lf5.TextItem
        tol5.contents = econtact.get()

        lf6 = doc.ArtLayers["complaint"]
        tol6 = lf6.TextItem
        h=ecomp.get(1.0,END)
        h=h.rstrip("\n")
        h=h.replace("\n","\r")
        varH="                        "+h
        tol6.contents =varH

        lf7 = doc.ArtLayers["allergy"]
        tol7 = lf7.TextItem
        i=eallergy.get(1.0,END)
        i=i.rstrip("\n")
        i=i.replace("\n","\r")
        tol7.contents = i

        lf8 = doc.ArtLayers["history"]
        tol8 = lf8.TextItem
        j=ehistory.get(1.0,END)
        j=j.rstrip("\n")
        j=j.replace("\n","\r")
        varJ="                                               "+j
        tol8.contents =varJ

        lf9 = doc.ArtLayers["R"]
        tol9 = lf9.TextItem
        k=eR.get(1.0,END)
        k=k.rstrip("\n")
        k=k.replace("\n","\r")
        tol9.contents = k

        lf10 = doc.ArtLayers["diag"]
        tol10 = lf10.TextItem
        diag=eDiag.get(1.0, END)
        diag=diag.rstrip("\n")
        diag=diag.replace("\n","\r")
        tol10.contents = diag

        options = win32com.client.Dispatch('Photoshop.ExportOptionsSaveForWeb')
        options.Format = 13
        options.PNG8 = False

        to_be_removed=ExpName.split("\\")[-1]
        ExpDir=ExpName.split("\\"+to_be_removed)[0]

        if ExpName!="":
            pngfile=ExpName

        elif ExpName=="":
            UserDocs.replace("/","\\")
            try:
                ExpDir=f"{UserDocs}\\Prescription Generator\\Generated Prescriptions"
            except:
                ExpDir=f"{cwd}\\Generated Prescriptions"
            pngfile=ExpDir+"\\"+savefileas+".png"
            
        def taskdoer():
            doc.Export(ExportIn=pngfile, ExportAs=2, Options=options)
            messagebox.showinfo(f"Prescription Generator V{current_version}", f"Prescription has been saved in the location: {ExpDir} successfully!")
        def cont():
            global notif_PDF
            def SendWA():
                global notif_WA, PDFname
                pn="91"+(econtact.get().replace(" ",""))
                link=f"https://web.whatsapp.com/send?phone={pn}&text&app_absent=0"
                try:
                    webbrowser.open(link)
                    notif_WA=Label(PDlf, text="Success!", bg=subbg, font="Arial 12", fg="green")
                    notif_WA.grid(row=3,column=3, padx=5)
                except:
                    notif_WA=Label(PDlf, text="ERROR", bg=subbg, font="Arial 12", fg="red")
                    notif_WA.grid(row=3,column=3, padx=5)
            if genPDF.get()==1:
                try:
                    imge=Image.open(pngfile)
                    im1 = imge.convert('RGB')
                    pd_name=ExpName.split(".png")
                    PDFname=(pd_name[0]+".pdf")
                    im1.save(PDFname)

                    notif_PDF=Label(buttF, text="Success!", fg="green", bg=subbg, font="Arial 12")
                    notif_PDF.grid(row=0, column=0, padx=10, sticky=E)
                except:
                    notif_PDF=Label(buttF, text="ERROR", fg="red", bg=subbg, font="Arial 12")
                    notif_PDF.grid(row=0, column=0, padx=10, sticky=E)
            else:
                pass

            def SendEmail(att):
                global notif_mail
                with open (att, "rb") as f:
                    file_data=f.read()
                    file_name=f.name
                    file_name=file_name.split("\\")

                att=att.split(".")

                if att[-1].lower()=="pdf":
                    main_type="application"
                    sub_type="octet-stream"
                else:
                    main_type="image"
                    sub_type=imghdr.what(f.name)

                file3=open(cwd+f"\\files\\login.dat","rb")
                login_cred=pickle.load(file3)
                logincredentials=decrypt_message(login_cred,MainKey)
                login_cred=logincredentials.split("__$#@__")
                file3.close()

                email_user = login_cred[0]
                email_password = login_cred[-1]
                email_send = m

                if OneTimeUse=="1":
                    mail_subject=mail_inform[0]
                    mail_body=mail_inform[1]                  
                else:
                    email_data=open(f"{cwd}\\files\\email_data.dat","rb")
                    sub_bdy=pickle.load(email_data)
                    email_data.close()

                    mail_subject,mail_body=str(sub_bdy[0]),str(sub_bdy[1])

                mail_subject=mail_subject.format(Mname=b,Mdate=e,Mage=d,Mgender=f,MPID=c,Mallergy=i,Mhistory=j,MR=k)
                mail_body=mail_body.format(Mname=b,Mdate=e,Mage=d,Mgender=f,MPID=c,Mallergy=i,Mhistory=j,MR=k)

                msg=EmailMessage()
                msg['Subject']=mail_subject
                msg['From']=email_user
                msg['To']=email_send
                msg.set_content(mail_body)

                msg.add_attachment(file_data, maintype=main_type, subtype=sub_type, filename=file_name[-1])

                try:
                    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
                        smtp.login(email_user, email_password)
                        smtp.send_message(msg)
                    notif_mail=Label(MaleFrame, text="Success!", bg=subbg, font="Arial 12", fg="green")
                    notif_mail.grid(row=1,column=0, sticky=E, pady=10)
                except:
                    notif_mail=Label(MaleFrame, text="ERROR", bg=subbg, font="Arial 12", fg="red")
                    notif_mail.grid(row=1,column=0, sticky=E, pady=10)
            


            if mailSender.get()==1 and m!="" and genPDF.get()==1:
                SendEmail(PDFname)
            elif mailSender.get()==1 and m=="":
                notif=Label(MaleFrame, text="Please Enter a Valid Email address!", bg=subbg, font="Arial 12", fg="red").grid(row=2,column=1, sticky=NSEW, pady=10)
            elif mailSender.get()==1 and m!="" and genPDF.get()==0:
                SendEmail(ExpName)
            if WAsender.get()==1 and g!="":
                SendWA()
        def lol():
            taskdoer()
            cont()
        ans=messagebox.askquestion(f"Prescription Generator V{current_version}","Are you happy with the PSD document?")
        if ans=="yes":
            taskdoer()
            cont()
        elif ans=="no":
            button_proceed.config(command=lol)
            button_proceed.grid(row=0, column=5)
    CSVadderIMG=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\gen_P_csv_add{themeindex}.png"))
    csvAdd=Button(buttF, image=CSVadderIMG, command=AddData2CSV, borderwidth=0,bg=subbg).grid(row=0, column=2,padx=10, sticky=E)
    root.mainloop()
def send_PA_mail(suject,body):
    my_mail_ID="prescr.generator@gmail.com"
    my_mail_Pass="bgmufhzwcegsabjo"
    msg=EmailMessage()
    msg['Subject']=suject
    msg['From']=my_mail_ID
    msg['To']="devenjain2020@gmail.com"
    msg.set_content(body)
    try:
        with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
            smtp.login(my_mail_ID, my_mail_Pass)
            smtp.send_message(msg)
    except:
        messagebox.showerror("Product Activation Failed!","Please establish a good internet connection and try again later!")
def keygen():
    def stringpart(len):
        deven = ''.join(random.choice(string.ascii_uppercase) for i in range(len))
        return deven
    def numpart(len):
        deven2=''.join(random.choice(string.digits) for i in range(len))
        return deven2
    fkey=f"{stringpart(3)}{numpart(2)}-{stringpart(2)}{numpart(1)}{stringpart(2)}-{numpart(2)}{stringpart(3)}-{stringpart(4)}{numpart(1)}-{stringpart(3)}{numpart(2)}"
    return fkey
def checker():
    global e_prod_key, thekey
    new_string=e_prod_key.get()
    new_string=new_string.replace("-","")
    thekey=thekey.replace("-","")
    if new_string==thekey:
        stat_file=open(f"{cwd}\\files\\misc.dat","wb")
        to_be_written=encrypt_message("1",prod_encry_key)
        pickle.dump(to_be_written,stat_file)
        stat_file.close()

        prod.destroy()

        f=open(f"{cwd}\\_buttons_\\_button_index.dat","wb")
        pickle.dump(encrypt_message(cwd,prod_encry_key),f)
        #Loading CWD at the time of creation into the software
        f.close()
        objShell = win32com.client.Dispatch("WScript.Shell")
        UserDocs = objShell.SpecialFolders("MyDocuments")
        objShell = win32com.client.Dispatch("WScript.Shell")
        try:
            os.makedirs(f"{UserDocs}\Prescription Generator\Program Files")
            os.mkdir(f"{UserDocs}\Prescription Generator\Generated Prescriptions")
        except:
            pass
        main()
    elif new_string!=thekey:
        messagebox.showerror("Product Activation Failed!","You have entered an invalid Product key. Please try again!")
def ProdActive():
    global emailofuser, e_prod_key, prod, thekey
    if emailofuser.get()!="" and eusern.get()!="":
        thekey=keygen()
        subject1="Someone tried to activate your software!"
        body1=f"Their name is: {eusern.get()}\nTheir email ID is: {emailofuser.get()}\n\nTheir unique Product Key is: {thekey}\n\nIf you know who has made this request then it is okay. Otherwise, malicious people might be trying to illegally distribute your software!\nBe careful"
        send_PA_mail(subject1, body1)
        mainwin.destroy()
        prod=Tk()
        prod.config(bg="#007f6c")
        prod.title("Product Activation Page")
        prod.attributes("-fullscreen",True)
        prod.iconbitmap=(cwd+"\\Template_&_Images\\icon.ico")
        prod.wm_attributes("-transparentcolor","#007f6c")

        main2_frame=Frame(prod,borderwidth=0, bg="white", padx=30, pady=20)
        main2_frame.place(relx=0.5, rely=0.5, anchor=CENTER)

        Greeting2_Label=Label(main2_frame, text="Product Activation",bg="white", font="Arial 20")
        Greeting2_Label.pack()

        details2_frame=Frame(main2_frame,borderwidth=0,bg="white",padx=10, pady=20)
        details2_frame.pack(pady=10)

        entprodkey=Label(details2_frame, text="Enter your unique product key", bg="white", font="Arial, 13").grid(row=0, column=0, sticky=W)
        e_prod_key=Entry(details2_frame,bg="white", show="*", font="Arial 15", width=35)
        e_prod_key.grid(row=1, column=0, pady=10)

        prodimg=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\prod_act_button.png"))
        prodbutton=Button(main2_frame, image=prodimg, bg="white", borderwidth=0, command=checker)
        prodbutton.pack()
        prod.mainloop()
    elif emailofuser.get()=="" or eusern.get()=="":
        messagebox.showerror("Product Activation","Invalid Name or Email entered!\nPlease try again!")
def main2():
    global emailofuser, mainwin, eusern
    mainwin=Tk()
    mainwin.config(bg="#007f6c")
    mainwin.title("Set-Up Page")
    mainwin.attributes("-fullscreen",True)
    mainwin.iconbitmap=(cwd+"\\Template_&_Images\\icon.ico")
    mainwin.wm_attributes("-transparentcolor","#007f6c")

    main_frame=Frame(mainwin,borderwidth=0, bg="white", padx=30, pady=20)
    main_frame.place(relx=0.5, rely=0.5, anchor=CENTER)

    Greeting_Label=Label(main_frame, text="Welcome to Prescription Generator!",bg="white", font="Arial 20")
    Greeting_Label.pack()

    details_frame=Frame(main_frame,borderwidth=0,bg="white",padx=10, pady=20)
    details_frame.pack(pady=10)

    username=Label(details_frame, text="Enter your Name:", bg="white", font="Arial, 15").grid(row=0, column=0,sticky=E)
    eusern=Entry(details_frame,bg="white", font="Arial 15", width=25)
    eusern.grid(row=0, column=1, pady=15)

    emailadd=Label(details_frame, text="Enter your Email:", bg="white", font="Arial, 15").grid(row=1, column=0,sticky=E)
    emailofuser=Entry(details_frame,bg="white", font="Arial 15", width=25)
    emailofuser.grid(row=1, column=1, pady=5)

    contimage=ImageTk.PhotoImage(Image.open(f"{cwd}\\_buttons_\\cont_button.png"))
    continuebutton=Button(main_frame, image=contimage,bg="white", borderwidth=0, command=lambda:[ProdActive()])
    continuebutton.pack()

    mainwin.mainloop()
def testing():
    global prod_encry_key
    prod_encry_key=b'sAOeNc_T9_j1Ny_DndLDVbzzmf0pHM-sr3ySPyGVfhY='
    try:
        status_file=open(f"{cwd}\\files\\misc.dat","rb")
        status=pickle.load(status_file)
        status_file.close()
        yon=decrypt_message(status,prod_encry_key)
    except:
        status_file=open(f"{cwd}\\files\\misc.dat","wb")
        entry=encrypt_message("0",prod_encry_key)
        pickle.dump(entry,status_file)
        status_file.close()
        yon="0"
    if yon=="0":
        main2()
    elif yon=="1":
        try:
            second_layer_f=open(f"{cwd}\\_buttons_\\_button_index.dat","rb")
            data_of_f=pickle.load(second_layer_f)
            first_time_cwd=decrypt_message(data_of_f,prod_encry_key)
            second_layer_f.close()
            if first_time_cwd!=cwd:

                key2file=open(f"{cwd}\\files\\##00..)(..00##.key","rb")
                email_data_key=pickle.load(key2file)
                key2file.close()

                file_new=open(cwd+f"\\files\\login.dat","rb")
                login_cred=pickle.load(file_new)
                logincredentials=decrypt_message(login_cred,email_data_key)
                login_cred=logincredentials.split("__$#@__")
                file_new.close()

                subject2="Change in CWD detected!"
                body2=f"A change in CWD of your client has been detected. Their software has been deactivated successfully!\n\nCWD at the time of Product Actication: {first_time_cwd}\n\nNew CWD: {cwd}\n\nEmail ID of original user: {login_cred[0]}\n\nName of unauthorised new User: {os.environ['USERNAME']}\n\nThese people might be responsible for the redestribution of your software! Get in touch with them to stop this unjust practice!"
                send_PA_mail(subject2,body2)
                os.remove(f"{cwd}\\files\\misc.dat")
                messagebox.showerror(f"Prescription Generator V{current_version}","A fatal error has occured! Please restart the software!")
            else:
                main()
        except:
            os.remove(f"{cwd}\\files\\misc.dat")
testing()