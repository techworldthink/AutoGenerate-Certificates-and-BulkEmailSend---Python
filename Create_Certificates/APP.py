import pandas as pd
from PIL import Image, ImageDraw, ImageFont,ImageTk
import PIL.Image
import csv
import os
import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
from tkinter import*
from tkinter import font
from tkinter import messagebox
from tkinter.scrolledtext import ScrolledText
from tkinter.filedialog import askopenfilename
from tkinter.messagebox import showinfo

class Convert:
    def __init__(self, master):
        self.master = master
        master.title("Certificate Generate")
        master.geometry("900x600")
        master.configure(bg="white")

        ButtonVar1 = StringVar()
        ButtonVar2 = StringVar()
        ButtonVar3 = StringVar()

        menubar = Menu(master)
        filemenu = Menu(menubar, tearoff=0)
        filemenu.add_command(label="New")
        filemenu.add_command(label="Open",command=lambda:choose_excel())
        filemenu.add_command(label="Save")
        filemenu.add_command(label="Save as...")
        filemenu.add_command(label="Close")
        filemenu.add_separator()
        filemenu.add_command(label="Exit", command=master.quit)
        menubar.add_cascade(label="File", menu=filemenu)
        
        editmenu = Menu(menubar, tearoff=0)
        editmenu.add_command(label="Undo")
        editmenu.add_separator()
        editmenu.add_command(label="Cut")
        editmenu.add_command(label="Copy")
        editmenu.add_command(label="Paste")
        editmenu.add_command(label="Delete")
        editmenu.add_command(label="Select All")
        menubar.add_cascade(label="Edit", menu=editmenu)
        
        helpmenu = Menu(menubar, tearoff=0)
        helpmenu.add_command(label="Help Index")
        helpmenu.add_command(label="About...")
        menubar.add_cascade(label="Help", menu=helpmenu)
        master.config(menu=menubar)
        
        #full window row configure
        master.grid_rowconfigure(0, weight=1)
        master.grid_rowconfigure(1, weight=1)
        #full window column configure
        master.columnconfigure(0, weight=1)
        master.columnconfigure(1, weight=1)
      
        #Fonts
        self.label_frame_font = font.Font(family="Helvetica",size=10,weight="bold")
        self.frame2_font = font.Font(family="Franklin Gothic Medium",size=10)

        # text color of name 
        self.text_color = (0, 0, 0)
        # font of name
        self.font = ImageFont.truetype("Fonts/Oswald-Bold.ttf", 50)
        

        #SET INDICATORS PROPERTIES
        WIDTH_BTN = 10
        WIDTH_LABEL = 5

        # Demo model certificate
        model = PIL.Image.open("Demo.jpg")
        newsize = (300, 150)
        model = model.resize(newsize)
        IMG_CERT = ImageTk.PhotoImage(model)

        # BG IMAGE

        bgmodel = PIL.Image.open("Open.jpg")
        newsize = (100, 30)
        bgmodel = bgmodel.resize(newsize)
        F_IMG_CERT = ImageTk.PhotoImage(bgmodel)

        bgmodel = PIL.Image.open("Generate.jpg")
        newsize = (100, 30)
        bgmodel = bgmodel.resize(newsize)
        S_IMG_CERT = ImageTk.PhotoImage(bgmodel)

        bgmodel = PIL.Image.open("send.jpg")
        newsize = (100, 30)
        bgmodel = bgmodel.resize(newsize)
        T_IMG_CERT = ImageTk.PhotoImage(bgmodel)

        self.LOG="Starting..."

        
        #labelled frames
        self.frame_left  =  LabelFrame(master,text="SELECT EXCEL FILE",labelanchor="n",bg="white",bd=5,fg="red",font=self.label_frame_font)
        self.frame_right  =  LabelFrame(master,text="GENERATE-CERTIFICATES",labelanchor="n",bg="white",bd=5,fg="red",font=self.label_frame_font)
        self.BTM = ScrolledText(master,height=1,width=5,bg="white",fg="green")
        self.BTM.insert(tk.INSERT,"-------------LOGS----------\nStarting...")
        
        #frame grids
        self.frame_left.grid(row=0,column=0,sticky="nsew")
        self.frame_right.grid(row=0,column=1,sticky="nsew")
        self.BTM.grid(row=1,column=0,sticky="nsew",columnspan=2)
   
        # -------------------------------- LEFT -------------------------------- #
        
        # LEFT frame
        self.frame_left.grid_rowconfigure(0, weight=1)
        self.frame_left.grid_rowconfigure(1, weight=1)
        self.frame_left.grid_rowconfigure(2, weight=1)
        self.frame_left.grid_rowconfigure(3, weight=1)
        #self.frame_left.grid_rowconfigure(4, weight=1)
        
        self.frame_left.columnconfigure(0, weight=1)
        self.frame_left.columnconfigure(1, weight=1)
        
       

        #componants for frame 1
        self.LABEL_LEFT = Label(self.frame_left,text="DETAILS OF PARTICIPANTS",padx=0,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        self.ENTRY_LEFT = Entry(self.frame_left,bg="white",fg="green",textvariable = ButtonVar1,bd=0)
        self.BOX_LEFT_1 = ScrolledText(self.frame_left,width=5)
        self.BOX_LEFT_2 = ScrolledText(self.frame_left,width=5) 
        self.BTN_LEFT   = Button(self.frame_left,text="",image=F_IMG_CERT,height = 20, width = WIDTH_BTN,bg="white",bd=0,fg="black",command=lambda:choose_excel())
        self.BTN_LEFT.image=F_IMG_CERT
        
        #componants grid for frame 1
        self.LABEL_LEFT.grid(row=0,column=0,sticky="nsew",columnspan=2)
        self.BOX_LEFT_1.grid(row=1,column=0,sticky="nsew")
        self.BOX_LEFT_2.grid(row=1,column=1,sticky="nsew")
        self.BTN_LEFT.grid(row=2,column=0,sticky="nsew",columnspan=2)
        self.ENTRY_LEFT.grid(row=3,column=0,sticky="ew",columnspan=2)
     
        # -------------------------------- RIGHT -------------------------------- #
        
        # LEFT frame
        self.frame_right.grid_rowconfigure(0, weight=1)
        self.frame_right.grid_rowconfigure(1, weight=1)
        self.frame_right.grid_rowconfigure(2, weight=1)
        self.frame_right.grid_rowconfigure(3, weight=1)
        self.frame_right.grid_rowconfigure(4, weight=1)
        
        self.frame_right.columnconfigure(0, weight=1)
        self.frame_right.columnconfigure(1, weight=1)
        self.frame_right.columnconfigure(2, weight=1)

        #componants for frame 2
        self.LABEL_RIGHT= Label(self.frame_right,text="DETAILS OF CERTIFICATE",padx=0,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        self.LABEL_RIGHT_2= Label(self.frame_right,image=IMG_CERT,text="DETAILS OF CERTIFICATE",padx=0,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        self.BTN_RIGHT_1 = Button(self.frame_right,text="",image=F_IMG_CERT,height = 50, width = WIDTH_BTN,bg="white",bd=0,fg="black",command=lambda:choose_image())
        self.LABEL_RIGHT_2.image=IMG_CERT
        self.BTN_RIGHT_1.image=F_IMG_CERT
        
        self.LABEL_RIGHT_IM = Label(self.frame_right,text="Size",padx=0,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        self.ENTRY_LEFT_IMW = Entry(self.frame_right,bg="white",fg="black",textvariable = ButtonVar2,bd=5)
        self.ENTRY_LEFT_IMH = Entry(self.frame_right,bg="white",fg="black",textvariable = ButtonVar3,bd=5)

        self.BTN_RIGHT_GNT = Button(self.frame_right,text="GENERATE",image=S_IMG_CERT,height = 40, width = WIDTH_BTN,bg="white",bd=0,fg="black",command=lambda:generate())
        self.BTN_RIGHT_GNT.image=S_IMG_CERT
        self.BTN_RIGHT_GNT_2 = Button(self.frame_right,text="GENERATE",image=T_IMG_CERT,height = 40, width = WIDTH_BTN,bg="white",bd=0,fg="black",command=lambda:send_to_emails())
        self.BTN_RIGHT_GNT_2.image=T_IMG_CERT
        
        #componants grid for frame 1
        self.LABEL_RIGHT.grid(row=0,column=0,sticky="nsew",columnspan=3)
        self.LABEL_RIGHT_2.grid(row=1,column=0,sticky="nsew",columnspan=3)
        self.BTN_RIGHT_1.grid(row=2,column=0,sticky="ew",columnspan=3)
        self.LABEL_RIGHT_IM.grid(row=3,column=0,sticky="ew")
        self.ENTRY_LEFT_IMW.grid(row=3,column=1,sticky="ew")
        self.ENTRY_LEFT_IMH.grid(row=3,column=2,sticky="ew")
        self.BTN_RIGHT_GNT.grid(row=4,column=1,sticky="ew")
        self.BTN_RIGHT_GNT_2.grid(row=4,column=2,sticky="ew")

        
        # -----------------------   define functions here    ----------------------- #
                                
        self.counts_n = 0
        self.counts_e = 0
        self.name_list=[]
        self.email_list = []
        self.model_img = model

        # name of csv file
        self.csvFile = "Send_list/Send_to_emails.csv"
        # head names
        self.fields = ['name', 'email', 'CID','certFileName']
        # open csv for write
        with open(self.csvFile, 'w',newline='') as toCsv:
            # creating a csv writer object
            csvwriter = csv.writer(toCsv)
            # writing the fields
            csvwriter.writerow(self.fields)
        #ID
        self.ID=str(0)


        def select_file():
            filetypes = (
                ('text files', '*.txt'),
                ('All files', '*.*')
            )

            filename = askopenfilename(
                title='Open a file',
                initialdir='/',
                filetypes=filetypes)

            showinfo(
                title='Selected File',
                message=filename
            )

            return filename
        
        def get_count(flag):
            if(flag == 0):
                self.counts_n += 1
                return self.counts_n
            if(flag == 1):
                self.counts_e += 1
                return self.counts_e
            
        
        def choose_excel():
            self.BTM.insert(tk.INSERT,"\nChoose Excel")
            filename = select_file()
            if not filename:
                self.BTM.insert(tk.INSERT,"\nSelect Correct format (EXCEL FILES ONLY")
                return
            self.BTM.insert(tk.INSERT,"\nOpen File : "+filename)
            self.BTM.insert(tk.INSERT,"\nFile Selected")
            self.ENTRY_LEFT.delete(0,"end")
            self.ENTRY_LEFT.insert(0, filename)
            self.counts_n = 0
            self.counts_e = 0
            i,j = 0,0;
            self.BOX_LEFT_1.configure(state ='normal')
            self.BOX_LEFT_2.configure(state ='normal')
            self.BOX_LEFT_1.delete('1.0', END)
            self.BOX_LEFT_2.delete('1.0', END)
            # read list contain participants name
            data = pd.read_excel(filename)
            # fetch and store name column values 
            self.name_list = data["Name"].tolist()
            self.email_list = data["Email"].tolist()
            name_result = map(lambda x: str(get_count(0))+". "+x+"\n", self.name_list)
            email_result = map(lambda x: str(get_count(1))+". "+x+"\n", self.email_list)
            self.BOX_LEFT_1.insert(tk.INSERT,''.join(list(name_result)))
            self.BOX_LEFT_2.insert(tk.INSERT,''.join(list(email_result)))
            self.BOX_LEFT_1.configure(state ='disabled')
            self.BOX_LEFT_2.configure(state ='disabled')
            self.BTM.insert(tk.INSERT,"\nData Displayed")

        def choose_image():
            self.BTM.insert(tk.INSERT,"\nChoose Model certificate")
            filename = select_file()
            if not filename:
                self.BTM.insert(tk.INSERT,"\nSelect A valid Image File (.png)")
                return
            self.BTM.insert(tk.INSERT,"\nOpen File : "+filename)
            model = PIL.Image.open(filename)
            self.model_img = model
            self.IM_width, self.IM_height = model.size
            newsize = (300, 150)
            model_show = model.resize(newsize)
            IMG_CERT = ImageTk.PhotoImage(model_show)
            self.LABEL_RIGHT_2.configure(image=IMG_CERT)
            self.LABEL_RIGHT_2.image = IMG_CERT
            
            self.ENTRY_LEFT_IMW.delete(0,"end")
            self.ENTRY_LEFT_IMH.delete(0,"end")
            self.ENTRY_LEFT_IMW.insert(0, self.IM_width)
            self.ENTRY_LEFT_IMH.insert(0, self.IM_height)
            self.BTM.insert(tk.INSERT,"\nFind H, W")
            
        def generate():
            self.BTM.insert(tk.INSERT,"\nCertificate Generating ..... wait ")
            for name in self.name_list:
                image = self.model_img      
                background = PIL.Image.new("RGB",image.size, (255, 255, 255))
                # 3 is the alpha channel
                background.paste(image, mask=image.split()[3])                 
                image=background
                d = ImageDraw.Draw(image)
                # text size
                w,h = d.textsize(name)
                # adjust name to center 
                location = ((self.IM_width-w)/2 -(w+10),(self.IM_height-h)/2 -30)
                # add name
                d.text(location, name, fill = self.text_color, font = self.font)
                # save certificates in pdf format
                image.save("Generate/CID_"+self.ID+"_"+ name + ".pdf")
    
                # generate csv file
                with open(self.csvFile, 'a',newline='') as toCsv:
                    csvwriter = csv.writer(toCsv)
                    row_data = [name,self.email_list[int(self.ID)],"CID_"+self.ID,"CID_"+self.ID+"_"+ name]
                    csvwriter.writerow(row_data)
    
                self.ID=str(int(self.ID)+1)
            self.BTM.insert(tk.INSERT,"\nCertificate Generating compleated.....")

        def send_to_emails():
            self.BTM.insert(tk.INSERT,"\nSEND TO EMAIL")
            
            emailWindow = Toplevel(master)
            emailWindow.title("Send to Email")
            emailWindow.geometry("800x500")

            #full window row configure
            emailWindow.grid_rowconfigure(0, weight=1)
            emailWindow.grid_rowconfigure(1, weight=1)
            #full window column configure
            emailWindow.columnconfigure(0, weight=1)
            emailWindow.columnconfigure(1, weight=1)
            #labelled frames
            frame_left  =  LabelFrame(emailWindow,text="Insert Your Gmail credentials",labelanchor="n",bg="white",bd=1,fg="red",font=self.label_frame_font)
            frame_right  =  LabelFrame(emailWindow,text="Sended List",labelanchor="n",bg="white",bd=1,fg="red",font=self.label_frame_font)
            BTM = ScrolledText(emailWindow,height=1,width=5,bg="white",fg="green",bd=1)
            self.BTM.insert(tk.INSERT,"-------------LOGS----------\nStarting...")
        
            #frame grids
            frame_left.grid(row=0,column=0,sticky="nsew")
            frame_right.grid(row=0,column=1,sticky="nsew")
            BTM.grid(row=1,column=0,sticky="nsew",columnspan=2)

            # LEFT frame
            frame_left.grid_rowconfigure(0, weight=1)
            frame_left.grid_rowconfigure(1, weight=1)
            frame_left.columnconfigure(0, weight=1)
            
            frame_left_top  =  LabelFrame(frame_left,text="",labelanchor="n",bg="white",bd=0,fg="red",font=self.label_frame_font)
            frame_left_bottom  =  LabelFrame(frame_left,text="Select",labelanchor="n",bg="white",bd=1,fg="red",font=self.label_frame_font)

            frame_left_top.grid(row=0,column=0,sticky="nsew")
            frame_left_bottom.grid(row=1,column=0,sticky="nsew")


            frame_left_top.grid_rowconfigure(0, weight=1)
            frame_left_top.grid_rowconfigure(1, weight=1)
            frame_left_top.grid_rowconfigure(2, weight=1)
            
            frame_left_top.columnconfigure(0, weight=1)
            frame_left_top.columnconfigure(1, weight=1)
            frame_left_top.columnconfigure(2, weight=1)
            frame_left_top.columnconfigure(3, weight=1)
            frame_left_top.columnconfigure(4, weight=1)

            LABEL_LEFT_1 = Label(frame_left_top,text="Email : ",padx=20,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
            ENTRY_LEFT_1 = Entry(frame_left_top,bg="white",fg="green",textvariable = ButtonVar1,bd=2)
            LABEL_LEFT_2 = Label(frame_left_top,text="Name : ",padx=20,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
            ENTRY_LEFT_2 = Entry(frame_left_top,bg="white",fg="green",textvariable = ButtonVar1,bd=2)
            LABEL_LEFT_3 = Label(frame_left_top,text="Password : ",padx=20,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
            ENTRY_LEFT_3 = Entry(frame_left_top,bg="white",fg="green",textvariable = ButtonVar1,bd=2)

            LABEL_LEFT_1.grid(row=0,column=1,sticky="w")
            ENTRY_LEFT_1.grid(row=0,column=3,sticky="w")
            LABEL_LEFT_2.grid(row=1,column=1,sticky="w")
            ENTRY_LEFT_2.grid(row=1,column=3,sticky="w")
            LABEL_LEFT_3.grid(row=2,column=1,sticky="w")
            ENTRY_LEFT_3.grid(row=2,column=3,sticky="w")



            frame_left_bottom.grid_rowconfigure(0, weight=1)
            frame_left_bottom.grid_rowconfigure(1, weight=1)
            frame_left_bottom.grid_rowconfigure(2, weight=1)
            frame_left_bottom.columnconfigure(0, weight=1)
            frame_left_bottom.columnconfigure(1, weight=1)
            frame_left_bottom.columnconfigure(2, weight=1)
            frame_left_bottom.columnconfigure(3, weight=1)
            frame_left_bottom.columnconfigure(4, weight=1)

            
            BTN_LEFT_21   = Button(frame_left_bottom,text="open CSV", width = WIDTH_BTN,bg="white",bd=1,fg="black",command=lambda:choose_excel())
            ENTRY_LEFT_21 = Entry(frame_left_bottom,bg="white",fg="green",textvariable = ButtonVar1,bd=2)
            BTN_LEFT_22   = Button(frame_left_bottom,text="CERT FOLDER", width = WIDTH_BTN,bg="white",bd=1,fg="black",command=lambda:choose_excel())
            ENTRY_LEFT_22 = Entry(frame_left_bottom,bg="white",fg="green",textvariable = ButtonVar1,bd=2)

            BTN_LEFT_21.grid(row=0,column=1,sticky="ew")
            ENTRY_LEFT_21.grid(row=0,column=3,sticky="ew")
            BTN_LEFT_22.grid(row=1,column=1,sticky="ew")
            ENTRY_LEFT_22.grid(row=1,column=3,sticky="ew")
            
            # RIGHT frame
            frame_right.grid_rowconfigure(0, weight=1)
            frame_right.columnconfigure(0, weight=1)

            BOX_right = ScrolledText(frame_right,width=5,bd=0)

            BOX_right.grid(row=0,column=0,sticky="nsew")
      
            
            


        #------------------ MENUS ----------------------------#
        def donothing():
           filewin = Toplevel(root)
           button = Button(filewin, text="Do nothing button")
           button.pack()
            
            
        
       
root = Tk()
hack_gui = Convert(root)
root.mainloop()
