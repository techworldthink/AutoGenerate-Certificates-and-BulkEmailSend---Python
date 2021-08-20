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

        """bgmodel = PIL.Image.open("Excel.jpg")
        newsize = (200, 100)
        bgmodel = bgmodel.resize(newsize)
        F_IMG_CERT = ImageTk.PhotoImage(bgmodel)"""

        self.LOG="Starting..."

        
        #labelled frames
        self.frame_left  =  LabelFrame(master,text="SELECT EXCEL FILE",labelanchor="n",bg="white",bd=15,fg="red",font=self.label_frame_font)
        self.frame_right  =  LabelFrame(master,text="GENERATE-CERTIFICATES",labelanchor="n",bg="white",bd=15,fg="red",font=self.label_frame_font)
        self.BTM= Label(master,text=self.LOG,bg="black",fg="green",font=self.frame2_font,width=1, anchor='w')
        
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
        self.LABEL_LEFT= Label(self.frame_left,text="DETAILS OF PARTICIPANTS",padx=0,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        self.ENTRY_LEFT = Entry(self.frame_left,bg="white",fg="green",textvariable = ButtonVar1,bd=0)
        self.BOX_LEFT_1 = ScrolledText(self.frame_left,width=5)
        self.BOX_LEFT_2 = ScrolledText(self.frame_left,width=5) 
        self.BTN_LEFT = Button(self.frame_left,text="Open Excel File",height = 2, width = WIDTH_BTN,bg="#d7d9db",bd=10,fg="black",command=lambda:choose_excel())
        #self.BTN_LEFT.image=F_IMG_CERT
        #self.LABEL_LEFT_B= Label(self.frame_left,text="m",padx=0,pady=0,bg="black",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        #componants grid for frame 1
        self.LABEL_LEFT.grid(row=0,column=0,sticky="nsew",columnspan=2)
        self.BOX_LEFT_1.grid(row=1,column=0,sticky="nsew")
        self.BOX_LEFT_2.grid(row=1,column=1,sticky="nsew")
        self.BTN_LEFT.grid(row=2,column=0,sticky="nsew",columnspan=2)
        self.ENTRY_LEFT.grid(row=3,column=0,sticky="ew",columnspan=2)
        #self.LABEL_LEFT_B.grid(row=4,column=0,sticky="sew",columnspan=2)
     
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
        self.BTN_RIGHT_1 = Button(self.frame_right,text="Open Model Certificate",height = 2, width = WIDTH_BTN,bg="#d7d9db",bd=10,fg="black",command=lambda:choose_image())
        self.LABEL_RIGHT_2.image=IMG_CERT
        
        self.LABEL_RIGHT_IM = Label(self.frame_right,text="Size",padx=0,pady=0,bg="white",fg="black",font=self.frame2_font,width=WIDTH_LABEL)
        self.ENTRY_LEFT_IMW = Entry(self.frame_right,bg="white",fg="black",textvariable = ButtonVar2,bd=1)
        self.ENTRY_LEFT_IMH = Entry(self.frame_right,bg="white",fg="black",textvariable = ButtonVar3,bd=1)

        self.BTN_RIGHT_GNT = Button(self.frame_right,text="GENERATE",height = 2, width = WIDTH_BTN,bg="#d7d9db",bd=10,fg="black",command=lambda:generate())

        #componants grid for frame 1
        self.LABEL_RIGHT.grid(row=0,column=0,sticky="nsew",columnspan=3)
        self.LABEL_RIGHT_2.grid(row=1,column=0,sticky="nsew",columnspan=3)
        self.BTN_RIGHT_1.grid(row=2,column=0,sticky="ew",columnspan=3)
        self.LABEL_RIGHT_IM.grid(row=3,column=0,sticky="ew")
        self.ENTRY_LEFT_IMW.grid(row=3,column=1,sticky="ew")
        self.ENTRY_LEFT_IMH.grid(row=3,column=2,sticky="ew")
        self.BTN_RIGHT_GNT.grid(row=4,column=0,sticky="ew",columnspan=3)

        
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
            self.BTM.config(text = self.BTM.cget("text")+"\n Choose Excel")
            filename = select_file()
            if not filename:
                self.BTM.config(text = self.BTM.cget("text")+"\n Select Correct format (EXCEL FILES ONLY)")
                return
            self.BTM.config(text = self.BTM.cget("text")+"\n File Selected")
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
            self.BTM.config(text = self.BTM.cget("text")+"\n Data Displayed")

        def choose_image():
            self.BTM.config(text = self.BTM.cget("text")+"\n Choose Model certificate")
            filename = select_file()
            if not filename:
                self.BTM.config(text = self.BTM.cget("text")+"\n Select A valid Image File (.png)")
                return
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
            self.BTM.config(text = self.BTM.cget("text")+"\n Find H, W")
            
        def generate():
            self.BTM.config(text = self.BTM.cget("text")+"\n Certificate Generating ..... wait ")
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
            self.BTM.config(text = self.BTM.cget("text")+"\n Certificate Generating compleated..... ")
            
            
        
       
root = Tk()
hack_gui = Convert(root)
root.mainloop()
