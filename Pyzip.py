# Import Modules.
import datetime
import os
import zipfile
from tkinter import *
from tkinter import filedialog, messagebox, ttk


class Pyzip:
    def __init__(self,root):
        self.root = root
        
        self.zip_file = ''
        self.nodes = dict()
        
        banner_img = PhotoImage(file="banner.png")
        banner = Canvas(self.root,width=400,height=100)
        banner.create_image(0,0, anchor=NW,image=banner_img)
        banner.image = banner_img
        banner.place(x=0,y=0)
        
        # Notebook / tabbed Widget
        main_frame = Frame(self.root,bd=2)
        main_frame.place(x=0,y=100,height=350,width=400)
        # Defines and places the notebook widget
        nb = ttk.Notebook(main_frame)
        nb.pack(expand=1, fill="both")
        
        # Adds tab 1 of the notebook
        page1 = ttk.Frame(nb)
        nb.add(page1, text='Unzip File')
        # Adds tab 2 of the notebook
        page2 = ttk.Frame(nb)
        nb.add(page2, text='Create Zip')
        
        # add widgets in page 1 for unzip files...
        
        self.unzip_file_entry = ttk.Entry(page1,width=45)
        self.unzip_file_entry.place(x=10,y=10)
        
        browse_unzip_file_path_button = ttk.Button(page1,text="Browse Path",command=self.get_zip)
        browse_unzip_file_path_button.place(x=300,y=9)
        
        # Create frame for adding table data..
        data_display_frame = Frame(page1,bd=2,relief=RIDGE)
        data_display_frame.place(x=0,y=45,width=390,height=180)
        
        # add treeview and scrollbar ..
        scrollbar_y = Scrollbar(data_display_frame)
        scrollbar_y.pack( side = RIGHT, fill = Y )
        
        scrollbar_x = Scrollbar(data_display_frame,orient=HORIZONTAL)
        scrollbar_x.pack( side = BOTTOM, fill = X )
        
        columns = ("Sno.","Filename","Modified","Size")
        self.treeview = ttk.Treeview(data_display_frame, height=7, 
		show="headings", columns=columns,yscrollcommand = scrollbar_y.set,
            xscrollcommand = scrollbar_x.set)
        
        # add column properties
        self.treeview.column("Sno.",width=5, anchor='center')
        self.treeview.column("Filename",stretch=YES,anchor='center')
        self.treeview.column("Modified",stretch=YES,anchor='center')
        self.treeview.column("Size",stretch=YES,anchor = 'center')
        # adding heading 
        self.treeview.heading("Sno.", text="Sno.")
        self.treeview.heading("Filename", text="Filename") # Show header
        self.treeview.heading("Modified", text="Modified")
        self.treeview.heading("Size",text = "Size")
        
        scrollbar_y.config( command = self.treeview.yview )
        scrollbar_x.config( command = self.treeview.xview )
        self.treeview.pack(side=TOP, fill=BOTH)
        
        self.unzip_file_path_entry = ttk.Entry(page1,width=45)
        self.unzip_file_path_entry.place(x=10,y=240)
        
        browse_path_button = ttk.Button(page1,text="Unzip Path",command=self.get_unzip_path)
        browse_path_button.place(x=300,y=239)
        
        ttk.Button(page1,text="Unzip File",width=15,command=self.unzip_file).place(x=30,y=280)
        ttk.Button(page1,text="Cancel",width=15,command=self.cancel_1).place(x=140,y=280)
        ttk.Button(page1,text="Exit App.",width=15,command=self.root.quit).place(x=260,y=280)
        
         # add widgets in page 2 for creating zip files...
        
        self.zip_folder_entry = ttk.Entry(page2,width=45)
        self.zip_folder_entry.place(x=10,y=10)
        
        browse_zip_folder_path_button = ttk.Button(page2,text="Browse Path",command=self.get_folder)
        browse_zip_folder_path_button.place(x=300,y=9)
        
        # Create frame for adding table data..
        data_display_frame_page2 = Frame(page2,bd=2,relief=RIDGE)
        data_display_frame_page2.place(x=0,y=45,width=390,height=180)
        
        # add treeview and scrollbar ..
        scrollbar_y_page2 = Scrollbar(data_display_frame_page2)
        scrollbar_y_page2.pack( side = RIGHT, fill = Y )
        
        scrollbar_x_page2 = Scrollbar(data_display_frame_page2,orient=HORIZONTAL)
        scrollbar_x_page2.pack( side = BOTTOM, fill = X )
        
        self.treeview_page2 = ttk.Treeview(data_display_frame_page2, height=7, 
		yscrollcommand = scrollbar_y_page2.set,
            xscrollcommand = scrollbar_x_page2.set)
        
        self.treeview_page2.heading('#0', text='->Selected Directory Contents', anchor='w')        
        
        scrollbar_y_page2.config( command = self.treeview_page2.yview )
        scrollbar_x_page2.config( command = self.treeview_page2.xview )
        self.treeview_page2.pack(side=TOP, fill=BOTH)
        
        self.zip_file_path_entry = ttk.Entry(page2,width=45)
        self.zip_file_path_entry.place(x=10,y=240)
        
        browse_zip_path_button = ttk.Button(page2,text="Zip Path",command=self.get_zip_path)
        browse_zip_path_button.place(x=300,y=239)
        
        ttk.Button(page2,text="Create Zip",width=15,command=self.create_zip).place(x=30,y=280)
        ttk.Button(page2,text="Cancel",width=15,command=self.cancel_2).place(x=140,y=280)
        ttk.Button(page2,text="Exit App.",width=15,command=self.root.quit).place(x=260,y=280)
        
    # Methods .............................
    def get_zip(self):
        self.zip_file = filedialog.askopenfilename()
        if  zipfile.is_zipfile(self.zip_file):
            self.unzip_file_entry.insert(END, self.zip_file)
            # zip file handler  
            zip = zipfile.ZipFile(self.zip_file)
            # list available files in the container
            #print ("using name list \n",zip.namelist()) # only print the name of zip file content
            #print('\n using info list \n',zip.infolist()) # prints filename and size of zip content
            # zip.printdir() # print filename, modified data, size but only in console..
            #print(zip.infolist())
            for i, info in enumerate(zip.infolist()):
                #print(info.filename)
                #print('\tModified:\t'+str(datetime.datetime(*info.date_time) ))
                #print('\tUncompressed:\t'+str(info.file_size)+' bytes')
                self.treeview.insert('', i, values=(i+1,info.filename, str(datetime.datetime(*info.date_time)), str(info.file_size)+' bytes'))
        else:
            messagebox.showerror("Error !",'Please Select ZIP file')
            
    def get_unzip_path(self):
        self.unzip_path = filedialog.askdirectory()
        if self.unzip_path:
            self.unzip_file_path_entry.insert(END, self.unzip_path)
        
    
    def unzip_file(self):
        if self.unzip_path:
            if self.zip_file:
                try:
                    # Opeaning the Zip File in Read Mode
                    with zipfile.ZipFile(self.zip_file,'r') as zip:
                        base=os.path.basename(self.zip_file)
                        filename = os.path.splitext(base)[0]
                        extract_path = os.path.join(self.unzip_path,filename)
                        os.mkdir(extract_path)
                        zip.extractall(extract_path)
                        messagebox.showinfo("Sucess","File is extracted at\n"+str(extract_path))
                except Exception as e:
                    messagebox.showerror("Error",e)
            else:
                messagebox.showwarning('Warning','First Select a Zip File ')
        else:
            messagebox.showwarning('Warning','First Select path where to \n unzip file ')
            
    def get_folder(self):
        self.zip_path = filedialog.askdirectory()
        if self.zip_path:
            self.zip_folder_entry.insert(END,self.zip_path)
            abspath = os.path.abspath(self.zip_path)
            self.insert_node('',abspath , abspath)
            self.treeview_page2.bind('<<TreeviewOpen>>', self.open_node)
    
    def insert_node(self, parent, text, abspath):
        node = self.treeview_page2.insert(parent, 'end', text=text, open=False)
        if os.path.isdir(abspath):
            self.nodes[node] = abspath
            self.treeview_page2.insert(node, 'end')

    def open_node(self, event):
        node = self.treeview_page2.focus()
        abspath = self.nodes.pop(node, None)
        if abspath:
            self.treeview_page2.delete(self.treeview_page2.get_children(node))
            for p in os.listdir(abspath):
                self.insert_node(node, p, os.path.join(abspath, p))
            
    
    
    def get_zip_path(self):
        self.zip_save_path = filedialog.askdirectory()
        if self.zip_save_path:
            self.zip_file_path_entry.insert(END, self.zip_save_path)
        
        
    
    
    def create_zip(self):
        try:
            # path to folder which needs to be zipped 
            directory = self.zip_path
            # initializing empty file paths list
            file_paths = [] 
            # crawling through directory and subdirectories
            for root, directories, files in os.walk(directory):
                for filename in files:
                    # join the two strings in order to form the full filepath.
                    filepath = os.path.join(root, filename)
                    file_paths.append(filepath)
            base=os.path.basename(directory)
            save_path = os.path.join(self.zip_save_path,str(base)+'.zip')
            with zipfile.ZipFile(save_path,'w') as zip: 
                # writing each file one by one 
                for file in file_paths:
                    zip.write(file) 
            
            messagebox.showinfo('Sucess',"Zip File is Create at \n"+save_path)
        except Exception as e:
            messagebox.showerror("Error",e)
    
    
    def cancel_1(self):
        self.unzip_file_entry.delete(0,END)
        self.unzip_file_path_entry.delete(0,END)
        self.treeview.delete(*self.treeview.get_children())  
        
    def cancel_2(self):
        self.zip_file_path_entry.delete(0,END)
        self.zip_folder_entry.delete(0,END)
        self.treeview_page2.delete(*self.treeview_page2.get_children())
               
            
            
            
            
            
    
        
        
               
root = Tk()
root.title("PyZip- Tool To Zip and UnZip Files")
root.geometry("400x450")
root.resizable(0,0)
Pyzip(root)
root.mainloop()

