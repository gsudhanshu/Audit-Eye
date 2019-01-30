from Tkinter import *
import pandas as pd
from ScrolledText import ScrolledText 
from PIL import ImageTk, Image
from tkFileDialog import askopenfilename
from tkcalendar import DateEntry
import ttk
import os
from shutil import copyfile
import numpy as np

def frame(root, side):
    w=Frame(root)
    w.pack(side=side, expand=YES, fill=BOTH)
    return w

class Application(Frame):
    
    class Project:
        def __init__(self, projectName, fy_end, timing, creator, sector, fname=''):
            self.fname = fname
            self.projectName = projectName
            self.fy_end = fy_end
            self.timing = timing
            self.creator = creator
            self.sector = sector
            self.dataSource = ''
            self.glInputFile = ''
            self.tbInputFile = ''
            self.caInputFile = ''
            self.gldata = None
            self.tbdata = None
            self.cadata = None
            self.min_entry_dt = ''
            self.max_entry_dt = ''
            self.min_eff_dt = ''
            self.max_eff_dt = ''
        def getProjectFName(self):
            return self.fname
        def getProjectName(self):
            return self.projectName
        def getFYend(self):
            return self.fy_end
        def getTiming(self):
            return self.timing
        def getCreator(self):
            return self.creator
        def getSector(self):
            return self.sector
        def getGLInputFile(self):
            return self.glInputFile
        def getTBInputFile(self):
            return self.tbInputFile
        def getCAInputFile(self):
            return self.caInputFile
        def setGLInputFile(self, glInputF):
            self.glInputFile = glInputF
            self.gldata = pd.read_excel(glInputF)
        def setTBInputFile(self, tbInputF):
            self.tbInputFile = tbInputF
            self.tbdata = pd.read_excel(tbInputF)
        def setCAInputFile(self, caInputF):
            self.caInputFile = caInputF
            self.cadata = pd.read_excel(caInputF)
        def getGLData(self):
            return self.gldata
        def getTBData(self):
            return self.tbdata
        def getCAData(self):
            return self.cadata
        def setEntryEffDates(self, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt):
            self.min_entry_dt = min_entry_dt
            self.max_entry_dt = max_entry_dt
            self.min_eff_dt = min_eff_dt
            self.max_eff_dt = max_eff_dt

    def init_dashboard(self):
        self.l1.destroy()
        #f1: left pane
        f1 = frame(self.w, LEFT)
        Label(f1, text="Financial Statement Profiling", bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Analyze Balance Sheet", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Analyze Income Statement", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f1, text="Business Unit Map", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f1, text="Financial Statement Tie-out", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f1, text="Significant Accounts Identification", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Income Analysis", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: second pane
        f2 = frame(self.w, LEFT)
        Label(f2, text="Validation", bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f2, text="Form 572: Data and Analytics delivery memo", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f2, text="Validation Results Summary", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f2, text="Trial Balance Roll-forward", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f2, text="Back Posting", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Label(f2, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f2, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f2, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #f3: third pane
        f3 = frame(self.w, LEFT)
        Label(f3, text="Process Analysis", bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Process Map", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Preparer Map", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Analyze preparers, approvers and segregation of duties", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Identify and Understand Booking Patterns", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Tagging Analysis - Journals", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Label(f3, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f3, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f3.pack(expand=YES, fill=BOTH)
        #f4: Last pane
        f4 = frame(self.w, LEFT)
        Label(f4, text="Account and Journal Entry Analysis", bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Analyze Correlation b/w 2 accounts", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Analyze Correlation b/w 3 accounts", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Analyze Relationship of 2 accounts", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Gross Margin Analysis", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Cutoff Analysis of GL accounts", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Additional Reports", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Custom Analytics - visualization", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        f4.pack(expand=YES, fill=BOTH)

    def __init__(self):
        pd.set_option('display.max_colwidth', -1)
        pd.set_option('display.max_rows', 500)
        Frame.__init__(self)
        self.project=None
        self.status = StringVar()
        self.status.set("Started")
        self.pack(expand=YES, fill=BOTH)
        self.master.title('DA Analyze')
        self.master.iconname("DA")
        mBar = Frame(self, relief=RAISED, borderwidth=2)
        mBar.pack(fill=X)
        fileBtn = self.makeFileMenu(mBar)
        toolsBtn = self.makeToolsMenu(mBar)
        helpBtn = self.makeHelpMenu(mBar)
        mBar.tk_menuBar(fileBtn, toolsBtn, helpBtn)
        self.w = Frame(self, relief=SUNKEN, borderwidth=1)
        self.w.pack(side=TOP, expand=YES, fill=BOTH)
        os.chdir("images")
        img = ImageTk.PhotoImage(Image.open("base.jpg"))
        os.chdir("..")
        self.l1 = Label(self.w, image=img, relief=SUNKEN)
        self.l1.pack(side=TOP, fill=BOTH, expand=YES, padx=5)
        self.l1.image = img
        lbl_status = Entry(self.w, textvariable=self.status, justify=LEFT, relief=RAISED)
        lbl_status.pack(side=BOTTOM, fill=BOTH, expand=YES, padx=5)

    def input_parameters_window(self):
        if self.project == None:
            self.status.set("First Create Project or Load existing Project!")
            return
        elif self.project.getGLInputFile() == '' or self.project.getTBInputFile() == '' or self.project.getCAInputFile() == '':
            self.status.set("First Upload Data Files. Select Tools -> Manage Data")
            return        
        ipw = Toplevel(self)
        ipw.wm_title("Validate Input Parameters")
        #read gl and get max and min effective and entry dates
        glData = self.project.getGLData()
        min_entry_dt = glData['Posting Date'].min()
        max_entry_dt = glData['Posting Date'].max()
        min_eff_dt = glData['Posting Date'].min()
        max_eff_dt = glData['Posting Date'].max()
        #f1: left pane
        f1 = frame(ipw, LEFT)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_dt = Label(f1, text="Start Date:", relief=FLAT, anchor="w")
        lbl_start_dt.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_dt = Label(f1, text="End Date:", relief=FLAT, anchor="w")
        lbl_end_dt.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Center pane
        f2 = frame(ipw, LEFT)
        Label(f2, text="Entry Date", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_entry = Label(f2, text=min_entry_dt, relief=SUNKEN)
        lbl_start_entry.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_entry = Label(f2, text=max_entry_dt, relief=SUNKEN)
        lbl_end_entry.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f2, text="Cancel", command=ipw.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        f2.pack(expand=YES, fill=BOTH)
        #f3: Right pane
        f3 = frame(ipw, LEFT)
        Label(f3, text="Effective Date", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_effective = Label(f3, text=min_eff_dt, relief=SUNKEN)
        lbl_start_effective.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_effective = Label(f3, text=max_eff_dt, relief=SUNKEN)
        lbl_end_effective.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Ok and Next", command=lambda: self.ipt_param_JEvalidate_window(self, ipw, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)

    def ipt_param_JEvalidate_window(self, master, ipw, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt):
        #save entry and effective dates in project object
        master.project.setEntryEffDates(min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt)
        ipw.destroy()
        #open next window
        ipjw = Toplevel(master)
        ipjw.wm_title("Validate Input Parameters")
        glData = master.project.getGLData()
        #A. highlight more than 5 JE line items
        glData_subset = glData[['Entry No_', 'Amount']]
        countli_byJE = glData_subset.pivot_table(index=['Entry No_'], aggfunc='count')
        countli_byJE = countli_byJE.rename(columns = {'Amount':'Count'})
        countli_byJE = countli_byJE.sort_values(by=['Count'], ascending=False)
        f1 = frame(ipjw, TOP)
        Label(f1, text="JE's with count of line items to check if multiple transactions within same JE").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        jeli_text = Text(f1, height=10, width=80)
        jeli_text.insert(END, countli_byJE) #display dataframe in text
        jeli_scroll = Scrollbar(f1, command= jeli_text.yview)
        jeli_text.configure(yscrollcommand=jeli_scroll.set)
        jeli_text.pack(side=LEFT)
        jeli_scroll.pack(side=RIGHT, fill=Y)
        f1.pack(expand=YES, fill=BOTH)
        f2 = frame(ipjw, TOP)
        #B. Unbalanced JEs
        amount_by_JE = glData_subset.pivot_table(index=['Entry No_'])
        amount_by_JE = amount_by_JE.replace(0, np.nan)
        unbalancedJE = amount_by_JE.dropna(how='any', axis=1) 
        unbalancedJE = amount_by_JE.replace(np.nan, 0) #to be on safe side
        Label(f2, text="Unbalanced JE's").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        jelist_text = Text(f2, height=10, width=80)
        jelist_text.insert(END, unbalancedJE) #display dataframe in text
        jelist_scroll = Scrollbar(f2, command= jelist_text.yview)
        jelist_text.configure(yscrollcommand=jelist_scroll.set)
        jelist_text.pack(side=LEFT)
        jelist_scroll.pack(side=RIGHT, fill=Y)
        f2.pack(expand=YES, fill=BOTH)
        f3 = frame(ipjw, TOP)
        Button(f3, text="Cancel", command=ipjw.destroy).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Approve and Next", command=ipjw.destroy).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)

    def input_data_window(self):
        if self.project == None:
            self.status.set("First Create Project or Load existing Project!")
            return
        idw = Toplevel(self)
        idw.wm_title("Manage Data Files")
        #f1: left pane
        f1 = frame(idw, LEFT)
        lbl_data_source = Label(f1, text="Data Source", relief=FLAT, anchor="w")
        lbl_data_source.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_gl = Label(f1, text="Journal Entry Data File", relief=FLAT, anchor="w")
        lbl_gl.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_tb = Label(f1, text="Trial Balance Data File", relief=FLAT, anchor="w")
        lbl_tb.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_ca = Label(f1, text="Chart of Accounts Data File", relief=FLAT, anchor="w")
        lbl_ca.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT, anchor="w").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: center pane
        f2 = frame(idw, LEFT)
        ipt_ds = ttk.Combobox(f2, values=("TALLY","SAP"))
        ipt_ds.set("TALLY")
        ipt_ds.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        self.glInputF = StringVar()
        self.glInputF.set(self.project.getGLInputFile())
        ipt_gl = Entry(f2, relief=SUNKEN, width=100, textvariable=self.glInputF, state='disabled')
        ipt_gl.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        self.tbInputF = StringVar()
        self.tbInputF.set(self.project.getTBInputFile())
        ipt_tb = Entry(f2, relief=SUNKEN, width=100, textvariable=self.tbInputF, state='disabled')
        ipt_tb.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        self.caInputF = StringVar()
        self.caInputF.set(self.project.getCAInputFile())
        ipt_ca = Entry(f2, relief=SUNKEN, width=100, textvariable=self.caInputF, state='disabled')
        ipt_ca.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        def onSubmit(parent, master, dataSource):
            if master.glInputF.get() == '' or master.tbInputF.get() == '' or master.caInputF.get() == '':
                master.status.set("All 3 data files are mandatory!")
                return
            if self.changeInputF == 0:
                master.status.set("No change in data files")
            else:
                #create a data folder in pwd if it does not exist
                try:
                    os.chdir("Data")
                except OSError:
                    if 'Data' not in os.listdir(os.getcwd()):
                        os.mkdir("Data")
                        os.chdir("Data")
                finally:
                    #copy all 3 files in data folder
                    cpflag = 0
                    try:
                        copyfile(master.glInputF.get(), ""+master.project.getProjectName()+"_gl")
                        cpflag +=1
                    except:
                        #log error!
                        master.status.set("Error in saving GL Input File")
                    try:
                        copyfile(master.tbInputF.get(), ""+master.project.getProjectName()+"_tb")
                        cpflag +=1
                    except:
                        master.status.set("Error in saving TB Input File")
                    try:
                        copyfile(master.caInputF.get(), ""+master.project.getProjectName()+"_ca")
                        cpflag +=1
                    except:
                        master.status.set("Error in saving CoA Input File")
                    if cpflag == 0:
                        master.status.set("Error in saving Data files. Try again!")
                        cwd = os.getcwd()
                        if cwd[-4:] == "Data":
                            os.chdir("..")
                        return
                    else:
                        master.status.set(str(cpflag)+" Data files saved. Now loading data...")
                    cwd = os.getcwd()
                    if cwd[-4:] == "Data":
                        os.chdir("..")
                    #save information in project.p file
                    pf = open(master.project.getProjectFName(), "w") #existing file will be overwritten
                    pf.write("ProjectName="+master.project.getProjectName()+"\n")
                    pf.write("FY_end="+master.project.getFYend()+"\n")
                    pf.write("ProjectTiming="+master.project.getTiming()+"\n")
                    pf.write("ProjectCreator="+master.project.getCreator()+"\n")
                    pf.write("Sector="+master.project.getSector()+"\n")
                    pf.write("GLinputFile="+os.path.abspath(""+master.project.getProjectName()+"_gl")+"\n")
                    pf.write("TBinputFile="+os.path.abspath(""+master.project.getProjectName()+"_tb")+"\n")
                    pf.write("CAinputFile="+os.path.abspath(""+master.project.getProjectName()+"_ca")+"\n")
                    pf.close()
                    #Load data files from saved location
                    master.load_project_file(master.project.getProjectFName())
            parent.destroy()
        Label(f2, text=" ", relief=FLAT, anchor="w").pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f2, text="Cancel", command=idw.destroy).pack(side=LEFT, padx=10, pady=10)
        Button(f2, text="Submit", command=lambda: onSubmit(idw, self, ipt_ds.get())).pack(side=LEFT, padx=10, pady=10)
        Label(f2, text=" ", relief=FLAT, anchor="w").pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #f3: rightmost pane
        f3 = frame(idw, LEFT)
        self.changeInputF = 0
        def browseGLinputF(master):
            master.glInputF.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.glInputF.get() == (): #in case of cancel or no selection
                master.glInputF.set('')
                return
            self.changeInputF += 1
        def browseTBinputF(master):
            master.tbInputF.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.tbInputF.get() == (): #in case of cancel or no selection
                master.tbInputF.set('')
                return
            self.changeInputF += 1
        def browseCAinputF(master):
            master.caInputF.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.caInputF.get() == (): #in case of cancel or no selection
                master.caInputF.set('')
                return
            self.changeInputF += 1
        Label(f3, text=" ", relief=FLAT, anchor="w").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Browse...", command=lambda: browseGLinputF(self)).pack(side=TOP, padx=10, pady=10)
        Button(f3, text="Browse...", command=lambda: browseTBinputF(self)).pack(side=TOP, padx=10, pady=10)
        Button(f3, text="Browse...", command=lambda: browseCAinputF(self)).pack(side=TOP, padx=10, pady=10)
        Label(f3, text=" ", relief=FLAT, anchor="w").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f3.pack(expand=YES, fill=BOTH)
        #grab_set to refrain any activity on main window
        idw.grab_set()

    def create_project_window(self):
        cpw = Toplevel(self)
        cpw.wm_title("Create Project")
        #f1: left pane
        f1 = frame(cpw, LEFT)
        lbl_project_name = Label(f1, text="Project Name", relief=FLAT, anchor="w")
        lbl_project_name.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_fy_end = Label(f1, text="Financial Year End", relief=FLAT, anchor="w")
        lbl_fy_end.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_timing = Label(f1, text="Project Timing", relief=FLAT, anchor="w")
        lbl_timing.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_creator = Label(f1, text="Project Creator", relief=FLAT, anchor="w")
        lbl_creator.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_sector = Label(f1, text="Sector", relief=FLAT, anchor="w")
        lbl_sector.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Cancel", command=cpw.destroy).pack(side=TOP, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: right pane
        f2 = frame(cpw, LEFT)
        ipt_project_name = Entry(f2, relief=SUNKEN)
        ipt_project_name.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_fy_end = DateEntry(f2, relief=SUNKEN, year=2019, month=3, day=31)
        ipt_fy_end.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_timing = DateEntry(f2, relief=SUNKEN, year=2019, month=3, day=31)
        ipt_timing.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_creator = Entry(f2, relief=SUNKEN)
        ipt_creator.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_sector = Entry(f2, relief=SUNKEN)
        ipt_sector.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        def onSubmit(parent, master, project_name, fy_end, timing, creator, sector):
            if project_name == '' or creator == '' or sector == '':
                master.status.set("Fill required details")
                return
            #create .p file with these details
            pf=open(project_name+".p","w") #existing file will be overwritten
            pf.write("ProjectName="+project_name+"\n")
            pf.write("FY_end="+fy_end+"\n")
            pf.write("ProjectTiming="+timing+"\n")
            pf.write("ProjectCreator="+creator+"\n")
            pf.write("Sector="+sector+"\n")
            pf.close()
            master.project = master.Project(project_name, fy_end, timing, creator, sector, os.path.abspath(project_name+".p"))
            self.winfo_toplevel().title("DA Analyze: "+project_name)
            master.status.set("Project Created. Now select Tools -> Manage Data")
            parent.destroy()
        Button(f2, text="Submit", command=lambda: onSubmit(cpw, self, ipt_project_name.get(), ipt_fy_end.get_date().strftime('%d/%m/%Y'), ipt_timing.get_date().strftime('%d/%m/%Y'), ipt_creator.get(), ipt_sector.get())).pack(side=TOP, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #grab_set to refrain any activity on main window
        cpw.grab_set()

    def load_project(self):
        fname = askopenfilename(filetypes=(("Project Files", "*.p"),("All Files", "*")))
        if fname == () or fname == '': #in case of cancel or no selection
            return
        self.load_project_file(fname)

    def load_project_file(self, fname):
        f = open(fname, 'r')
        inputFileSetFlag = 1
        for line in f:
            if line[:11] == "ProjectName":
                projectName = line[12:-1]
                self.winfo_toplevel().title("DA Analyze: "+projectName)
                self.status.set("Loading Project Files...")
                self.update()
            elif line[:6] == "FY_end":
                fy_end = line[7:-1]
            elif line[:13] == "ProjectTiming":
                timing = line[14:-1]
            elif line[:14] == "ProjectCreator":
                creator = line[15:-1]
            elif line[:6] == "Sector":
                sector = line[7:-1]
            elif line[:11] == "GLinputFile":
                inputFileSetFlag = 0
                glInputFile = line[12:-1]
            elif line[:11] == "TBinputFile":
                inputFileSetFlag = 0
                tbInputFile = line[12:-1]
            elif line[:11] == "CAinputFile":
                inputFileSetFlag = 0
                caInputFile = line[12:-1]
        f.close()
        if inputFileSetFlag == 1:
            self.project = self.Project(projectName, fy_end, timing, creator, sector, fname)
            self.status.set("Loading Project File...Done. Now select Tools -> Manage Data")
            #check on garbage collection in python
        else:
            self.project = self.Project(projectName, fy_end, timing, creator, sector, fname)
            try:
                self.project.setGLInputFile(glInputFile)
                self.project.setTBInputFile(tbInputFile)
                self.project.setCAInputFile(caInputFile)
                self.status.set("Loading Project...Done")
            except IOError:
                self.status.set("Missing Data Files... Select Tools -> Manage Data; and upload data files again.")
                return
            else:
                #Display dashboard
                #self.init_dashboard()
                self.status.set("Loading Project...Done. Now select Tools -> Input Parameters")

    def makeFileMenu(self, mBar):
        CmdBtn = Menubutton(mBar, text='File', underline=0)
        CmdBtn.pack(side=LEFT, padx="2m")
        CmdBtn.menu = Menu(CmdBtn)
        CmdBtn.menu.add_command(label="Create Project...", underline=0, command=self.create_project_window)
        #CmdBtn.menu.entryconfig(0, state=DISABLED)
        CmdBtn.menu.add_command(label='Load/Open Project...', underline=5, command=self.load_project)
        CmdBtn.menu.add_command(label='Manage Projects', underline=0, state=DISABLED)#, command=manage_projects)
        CmdBtn.menu.add('separator')
        CmdBtn.menu.add_command(label='Quit', underline=0, command=CmdBtn.quit)
        CmdBtn['menu'] = CmdBtn.menu
        return CmdBtn

    def makeToolsMenu(self, mBar):
        CmdBtn = Menubutton(mBar, text='Tools', underline=0)
        CmdBtn.pack(side=LEFT, padx="2m")
        CmdBtn.menu = Menu(CmdBtn)
        CmdBtn.menu.add_command(label="Configure", underline=0, state=DISABLED)#, command=configure)
        #CmdBtn.menu.entryconfig(0, state=DISABLED)
        CmdBtn.menu.add_command(label='Manage Data', underline=6, command=self.input_data_window)
        CmdBtn.menu.add_command(label='Input Parameters', underline=0, command=self.input_parameters_window)
        CmdBtn['menu'] = CmdBtn.menu
        return CmdBtn

    def makeHelpMenu(self, mBar):
        Help_button = Menubutton(mBar, text='Help', underline=0)
        Help_button.pack(side=LEFT, padx='2m')
        Help_button["state"] = DISABLED
        return Help_button

if __name__ == '__main__':
    Application().mainloop()
