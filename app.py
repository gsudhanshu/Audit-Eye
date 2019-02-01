from Tkinter import *
import pandas as pd
from PIL import ImageTk, Image
from tkFileDialog import askopenfilename
from tkFileDialog import asksaveasfilename
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
            self.jeField = ''
            self.JEvalidated = ''
            self.sys_man_entries = ''
            self.sysField = ''
            self.sysvalues = [] #take care while loading value from project file
            self.AccDefvalidated = ''
            self.sourceInputF = ''
            self.sourceInput = None
            self.preparerInputF = ''
            self.preparerInput = None
            self.BUInputF = ''
            self.BUInput = None
            self.SG01FileName = ''
            self.SG01File = None
            self.SG02FileName = ''
            self.SG02File = None
            self.SG03FileName = ''
            self.SG03File = None
            self.SG04FileName = ''
            self.SG04File = None
            self.ip_saved = ''
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
        def setEntryEffDates(self, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt, jeField):
            self.min_entry_dt = min_entry_dt
            self.max_entry_dt = max_entry_dt
            self.min_eff_dt = min_eff_dt
            self.max_eff_dt = max_eff_dt
            self.jeField = jeField
        def setJEField(self, jeField):
            self.jeField = jeField
        def getJEField(self):
            return self.jeField
        def saveJEvalidated(self):
            self.JEvalidated = 'True'
        def setJEvalidated(self, JEvalidated):
            self.JEvalidated = JEvalidated
        def getJEvalidated(self):
            return self.JEvalidated
        def saveSys_Manual_fields(self, sys_man_entries, sysField, sysvalues):
            self.sys_man_entries = sys_man_entries
            self.sysField = sysField
            self.sysvalues = sysvalues
        def getsys_man_entries(self):
            return self.sys_man_entries
        def getsysField(self):
            return self.sysField
        def getsysValues(self):
            return self.sysvalues
        def setAccDefvalidated(self, AccDefvalidated):
            self.AccDefvalidated = AccDefvalidated
        def getAccDefvalidated(self):
            return self.AccDefvalidated
        def setSourceInputF(self, sourceFileName):
            if sourceFileName != '':
                self.sourceInput = pd.read_excel(sourceFileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_src", engine='xlsxwriter')
                self.sourceInput.to_excel(writer)
                writer.save()
                self.sourceInputF = os.path.abspath(""+self.getProjectName()+"_src")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.sourceInputF = ''
                self.sourceInput = None
        def getSourceInputF(self):
            return self.sourceInputF
        def getSourceInput(self):
            return self.sourceInput
        def setPreparerInputF(self, preparerFileName):
            #self.preparerInputF = preparerFileName
            if preparerFileName != '':
                self.preparerInput = pd.read_excel(preparerFileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_prep", engine='xlsxwriter')
                self.preparerInput.to_excel(writer)
                writer.save()
                self.preparerInputF = os.path.abspath(""+self.getProjectName()+"_prep")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.preparerInput = None
                self.preparerInputF = ''
        def getPreparerInputF(self):
            return self.preparerInputF
        def getPreparerInput(self):
            return self.preparerInput
        def setBUInputF(self, BUFileName):
            #self.BUInputF = BUFileName
            if BUFileName != '':
                self.BUInput = pd.read_excel(BUFileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_BU", engine='xlsxwriter')
                self.BUInput.to_excel(writer)
                writer.save()
                self.BUInputF = os.path.abspath(""+self.getProjectName()+"_BU")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.BUInput = None
                self.BUInputF = ''
        def getBUInputF(self):
            return self.BUInputF
        def getBUInput(self):
            return self.BUInput
        def setSegmentFiles(self, SG01FileName, SG02FileName, SG03FileName, SG04FileName):
            if SG01FileName != '':
                self.SG01File = pd.read_excel(SG01FileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG01", engine='xlsxwriter')
                self.SG01File.to_excel(writer)
                writer.save()
                self.SG01FileName = os.path.abspath(""+self.getProjectName()+"_SG01")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.SG01File = None
                self.SG01FileName = ''
            self.SG02FileName = SG02FileName
            if SG02FileName != '':
                self.SG02File = pd.read_excel(SG02FileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG02", engine='xlsxwriter')
                self.SG02File.to_excel(writer)
                writer.save()
                self.SG02FileName = os.path.abspath(""+self.getProjectName()+"_SG02")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.SG02File = None
                self.SG02FileName = ''
            if SG03FileName != '':
                self.SG03File = pd.read_excel(SG03FileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG03", engine='xlsxwriter')
                self.SG03File.to_excel(writer)
                writer.save()
                self.SG03FileName = os.path.abspath(""+self.getProjectName()+"_SG03")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.SG03File = None
                self.SG04FileName = ''
            if SG04FileName != '':
                self.SG04File = pd.read_excel(SG04FileName)
                cwd = os.getcwd()
                if cwd[:-4] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG04", engine='xlsxwriter')
                self.SG04File.to_excel(writer)
                writer.save()
                self.SG04FileName = os.path.abspath(""+self.getProjectName()+"_SG04")
                cwd = os.getcwd()
                if cwd[:-4] == 'Data':
                    os.chdir('..')
            else:
                self.SG04File = None
                self.SG04FileName = ''
        def getSG01FileName(self):
            return self.SG01FileName
        def getSG01File(self):
            return self.SG01File
        def getSG02FileName(self):
            return self.SG02FileName
        def getSG02File(self):
            return self.SG02File
        def getSG03FileName(self):
            return self.SG03FileName
        def getSG03File(self):
            return self.SG03File
        def getSG04FileName(self):
            return self.SG04FileName
        def getSG04File(self):
            return self.SG04File
        def setIPSaved(self, ip_saved):
            self.ip_saved = ip_saved
        def getIPSaved(self):
            return self.ip_saved

    def save_project_file(master):#modify to reflect latest Project Data
        pf = open(master.project.getProjectFName(), "w") #existing file will be overwritten
        pf.write("ProjectName="+master.project.getProjectName()+"\n")
        pf.write("FY_end="+master.project.getFYend()+"\n")
        pf.write("ProjectTiming="+master.project.getTiming()+"\n")
        pf.write("ProjectCreator="+master.project.getCreator()+"\n")
        pf.write("Sector="+master.project.getSector()+"\n")
        cwd = os.getcwd()
        if cwd[-4:] != "Data":
            os.chdir("Data")
        pf.write("GLinputFile="+os.path.abspath(""+master.project.getProjectName()+"_gl")+"\n")
        pf.write("TBinputFile="+os.path.abspath(""+master.project.getProjectName()+"_tb")+"\n")
        pf.write("CAinputFile="+os.path.abspath(""+master.project.getProjectName()+"_ca")+"\n")
        pf.write("JEField="+master.project.getJEField()+"\n")
        pf.write("JEvalidated="+master.project.getJEvalidated()+"\n")
        pf.write("sys_man_entries="+master.project.getsys_man_entries()+"\n")
        pf.write("sysField="+master.project.getsysField()+"\n")
        pf.write("sysValues="+str(master.project.getsysValues())+"\n")
        pf.write("AccDefValidated="+master.project.getAccDefvalidated()+"\n")
        pf.write("SourceInputF="+master.project.getSourceInputF()+"\n")
        pf.write("PreparerInputF="+master.project.getPreparerInputF()+"\n")
        pf.write("BUInputF="+master.project.getBUInputF()+"\n")
        pf.write("SG01InputF="+master.project.getSG01FileName()+"\n")
        pf.write("SG02InputF="+master.project.getSG02FileName()+"\n")
        pf.write("SG03InputF="+master.project.getSG03FileName()+"\n")
        pf.write("SG04InputF="+master.project.getSG04FileName()+"\n")
        pf.write("IPSaved="+master.project.getIPSaved()+"\n")
        cwd = os.getcwd()
        if cwd[-4:] == "Data":
            os.chdir("..")
        pf.close()

    def cleanup(master):
        master.project.setEntryEffDates('','','','','')
        master.project.setJEvalidated('')
        master.project.saveSys_Manual_fields('','',[])
        master.project.setAccDefvalidated('')
        master.project.setSourceInputF('')
        master.project.setPreparerInputF('')
        master.project.setBUInputF('')
        master.project.setSegmentFiles('','','','')
        master.project.setIPSaved('True')

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
        ipw.wm_title("Validate Input Parameters: Journal Entry Dates")
        #read gl and get max and min effective and entry dates
        glData = self.project.getGLData()
        min_entry_dt = glData['Posting Date'].min()
        max_entry_dt = glData['Posting Date'].max()
        min_eff_dt = glData['Posting Date'].min()
        max_eff_dt = glData['Posting Date'].max()
        columns = tuple(glData)
        #f1: left pane
        f1 = frame(ipw, LEFT)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_dt = Label(f1, text="Start Date:", relief=FLAT, anchor="w")
        lbl_start_dt.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_dt = Label(f1, text="End Date:", relief=FLAT, anchor="w")
        lbl_end_dt.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Center pane
        f2 = frame(ipw, LEFT)
        Label(f2, text="Entry Date", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_entry = Label(f2, text=min_entry_dt, relief=SUNKEN)
        lbl_start_entry.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_entry = Label(f2, text=max_entry_dt, relief=SUNKEN)
        lbl_end_entry.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_entry = Label(f2, text="JE No. Field", relief=FLAT)
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
        ipt_jeField = ttk.Combobox(f3, values=columns)
        ipt_jeField.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        def onOK(master):
            if ipt_jeField.get() == '':
                master.status.set("Please select JE No. Field")
                return
            else:
                master.status.set("")
            master.ipt_param_JEvalidate_window(master, ipw, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt, ipt_jeField.get())
        Button(f3, text="Ok and Next", command=lambda: onOK(self)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)

    def ipt_param_JEvalidate_window(self, master, ipw, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt, jeField):
        #save entry and effective dates in project object
        master.project.setEntryEffDates(min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt, jeField)
        ipw.destroy()
        #open next window
        ipjw = Toplevel(master)
        ipjw.wm_title("Validate Input Parameters: Journal Entries")
        glData = master.project.getGLData()
        #A. highlight more than 5 JE line items
        glData_subset = glData[[jeField, 'Amount']]
        countli_byJE = glData_subset.pivot_table(index=[jeField], aggfunc='count')
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
        amount_by_JE = glData_subset.pivot_table(index=[jeField])
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
        def onCancel(master, ipjw):
            master.cleanup()
            ipjw.destroy()
        Button(f3, text="Cancel", command=lambda: onCancel(master, ipjw)).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)        
        def onApprove(master, ipjw):
            master.project.saveJEvalidated()
            master.ipt_select_sysvalues_window(master, ipjw)
        Button(f3, text="Approve and Next", command=lambda: onApprove(master, ipjw)).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        f3.pack(expand=YES, fill=BOTH)

    def ipt_select_sysvalues_window(self, master, ipjw):
        ipjw.destroy()
        ssw = Toplevel(master)
        ssw.wm_title("Validate Input Parameters: System / Manual entries")
        #read gl and get column list
        glData = self.project.getGLData()
        columns = tuple(glData)
        #f1: left pane
        f1 = frame(ssw, LEFT)
        Label(f1, text="Journal Entry Data file contains:", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text="Select Field:", relief=FLAT, anchor="w").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text="Select System Value(s):", relief=FLAT, anchor="w").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=0, pady=0)
        def onCancel(master, ssw):
            master.cleanup()
            ssw.destroy()
        Button(f1, text="Cancel", command=lambda: onCancel(master, ssw)).pack(side=BOTTOM, padx=10, pady=10)        
        f1.pack(expand=YES, fill=BOTH)
        #f2: Center pane
        f2 = frame(ssw, LEFT)
        ipt_entries = ttk.Combobox(f2, values=("Only Manual Entries","Only System Entries","Both System and Manual Entries"))
        ipt_entries.set("Only Manual Entries")
        ipt_entries.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_sysField = ttk.Combobox(f2, values=columns)
        ipt_sysValues = Listbox(f2,selectmode='multiple')
        def sysFieldSelected(self):
            if not ipt_sysField.get() == '':
                #get unique values in column from glData
                uniqueValues = glData[ipt_sysField.get()].unique().tolist()
                for item in uniqueValues:
                    ipt_sysValues.insert(END, item)
                ssw.update()
        ipt_sysField.bind("<<ComboboxSelected>>", sysFieldSelected)
        ipt_sysField.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_sysValues.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        def onOK(master, ssw):
            selected_values = [ipt_sysValues.get(i) for i in ipt_sysValues.curselection()]
            master.project.saveSys_Manual_fields(ipt_entries.get(), ipt_sysField.get(), selected_values)
            master.ipt_acc_def_window(master, ssw)
        Button(f2, text="Ok and Next", command=lambda: onOK(master, ssw)).pack(side=BOTTOM, padx=10, pady=10)        
        f2.pack(expand=YES, fill=BOTH)

    def ipt_acc_def_window(self, master, ssw):
        ssw.destroy()
        iadw = Toplevel(master)
        iadw.wm_title("Validate Input Parameters: Account Definition")
        #read CoA
        caData = master.project.getCAData()
        #f1: Top pane
        f1 = frame(iadw, TOP)
        Label(f1, text="Review the order of all levels within the account hierarchy:", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        accTree = ttk.Treeview(f1)
        caData_subset = caData[['Account Category', 'Account Class', 'Account Subclass', 'GL Account no.']]
        df2 = pd.DataFrame({'Account Category': caData_subset['Account Category'].unique()})
        df2['Account Class'] = [list(set(caData_subset['Account Class'].loc[caData_subset['Account Category'] == x['Account Category']])) for _, x in df2.iterrows()]
        df3 = pd.DataFrame({'Account Class': caData_subset['Account Class'].unique()})
        df3['Account Subclass'] = [list(set(caData_subset['Account Subclass'].loc[caData_subset['Account Class'] == x['Account Class']])) for _, x in df3.iterrows()]
        df4 = pd.DataFrame({'Account Subclass': caData_subset['Account Subclass'].unique()})
        df4['GL Account no.'] = [list(set(caData_subset['GL Account no.'].loc[caData_subset['Account Subclass'] == x['Account Subclass']])) for _, x in df4.iterrows()]
        gi = 0
        for item in caData_subset['Account Category'].unique().tolist():
            accTree.insert('', 'end', item, text=item)
            for ite in df2['Account Class'].loc[df2['Account Category'] == item]:
                for x in ite:
                    accTree.insert(item, 'end', x, text=x)
                    for it in df3['Account Subclass'].loc[df3['Account Class'] == x]:
                        for y in it:
                            accTree.insert(x, 'end', y, text=y)
                            for i in df4['GL Account no.'].loc[df4['Account Subclass'] == y]:
                                for z in i:
                                    gl_no = str(z)[:-3]
                                    accTree.insert(y, 'end', gi, text=gl_no)
                                    gi += 1
        accTree.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Middle pane
        f2 = frame(iadw, TOP)
        def uploadNewCOA(master):
            iadw.destroy()
            master.input_data_window()
        Button(f2, text="Manage Data to upload new COA", command=lambda: uploadNewCOA(master)).pack(side=RIGHT, padx=10, pady=10)        
        def export_COA(master):
            caData = master.project.getCAData()
            savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
            writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
            caData.to_excel(writer)
            writer.save()
        Button(f2, text="Export COA", command=lambda: export_COA(master)).pack(side=RIGHT, padx=10, pady=10)        
        f2.pack(expand=YES, fill=BOTH)
        #f3: Bottom pane
        f3 = frame(iadw, BOTTOM)
        def onOK(master, iadw):
            master.project.setAccDefvalidated('True')
            master.ipt_upload_source_window(master, iadw)
        Button(f3, text="Ok and Next", command=lambda: onOK(master, iadw)).pack(side=RIGHT, padx=10, pady=10)        
        def onCancel(master, iadw):
            master.cleanup()
            iadw.destroy()
        Button(f3, text="Cancel", command=lambda: onCancel(master, iadw)).pack(side=RIGHT, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)

    def ipt_upload_source_window(self, master, iadw):
        iadw.destroy()
        iusw = Toplevel(master)
        iusw.wm_title("Validate Input Parameters: Source")
        #f1: Top pane
        f1 = frame(iusw, TOP)
        Label(f1, text="Verify that source file has following fields: Source, SourceDescription and SourceGroup", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        master.sourceFileName = StringVar()
        master.sourceFileName.set('')
        master.changeInputF = 0
        Button(f1, text="Source File...", command=lambda: browseSourceF(master)).pack(side=LEFT, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Middle pane
        f2 = frame(iusw, TOP)
        text_source = Text(f2, height=20, width=100)
        def browseSourceF(master):
            master.sourceFileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.sourceFileName.get() == (): #in case of cancel or no selection
                master.sourceFileName.set('')
                return
            master.changeInputF += 1
            master.project.setSourceInputF(master.sourceFileName.get())
            text_source.insert(END, master.project.getSourceInput()) #display dataframe in text
        source_scroll = Scrollbar(f2, command= text_source.yview)
        text_source.configure(yscrollcommand=source_scroll.set)
        text_source.pack(side=LEFT)
        source_scroll.pack(side=RIGHT, fill=Y)
        f2.pack(expand=YES, fill=BOTH)
        #f3: Bottom pane
        f3 = frame(iusw, BOTTOM)
        def onOK(master, iusw):
            master.ipt_upload_preparer_window(master, iusw)
        Button(f3, text="Ok and Next", command=lambda: onOK(master, iusw)).pack(side=RIGHT, padx=10, pady=10)        
        def onCancel(master, iusw):
            master.cleanup()
            iusw.destroy()
        Button(f3, text="Cancel", command=lambda: onCancel(master, iusw)).pack(side=RIGHT, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)
        
    def ipt_upload_preparer_window(self, master, iusw):
        iusw.destroy()
        iupw = Toplevel(master)
        iupw.wm_title("Validate Input Parameters: Source")
        #f1: Top pane
        f1 = frame(iupw, TOP)
        Label(f1, text="Verify that preparer file has following fields: UserName, FullName, Title, Department and Role", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        master.preparerFileName = StringVar()
        master.preparerFileName.set('')
        master.changeInputF = 0
        Button(f1, text="Preparer File...", command=lambda: browsePreparerF(master)).pack(side=LEFT, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Middle pane
        f2 = frame(iupw, TOP)
        text_preparer = Text(f2, height=20, width=100)
        def browsePreparerF(master):
            master.preparerFileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.preparerFileName.get() == (): #in case of cancel or no selection
                master.preparerFileName.set('')
                return
            master.changeInputF += 1
            master.project.setPreparerInputF(master.preparerFileName.get())
            text_preparer.insert(END, master.project.getPreparerInput()) #display dataframe in text
        preparer_scroll = Scrollbar(f2, command= text_preparer.yview)
        text_preparer.configure(yscrollcommand=preparer_scroll.set)
        text_preparer.pack(side=LEFT)
        preparer_scroll.pack(side=RIGHT, fill=Y)
        f2.pack(expand=YES, fill=BOTH)
        #f3: Bottom pane
        f3 = frame(iupw, BOTTOM)
        def onOK(master, iupw):
            master.ipt_upload_BU_window(master, iupw)
        Button(f3, text="Ok and Next", command=lambda: onOK(master, iupw)).pack(side=RIGHT, padx=10, pady=10)        
        def onCancel(master, iupw):
            master.cleanup()
            iupw.destroy()
        Button(f3, text="Cancel", command=lambda: onCancel(master, iupw)).pack(side=RIGHT, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)

    def ipt_upload_BU_window(self, master, iupw):
        iupw.destroy()
        iubw = Toplevel(master)
        iubw.wm_title("Validate Input Parameters: Business Unit")
        #f1: Top pane
        f1 = frame(iubw, TOP)
        Label(f1, text="Verify that business unit file has following fields: BusinessUnit, BusinessUnitDescription and BusinessUnitGroup", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        master.BUFileName = StringVar()
        master.BUFileName.set('')
        master.changeInputF = 0
        Button(f1, text="Business Unit File...", command=lambda: browseBUFile(master)).pack(side=LEFT, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Middle pane
        f2 = frame(iubw, TOP)
        text_BU = Text(f2, height=20, width=120)
        def browseBUFile(master):
            master.BUFileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.BUFileName.get() == (): #in case of cancel or no selection
                master.BUFileName.set('')
                return
            master.changeInputF += 1
            master.project.setBUInputF(master.BUFileName.get())
            text_BU.insert(END, master.project.getBUInput()) #display dataframe in text
        BU_scroll = Scrollbar(f2, command= text_BU.yview)
        text_BU.configure(yscrollcommand=BU_scroll.set)
        text_BU.pack(side=LEFT)
        BU_scroll.pack(side=RIGHT, fill=Y)
        f2.pack(expand=YES, fill=BOTH)
        #f3: Bottom pane
        f3 = frame(iubw, BOTTOM)
        def onOK(master, iubw):
            master.ipt_upload_seGments_window(master, iubw)
        Button(f3, text="Ok and Next", command=lambda: onOK(master, iubw)).pack(side=RIGHT, padx=10, pady=10)        
        def onCancel(master, iubw):
            master.cleanup()
            iubw.destroy()
        Button(f3, text="Cancel", command=lambda: onCancel(master, iubw)).pack(side=RIGHT, padx=10, pady=10)        
        f3.pack(expand=YES, fill=BOTH)

    def ipt_upload_seGments_window(self, master, iubw):
        iubw.destroy()
        iugw = Toplevel(master)
        iugw.wm_title("Validate Input Parameters: Segments")
        #f1: Top pane
        f1 = frame(iugw, TOP)
        Label(f1, text="Verify that segment files has following fields: \nSegment0x, Segment0xDescription and Segment0xGroup", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        master.SG01FileName = StringVar()
        master.SG01FileName.set('')
        master.SG02FileName = StringVar()
        master.SG02FileName.set('')
        master.SG03FileName = StringVar()
        master.SG03FileName.set('')
        master.SG04FileName = StringVar()
        master.SG04FileName.set('')
        master.SG05FileName = StringVar()
        master.SG05FileName.set('')
        master.changeInputF = 0
        f1.pack(expand=YES, fill=BOTH)
        #f2: Mid pane
        f2 = frame(iugw, TOP)
        def browseSG01File(master):
            master.SG01FileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.SG01FileName.get() == (): #in case of cancel or no selection
                master.SG01FileName.set('')
                return
            master.changeInputF += 1
        Label(f2, text="Segment01:", relief=FLAT, anchor="e").pack(side=LEFT, padx=10, pady=10)
        Button(f2, text="Segment01 File...", command=lambda: browseSG01File(master)).pack(side=LEFT, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #f3: Mid pane
        f3 = frame(iugw, TOP)
        def browseSG02File(master):
            master.SG02FileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.SG02FileName.get() == (): #in case of cancel or no selection
                master.SG02FileName.set('')
                return
            master.changeInputF += 1
        Label(f3, text="Segment02:", relief=FLAT, anchor="e").pack(side=LEFT, padx=10, pady=10)
        Button(f3, text="Segment02 File...", command=lambda: browseSG02File(master)).pack(side=LEFT, padx=10, pady=10)
        f3.pack(expand=YES, fill=BOTH)
        #f4: Mid pane
        f4 = frame(iugw, TOP)
        def browseSG03File(master):
            master.SG03FileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.SG03FileName.get() == (): #in case of cancel or no selection
                master.SG03FileName.set('')
                return
            master.changeInputF += 1
        Label(f4, text="Segment03:", relief=FLAT, anchor="e").pack(side=LEFT, padx=10, pady=10)
        Button(f4, text="Segment03 File...", command=lambda: browseSG03File(master)).pack(side=LEFT, padx=10, pady=10)
        f4.pack(expand=YES, fill=BOTH)
        #f5: Mid pane
        f5 = frame(iugw, TOP)
        def browseSG04File(master):
            master.SG04FileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"),("All Files", "*"))))
            if master.SG04FileName.get() == (): #in case of cancel or no selection
                master.SG04FileName.set('')
                return
            master.changeInputF += 1
        Label(f5, text="Segment04:", relief=FLAT, anchor="e").pack(side=LEFT, padx=10, pady=10)
        Button(f5, text="Segment04 File...", command=lambda: browseSG04File(master)).pack(side=LEFT, padx=10, pady=10)
        f5.pack(expand=YES, fill=BOTH)
        #f6: Bottom pane
        f6 = frame(iugw, BOTTOM)
        def onSave(master, iugw):
            master.project.setIPSaved('True')
            #save segment files
            if master.changeInputF > 0:
                master.project.setSegmentFiles(master.SG01FileName.get(), master.SG02FileName.get(), master.SG03FileName.get(), master.SG04FileName.get())
            #write project file
            master.save_project_file()
            iugw.destroy()
            master.init_dashboard()
        Button(f6, text="Save and Close", command=lambda: onSave(master, iugw)).pack(side=RIGHT, padx=10, pady=10)        
        def onCancel(master, iugw):
            master.cleanup()
            iugw.destroy()
        Button(f6, text="Cancel", command=lambda: onCancel(master, iugw)).pack(side=RIGHT, padx=10, pady=10)
        f6.pack(expand=YES, fill=BOTH)

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
                    master.save_project_file()
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
            elif line[:7] == "JEField":
                jeField = line[8:-1]
            elif line[:11] == "JEvalidated":
                jeValidated = line[12:-1]
            elif line[:15] == "sys_man_entries":
                sys_man_entries = line[16:-1]
            elif line[:8] == "sysField":
                sysField = line[9:-1]
            elif line[:9] == "sysValues":
                temp = line[10:-1]
                sysValues = []
                x = ''
                for char in temp:
                    if char not in ('[',']'):
                        if char == ',':
                            sysValues.append(x)
                            x = ''
                        else:
                            x += char
            elif line[:15] == "AccDefValidated":
                AccDefValidated = line[16:-1]
            elif line[:12] == "SourceInputF":
                SourceInputF = line[13:-1]
            elif line[:14] == "PreparerInputF":
                PreparerInputF = line[15:-1]
            elif line[:8] == "BUInputF":
                BUInputF = line[9:-1]
            elif line[:10] == "SG01InputF":
                SG01InputF = line[11:-1]
            elif line[:10] == "SG02InputF":
                SG02InputF = line[11:-1]
            elif line[:10] == "SG03InputF":
                SG03InputF = line[11:-1]
            elif line[:10] == "SG04InputF":
                SG04InputF = line[11:-1]
            elif line[:7] == "IPSaved":
                IPSaved = line[8:-1]
        f.close()

        if inputFileSetFlag == 1:
            self.project = self.Project(projectName, fy_end, timing, creator, sector, fname)
            self.status.set("Loading Project File...Done. Now select Tools -> Manage Data")
            #check on garbage collection in python
        elif IPSaved != 'True':
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
                self.status.set("Loading Project...Done. Now select Tools -> Input Parameters")
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
                self.project.setJEField(jeField)
                self.project.setJEvalidated(jeValidated)
                self.project.saveSys_Manual_fields(sys_man_entries, sysField, sysValues)
                self.project.setAccDefvalidated(AccDefValidated)
                self.project.setSourceInputF(SourceInputF)
                self.project.setPreparerInputF(PreparerInputF)
                self.project.setBUInputF(BUInputF)
                self.project.setSegmentFiles(SG01InputF, SG02InputF, SG03InputF, SG04InputF)
                self.project.setIPSaved(IPSaved)
                self.status.set("Loading Project...Done.")
                self.init_dashboard()#Display dashboard

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
