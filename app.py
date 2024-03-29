from Tkinter import *
import pandas as pd #v 0.23.4
from PIL import ImageTk, Image
from tkFileDialog import askopenfilename
from tkFileDialog import asksaveasfilename
from tkcalendar import DateEntry
import ttk
import os
from shutil import copyfile
import numpy as np
import unicodedata as uni
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from pandastable import Table
import calendar
from urllib2 import urlopen
import logging
from datetime import datetime

def frame(root, side):
    w=Frame(root)
    w.pack(side=side, expand=YES, fill=BOTH, padx=2, pady=2)
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
            self.tags = {}
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
            self.gldata = pd.read_excel(glInputF, skiprows=3)
        def setTBInputFile(self, tbInputF):
            self.tbInputFile = tbInputF
            self.tbdata = pd.read_excel(tbInputF, skiprows=3)
        def setCAInputFile(self, caInputF):
            self.caInputFile = caInputF
            self.cadata = pd.read_excel(caInputF, skiprows=3)
        def getGLData(self):
            return self.gldata.copy()
        def getTBData(self):
            return self.tbdata.copy()
        def getCAData(self):
            return self.cadata.copy()
        def setEntryEffDates(self, min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt):
            self.min_entry_dt = min_entry_dt
            self.max_entry_dt = max_entry_dt
            self.min_eff_dt = min_eff_dt
            self.max_eff_dt = max_eff_dt
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
        def addTag(self, jvno, tag):
            self.tags[jvno] = self.tags[jvno]+"; "+tag if jvno in list(self.tags.keys()) else tag
            self.__storeTags()
        def removeTag(self, jvno):
            del(self.tags[jvno])
            self.__storeTags()
        def setTags(self):
            cwd = os.getcwd()
            if cwd[-4:] != "Data":
                os.chdir('Data')
            tagsdata = pd.read_excel(""+self.getProjectName()+"_tags")
            for index, row in tagsdata.iterrows():
                self.tags[row['JV']] = row['Tags']
            cwd = os.getcwd()
            if cwd[-4:] == 'Data':
                os.chdir('..')
        def __storeTags(self):
            if self.tags:
                #store tags in Data folder as a excel file
                tgs = {'JV':list(self.tags.keys())}
                tagsData = pd.DataFrame.from_dict(tgs)
                tagsData['Tags'] = tagsData.apply(lambda row: self.tags[row['JV']], axis = 1)
                cwd = os.getcwd()
                if cwd[-4:] != "Data":
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_tags", engine='xlsxwriter')
                tagsData.to_excel(writer)
                writer.save()
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                cwd = os.getcwd()
                if cwd[-4:] != "Data":
                    os.chdir('Data')
                try:
                    os.remove(""+self.getProjectName()+"_tags")
                except OSError:
                    pass
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
        def getTags(self):
            return self.tags
        def setSourceInputF(self, sourceFileName):
            if sourceFileName == '':
                self.sourceInputF = ''
                self.sourceInput = None
            elif sourceFileName[-4:] != '_src':
                self.sourceInput = pd.read_excel(sourceFileName)
                cwd = os.getcwd()
                if cwd[-4:] != "Data":
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_src", engine='xlsxwriter')
                self.sourceInput.to_excel(writer)
                writer.save()
                self.sourceInputF = os.path.abspath(""+self.getProjectName()+"_src")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.sourceInputF = sourceFileName
                self.sourceInput = pd.read_excel(sourceFileName)
        def getSourceInputF(self):
            return self.sourceInputF
        def getSourceInput(self):
            return self.sourceInput.copy()
        def setPreparerInputF(self, preparerFileName):
            if preparerFileName == '':
                self.preparerInput = None
                self.preparerInputF = ''
            elif preparerFileName[-5:] != '_prep':
                self.preparerInput = pd.read_excel(preparerFileName)
                cwd = os.getcwd()
                if cwd[-4:] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_prep", engine='xlsxwriter')
                self.preparerInput.to_excel(writer)
                writer.save()
                self.preparerInputF = os.path.abspath(""+self.getProjectName()+"_prep")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.preparerInputF = preparerFileName
                self.preparerInput = pd.read_excel(preparerFileName)
        def getPreparerInputF(self):
            return self.preparerInputF
        def getPreparerInput(self):
            return self.preparerInput.copy()
        def setBUInputF(self, BUFileName):
            if BUFileName == '':
                self.BUInput = None
                self.BUInputF = ''
            elif BUFileName[-3:] != '_BU':
                self.BUInput = pd.read_excel(BUFileName)
                cwd = os.getcwd()
                if cwd[-4:] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_BU", engine='xlsxwriter')
                self.BUInput.to_excel(writer)
                writer.save()
                self.BUInputF = os.path.abspath(""+self.getProjectName()+"_BU")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.BUInput = pd.read_excel(BUFileName)
                self.BUInputF = BUFileName
        def getBUInputF(self):
            return self.BUInputF
        def getBUInput(self):
            return self.BUInput.copy()
        def setSegmentFiles(self, SG01FileName, SG02FileName, SG03FileName, SG04FileName):
            if SG01FileName == '':
                self.SG01File = None
                self.SG01FileName = ''
            elif SG01FileName[-5:] != '_SG01':
                self.SG01File = pd.read_excel(SG01FileName)
                cwd = os.getcwd()
                if cwd[-4:] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG01", engine='xlsxwriter')
                self.SG01File.to_excel(writer)
                writer.save()
                self.SG01FileName = os.path.abspath(""+self.getProjectName()+"_SG01")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.SG01File = pd.read_excel(SG01FileName)
                self.SG01FileName = SG01FileName
            if SG02FileName == '':
                self.SG02File = None
                self.SG02FileName = ''
            elif SG02FileName[-5:] != '_SG02':
                self.SG02File = pd.read_excel(SG02FileName)
                cwd = os.getcwd()
                if cwd[-4:] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG02", engine='xlsxwriter')
                self.SG02File.to_excel(writer)
                writer.save()
                self.SG02FileName = os.path.abspath(""+self.getProjectName()+"_SG02")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.SG02File = pd.read_excel(SG02FileName)
                self.SG02FileName = SG02FileName
            if SG03FileName == '':
                self.SG03File = None
                self.SG04FileName = ''
            elif SG03FileName[-5:] != '_SG03':
                self.SG03File = pd.read_excel(SG03FileName)
                cwd = os.getcwd()
                if cwd[-4:] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG03", engine='xlsxwriter')
                self.SG03File.to_excel(writer)
                writer.save()
                self.SG03FileName = os.path.abspath(""+self.getProjectName()+"_SG03")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.SG03File = pd.read_excel(SG03FileName)
                self.SG04FileName = SG03FileName
            if SG04FileName == '':
                self.SG04File = None
                self.SG04FileName = ''
            elif SG04FileName[-5:] != '_SG04':
                self.SG04File = pd.read_excel(SG04FileName)
                cwd = os.getcwd()
                if cwd[-4:] != 'Data':
                    os.chdir('Data')
                writer = pd.ExcelWriter(""+self.getProjectName()+"_SG04", engine='xlsxwriter')
                self.SG04File.to_excel(writer)
                writer.save()
                self.SG04FileName = os.path.abspath(""+self.getProjectName()+"_SG04")
                cwd = os.getcwd()
                if cwd[-4:] == 'Data':
                    os.chdir('..')
            else:
                self.SG04File = pd.read_excel(SG04FileName)
                self.SG04FileName = SG04FileName
        def getSG01FileName(self):
            return self.SG01FileName
        def getSG01File(self):
            return self.SG01File.copy()
        def getSG02FileName(self):
            return self.SG02FileName
        def getSG02File(self):
            return self.SG02File.copy()
        def getSG03FileName(self):
            return self.SG03FileName
        def getSG03File(self):
            return self.SG03File.copy()
        def getSG04FileName(self):
            return self.SG04FileName
        def getSG04File(self):
            return self.SG04File.copy()
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
        #master.project.setEntryEffDates('','','','','')
        #master.project.setJEvalidated('')
        master.project.saveSys_Manual_fields('','',[])
        master.project.setAccDefvalidated('')
        master.project.setSourceInputF('')
        master.project.setPreparerInputF('')
        master.project.setBUInputF('')
        master.project.setSegmentFiles('','','','')
        master.project.setIPSaved('')
        cwd = os.getcwd()
        if cwd[-4:] == 'Data':
            os.chdir('..')
        master.status.set("Project Input Parameters reset. To input correct parameters select Tools -> Input Parameters")

    def gross_margin_window(self):
        gmw = Toplevel(self)
        gmw.wm_title("Gross Margin Analysis")
        caData = self.project.getCAData()
        glData = self.project.getGLData()
        tbData = self.project.getTBData()
        #ftop: Top Pane
        ftop = frame(gmw, TOP)
        Label(ftop, text="Select 'Sales' Accounts:", relief=FLAT).pack(side=LEFT, padx=10, pady=10)
        Label(ftop, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmid = frame(gmw, TOP)
        f1 = frame(fmid, LEFT)
        Label(f1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCat = Listbox(f1,selectmode='multiple', exportselection=False)
        scroll_accCat = Scrollbar(f1, orient=VERTICAL, command=ipt_accCat.yview)
        ipt_accCat.config(yscrollcommand=scroll_accCat.set)
        acc_categories = caData['Account Category'].unique().tolist()
        for s in acc_categories:
            if str(s) != 'nan':
                ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCat.pack(side=RIGHT, fill=Y)
        f2 = frame(fmid, LEFT)
        Label(f2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClass = Listbox(f2,selectmode='multiple', exportselection=False)
        scroll_accClass = Scrollbar(f2, orient=VERTICAL, command=ipt_accClass.yview)
        ipt_accClass.config(yscrollcommand=scroll_accClass.set)
        ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClass.pack(side=RIGHT, fill=Y)
        f3 = frame(fmid, LEFT)
        Label(f3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclass = Listbox(f3,selectmode='multiple', exportselection=False)
        scroll_accSubclass = Scrollbar(f3, orient=VERTICAL, command=ipt_accSubclass.yview)
        ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
        ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclass.pack(side=RIGHT, fill=Y)
        f4 = frame(fmid, LEFT)
        Label(f4, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAcc = Listbox(f4,selectmode='multiple', exportselection=False)
        scroll_glAcc = Scrollbar(f4, orient=VERTICAL, command=ipt_glAcc.yview)
        ipt_glAcc.config(yscrollcommand=scroll_glAcc.set)
        ipt_glAcc.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAcc.pack(side=RIGHT, fill=Y)
        def accCatSelectionChange(evt):
            ipt_accClass.delete(0, END)
            w = evt.widget
            sel_list_accCat = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCat.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                acc_classes = tempData['Account Class'].unique().tolist()
                for s in acc_classes:
                    ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
        def accClassSelectionChange(evt):
            ipt_accSubclass.delete(0, END)
            w = evt.widget
            sel_list_accClass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                acc_subclasses = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclasses:
                    ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
        def accSubclassSelectionChange(evt):
            ipt_glAcc.delete(0, END)
            w = evt.widget
            sel_list_accSubclass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclass))]
                particulars = tempData['Particulars'].unique().tolist()
                for s in particulars:
                    ipt_glAcc.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAcc.select_set(0, END)
        ipt_accSubclass.bind('<<ListboxSelect>>', accSubclassSelectionChange)
        #Cost Accounts
        ftopB = frame(gmw, TOP)
        Label(ftopB, text="Select 'Cost' Accounts:", relief=FLAT).pack(side=LEFT, padx=10, pady=10)
        Label(ftopB, text="", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmidB = frame(gmw, TOP)
        f1B = frame(fmidB, LEFT)
        Label(f1B, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCatB = Listbox(f1B,selectmode='multiple', exportselection=False)
        scroll_accCatB = Scrollbar(f1B, orient=VERTICAL, command=ipt_accCatB.yview)
        ipt_accCatB.config(yscrollcommand=scroll_accCatB.set)
        acc_categoriesB = caData['Account Category'].unique().tolist()
        for s in acc_categoriesB:
            if str(s) != 'nan':
                ipt_accCatB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCatB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCatB.pack(side=RIGHT, fill=Y)
        f2B = frame(fmidB, LEFT)
        Label(f2B, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClassB = Listbox(f2B,selectmode='multiple', exportselection=False)
        scroll_accClassB = Scrollbar(f2B, orient=VERTICAL, command=ipt_accClassB.yview)
        ipt_accClassB.config(yscrollcommand=scroll_accClassB.set)
        ipt_accClassB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClassB.pack(side=RIGHT, fill=Y)
        f3B = frame(fmidB, LEFT)
        Label(f3B, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclassB = Listbox(f3B,selectmode='multiple', exportselection=False)
        scroll_accSubclassB = Scrollbar(f3B, orient=VERTICAL, command=ipt_accSubclassB.yview)
        ipt_accSubclassB.config(yscrollcommand=scroll_accSubclassB.set)
        ipt_accSubclassB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclassB.pack(side=RIGHT, fill=Y)
        f4B = frame(fmidB, LEFT)
        Label(f4B, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAccB = Listbox(f4B,selectmode='multiple', exportselection=False)
        scroll_glAccB = Scrollbar(f4B, orient=VERTICAL, command=ipt_glAccB.yview)
        ipt_glAccB.config(yscrollcommand=scroll_glAccB.set)
        ipt_glAccB.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAccB.pack(side=RIGHT, fill=Y)
        def accCatBSelectionChange(evt):
            ipt_accClassB.delete(0, END)
            w = evt.widget
            sel_list_accCatB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCatB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCatB))]
                acc_classesB = tempData['Account Class'].unique().tolist()
                for s in acc_classesB:
                    ipt_accClassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCatB.bind('<<ListboxSelect>>', accCatBSelectionChange)
        def accClassBSelectionChange(evt):
            ipt_accSubclassB.delete(0, END)
            w = evt.widget
            sel_list_accClassB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClassB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClassB))]
                acc_subclassesB = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclassesB:
                    ipt_accSubclassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClassB.bind('<<ListboxSelect>>', accClassBSelectionChange)
        def accSubclassBSelectionChange(evt):
            ipt_glAccB.delete(0, END)
            w = evt.widget
            sel_list_accSubclassB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclassB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclassB))]
                particularsB = tempData['Particulars'].unique().tolist()
                for s in particularsB:
                    ipt_glAccB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAccB.select_set(0, END)
        ipt_accSubclassB.bind('<<ListboxSelect>>', accSubclassBSelectionChange)
        def generateGM(master):
            sel_list_glAccA = []
            for i in ipt_glAcc.curselection():
                sel_list_glAccA.append(ipt_glAcc.get(i))
            sel_list_glAccB = []
            for i in ipt_glAccB.curselection():
                sel_list_glAccB.append(ipt_glAccB.get(i))
            if sel_list_glAccA == [] or sel_list_glAccB == []:
                master.status.set("Select Sales and Cost accounts properly!")
                return
            gw = Toplevel(gmw)
            gw.wm_title("Gross Margin Analysis")
            salesData = glData.loc[(glData["Particulars"].isin(sel_list_glAccA))]
            costData = glData.loc[(glData["Particulars"].isin(sel_list_glAccB))]
            ftop = frame(gw, TOP)
            fg1 = frame(ftop, LEFT)
            Label(fg1, text="Gross Margin %   = ", relief=FLAT, bg="white", anchor="e").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="Gross Margin Amt = ", relief=FLAT, bg="white", anchor="e").pack(side=TOP, fill=BOTH, expand=YES)
            fg2 = frame(ftop, LEFT)
            Label(fg2, text='{:,.0f}'.format((abs(salesData['Amount'].sum()) - abs(costData['Amount'].sum()))*100/abs(salesData['Amount'].sum()))+" %", relief=FLAT, bg="white", anchor="w").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text='{:,.0f}'.format(abs(salesData['Amount'].sum()) - abs(costData['Amount'].sum())), relief=FLAT, bg="white", anchor="w").pack(side=TOP, fill=BOTH, expand=YES)
            graphF = frame(gw, TOP)
            sData0 = salesData.groupby(salesData.Date.dt.to_period('M')).sum()
            sData0 = sData0.reset_index()
            cData0 = costData.groupby(costData.Date.dt.to_period('M')).sum()
            cData0 = cData0.reset_index()
            Data0 = pd.DataFrame()
            Data0['Date'] = sData0['Date']
            Data0["Gross Margin %"] = (abs(sData0['Amount']) - abs(cData0['Amount']))*100/abs(sData0['Amount'])
            Data0.set_index(['Date'], drop=True, inplace=True)
            figure = plt.Figure(figsize=(5,3), dpi=100)
            line = FigureCanvasTkAgg(figure, graphF)
            line.get_tk_widget().pack(side=TOP, fill=BOTH)
            ax1 = figure.add_subplot(111)
            Data0.plot.line(legend=True, ax=ax1)
            os.chdir('images')
            figure.savefig('myplot.png')
            os.chdir('..')
            tableF = frame(gw, TOP)
            sData0 = sData0.reset_index()
            sData0 = sData0.rename(columns = {'Amount':"Sales"})
            sData0 = sData0[['Date','Sales']]
            cData0 = cData0.reset_index()
            cData0 = cData0.rename(columns = {'Amount':"Cost"})
            cData0 = cData0[['Date','Cost']]
            Data0 = Data0.reset_index()
            Data = pd.merge(sData0, cData0, on=["Date"])
            Data["Gross Margin Amt"] = abs(Data["Sales"]) - abs(Data["Cost"])
            Data = pd.merge(Data, Data0, on=["Date"])
            i=0
            for col in tuple(Data):
                if not i == 0:
                    Data[col] = Data[col].map(master.format)
                i = i+1
            t0 = Table(tableF, dataframe=Data, width=600, height=60, showtoolbar=False, showstatusbar=False)
            t0.show()
            Label(gw, text="Sales and Cost Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            tableF1 = frame(gw, TOP)
            Data1 = salesData.groupby(['Particulars', salesData.Date.dt.to_period('M')]).sum()
            Data1 = Data1.reset_index()
            Data1 = pd.pivot_table(Data1, values='Amount', index=['Date'], columns='Particulars', aggfunc=np.sum).reset_index()
            i=0
            for col in tuple(Data1):
                if not i == 0:
                    Data1[col] = Data1[col].map(master.format)
                i = i+1
            t1 = Table(tableF1, dataframe=Data1, width=600, height=60, showtoolbar=False, showstatusbar=False)
            t1.show()
            tableF2 = frame(gw, TOP)
            Data2 = costData.groupby(['Particulars', costData.Date.dt.to_period('M')]).sum()
            Data2 = Data2.reset_index()
            Data2 = pd.pivot_table(Data2, values='Amount', index=['Date'], columns='Particulars', aggfunc=np.sum).reset_index()
            i=0
            for col in tuple(Data2):
                if not i == 0:
                    Data2[col] = Data2[col].map(master.format)
                i = i+1
            t1 = Table(tableF2, dataframe=Data2, width=600, height=60, showtoolbar=False, showstatusbar=False)
            t1.show()
            buttonF = frame(gw, BOTTOM)
            def export_to_excel(df):
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                os.chdir('images')
                worksheet.insert_image('B6', 'myplot.png')
                writer.save()
                os.chdir('..')
            Button(buttonF, text="Export to Excel", command=lambda: export_to_excel(Data)).pack(side=TOP, padx=10)
        fbot0 = frame(gmw, TOP)
        Button(fbot0, text="Generate Gross Margin Analysis", command=lambda: generateGM(self)).pack(side=TOP)
        fbot1 = frame(gmw, TOP)
        Button(fbot1, text="Done", command=gmw.destroy).pack(side=RIGHT, padx=10)
        Button(fbot1, text="Cancel", command=gmw.destroy).pack(side=RIGHT, padx=10)

    def analyze_sod(self):
        sod = Toplevel(self)
        sod.wm_title("Analyze preparers, approvers and segregation of duties")
        caData = self.project.getCAData()
        glData = self.project.getGLData()
        #ftop: Top Pane
        ftop = frame(sod, TOP)
        Label(ftop, text="Select Account Sub-Classes for analysis:", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmid = frame(sod, TOP)
        f1 = frame(fmid, LEFT)
        Label(f1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCat = Listbox(f1,selectmode='multiple', exportselection=False)
        scroll_accCat = Scrollbar(f1, orient=VERTICAL, command=ipt_accCat.yview)
        ipt_accCat.config(yscrollcommand=scroll_accCat.set)
        acc_categories = caData['Account Category'].unique().tolist()
        for s in acc_categories:
            if str(s) != 'nan':
                ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCat.pack(side=RIGHT, fill=Y)
        f2 = frame(fmid, LEFT)
        Label(f2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClass = Listbox(f2,selectmode='multiple', exportselection=False)
        scroll_accClass = Scrollbar(f2, orient=VERTICAL, command=ipt_accClass.yview)
        ipt_accClass.config(yscrollcommand=scroll_accClass.set)
        ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClass.pack(side=RIGHT, fill=Y)
        f3 = frame(fmid, LEFT)
        Label(f3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclass = Listbox(f3,selectmode='multiple', exportselection=False)
        scroll_accSubclass = Scrollbar(f3, orient=VERTICAL, command=ipt_accSubclass.yview)
        ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
        ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclass.pack(side=RIGHT, fill=Y)
        def accCatSelectionChange(evt):
            ipt_accClass.delete(0, END)
            w = evt.widget
            sel_list_accCat = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCat.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                acc_classes = tempData['Account Class'].unique().tolist()
                for s in acc_classes:
                    ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
        def accClassSelectionChange(evt):
            ipt_accSubclass.delete(0, END)
            w = evt.widget
            sel_list_accClass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                acc_subclasses = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclasses:
                    ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_accSubclass.select_set(0, END)
        ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
        fbot = frame(sod, TOP)
        def gen_sod(master):
            sel_list_accSubclass = []
            for i in ipt_accSubclass.curselection():
                sel_list_accSubclass.append(ipt_accSubclass.get(i))
            if sel_list_accSubclass == []:
                master.status.set("Select Account Sub-Classes for analysis!")
                return
            gsw = Toplevel(sod)
            gsw.wm_title("SoD Account Subclass")
            f1 = frame(gsw, TOP)
            f2 = frame(gsw, TOP)
            tempData = glData.merge(caData, on=['Particulars'])
            tempData = tempData.loc[(tempData["Account Subclass"].isin(sel_list_accSubclass))]
            tempDrData = tempData.loc[(tempData["Amount"] >= 0)]
            tempCrData = tempData.loc[(tempData["Amount"] < 0)]
            temp0Data = tempData.groupby(['Account Type', 'Account Subclass'])['Preparer'].nunique().reset_index().rename(columns={'Preparer':'CY Preparers'})
            temp1Data = tempData.groupby(['Account Type', 'Account Subclass'])['JV Number'].nunique().reset_index().rename(columns={'JV Number':'CY JE Count'})
            Data = pd.merge(temp0Data, temp1Data, on=['Account Type', 'Account Subclass'])
            Data['CY Entries per Preparer'] = Data['CY JE Count'] / Data['CY Preparers']
            temp2Data = Data
            temp3Data = tempData.groupby(['Account Type', 'Account Subclass']).sum().reset_index()
            temp3Data = temp3Data[['Account Type', 'Account Subclass', 'Amount']].rename(columns={'Amount':'CY Amount'})
            Data = pd.merge(Data, temp3Data, on=['Account Type', 'Account Subclass'])
            temp4Data = tempDrData.groupby(['Account Type', 'Account Subclass']).sum().reset_index()
            temp4Data = temp4Data[['Account Type', 'Account Subclass', 'Amount']].rename(columns={'Amount':'CY Debit'})
            Data = pd.merge(Data, temp4Data, on=['Account Type', 'Account Subclass'])
            temp5Data = tempCrData.groupby(['Account Type', 'Account Subclass']).sum().reset_index()
            temp5Data = temp5Data[['Account Type', 'Account Subclass', 'Amount']].rename(columns={'Amount':'CY Credit'})
            Data = pd.merge(Data, temp5Data, on=['Account Type', 'Account Subclass'])
            Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, pady=5)
            if Data.empty:
                Label(f1, text="No Data to Analyze!", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            else:
                i = 0
                for col in tuple(Data):
                    if i > 1:
                        Data[col] = Data[col].map(master.format)
                    i = i+1
                t = Table(f1, dataframe=Data, width=800, showtoolbar=False, showstatusbar=False)
                t.show()
                def charts(master):
                    gw = Toplevel(gsw)
                    gw.wm_title("SoD Account Subclass: Charts")
                    Data0 = temp0Data[['Account Subclass', 'CY Preparers']]
                    #Data0 = Data0.set_index(['Account Subclass'])
                    Data1 = temp1Data[['Account Subclass', 'CY JE Count']]
                    #Data1 = Data1.set_index(['Account Subclass'])
                    Data2 = temp2Data[['Account Subclass', 'CY Entries per Preparer']]
                    Data2 = Data2.set_index(['Account Subclass'])
                    graphF = frame(gw, TOP)
                    figure = plt.Figure(figsize=(5,6), dpi=100)
                    bar = FigureCanvasTkAgg(figure, graphF)
                    bar.get_tk_widget().pack(side=TOP, fill=BOTH)
                    ax1 = figure.add_subplot(411)
                    Data0.plot.bar(stacked=False, legend=True, ax=ax1)
                    ax1.xaxis.set_tick_params(rotation=0)
                    ax2 = figure.add_subplot(412)
                    Data1.plot.bar(stacked=False, legend=True, ax=ax2)
                    ax2.xaxis.set_tick_params(rotation=0)
                    ax3 = figure.add_subplot(413)
                    Data2.plot.bar(stacked=False, legend=True, ax=ax3)
                    ax3.xaxis.set_tick_params(rotation=30)
                    os.chdir('images')
                    figure.savefig('myplot.png')
                    os.chdir('..')
                    fbot = frame(gw, TOP)
                    Button(fbot, text="Export to Excel", command= gw.destroy).pack(side=TOP, padx=10, pady=5)
                    Button(fbot, text="Done", command= gw.destroy).pack(side=TOP, padx=10, pady=5)
                Button(f2, text="Graphical Presentation", command= lambda: charts(master)).pack(side=LEFT, padx=5, pady=5)
            Button(f2, text="Analyze Unexpected Preparer Relationships", bg="white", fg="RoyalBlue4", command= lambda: master.unex_relationships(sod, gsw)).pack(side=LEFT, padx=5, pady=5)
            Button(f2, text="Done", command= gsw.destroy).pack(side=LEFT, padx=5, pady=5)
            Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, pady=5)
        Button(fbot, text="Generate Report", command= lambda: gen_sod(self)).pack(side=TOP, padx=10, pady=5)
        Button(fbot, text="Done", command= sod.destroy).pack(side=TOP, padx=10, pady=5)

    def unex_relationships(master, sod, gsw):
        gsw.destroy()
        urw = Toplevel(sod)
        f1 = frame(urw, TOP)
        Label(f1, text="Analyze unexpected preparer relationships:", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f2 = frame(urw, TOP)
        f3 = frame(f2, LEFT)
        f4 = frame(f2, LEFT)
        def sod_relationship(master, x, y, prev):
            if x == '' or y == '':
                master.status.set("Set names for Primary and Secondary Accounts!")
                return
            if not prev is None:
                prev.destroy()
            srd = Toplevel(urw)
            srd.wm_title("Analyze unexpected preparer relationships: "+x+" and "+y)
            caData = master.project.getCAData()
            glData = master.project.getGLData()
            #ftop: Top Pane
            ftop = frame(srd, TOP)
            Label(ftop, text="Select Account Sub-Classes for '"+x+"':", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
            #fmid: Listboxes
            fmid = frame(srd, TOP)
            f1 = frame(fmid, LEFT)
            Label(f1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            ipt_accCat = Listbox(f1,selectmode='multiple', exportselection=False)
            scroll_accCat = Scrollbar(f1, orient=VERTICAL, command=ipt_accCat.yview)
            ipt_accCat.config(yscrollcommand=scroll_accCat.set)
            acc_categories = caData['Account Category'].unique().tolist()
            for s in acc_categories:
                if str(s) != 'nan':
                    ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
            ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
            scroll_accCat.pack(side=RIGHT, fill=Y)
            f2 = frame(fmid, LEFT)
            Label(f2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            ipt_accClass = Listbox(f2,selectmode='multiple', exportselection=False)
            scroll_accClass = Scrollbar(f2, orient=VERTICAL, command=ipt_accClass.yview)
            ipt_accClass.config(yscrollcommand=scroll_accClass.set)
            ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
            scroll_accClass.pack(side=RIGHT, fill=Y)
            f3 = frame(fmid, LEFT)
            Label(f3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            ipt_accSubclass = Listbox(f3,selectmode='multiple', exportselection=False)
            scroll_accSubclass = Scrollbar(f3, orient=VERTICAL, command=ipt_accSubclass.yview)
            ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
            ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
            scroll_accSubclass.pack(side=RIGHT, fill=Y)
            def accCatSelectionChange(evt):
                ipt_accClass.delete(0, END)
                w = evt.widget
                sel_list_accCat = []
                selected = False
                for i in w.curselection():
                    selected = True
                    sel_list_accCat.append(w.get(i))
                if selected:
                    tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                    acc_classes = tempData['Account Class'].unique().tolist()
                    for s in acc_classes:
                        ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
            ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
            def accClassSelectionChange(evt):
                ipt_accSubclass.delete(0, END)
                w = evt.widget
                sel_list_accClass = []
                selected = False
                for i in w.curselection():
                    selected = True
                    sel_list_accClass.append(w.get(i))
                if selected:
                    tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                    acc_subclasses = tempData['Account Subclass'].unique().tolist()
                    for s in acc_subclasses:
                        ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                    ipt_accSubclass.select_set(0, END)
            ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
            #Account B
            ftopB = frame(srd, TOP)
            Label(ftopB, text="Select Account Sub-Classes for '"+y+"':", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
            #fmid: Listboxes
            fmidB = frame(srd, TOP)
            f1B = frame(fmidB, LEFT)
            Label(f1B, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            ipt_accCatB = Listbox(f1B,selectmode='multiple', exportselection=False)
            scroll_accCatB = Scrollbar(f1B, orient=VERTICAL, command=ipt_accCatB.yview)
            ipt_accCatB.config(yscrollcommand=scroll_accCatB.set)
            acc_categoriesB = caData['Account Category'].unique().tolist()
            for s in acc_categoriesB:
                if str(s) != 'nan':
                    ipt_accCatB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
            ipt_accCatB.pack(side=LEFT, fill=X, expand=YES)
            scroll_accCatB.pack(side=RIGHT, fill=Y)
            f2B = frame(fmidB, LEFT)
            Label(f2B, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            ipt_accClassB = Listbox(f2B,selectmode='multiple', exportselection=False)
            scroll_accClassB = Scrollbar(f2B, orient=VERTICAL, command=ipt_accClassB.yview)
            ipt_accClassB.config(yscrollcommand=scroll_accClassB.set)
            ipt_accClassB.pack(side=LEFT, fill=X, expand=YES)
            scroll_accClassB.pack(side=RIGHT, fill=Y)
            f3B = frame(fmidB, LEFT)
            Label(f3B, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            ipt_accSubclassB = Listbox(f3B,selectmode='multiple', exportselection=False)
            scroll_accSubclassB = Scrollbar(f3B, orient=VERTICAL, command=ipt_accSubclassB.yview)
            ipt_accSubclassB.config(yscrollcommand=scroll_accSubclassB.set)
            ipt_accSubclassB.pack(side=LEFT, fill=X, expand=YES)
            scroll_accSubclassB.pack(side=RIGHT, fill=Y)
            def accCatBSelectionChange(evt):
                ipt_accClassB.delete(0, END)
                w = evt.widget
                sel_list_accCatB = []
                selected = False
                for i in w.curselection():
                    selected = True
                    sel_list_accCatB.append(w.get(i))
                if selected:
                    tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCatB))]
                    acc_classesB = tempData['Account Class'].unique().tolist()
                    for s in acc_classesB:
                        ipt_accClassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
            ipt_accCatB.bind('<<ListboxSelect>>', accCatBSelectionChange)
            def accClassBSelectionChange(evt):
                ipt_accSubclassB.delete(0, END)
                w = evt.widget
                sel_list_accClassB = []
                selected = False
                for i in w.curselection():
                    selected = True
                    sel_list_accClassB.append(w.get(i))
                if selected:
                    tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClassB))]
                    acc_subclassesB = tempData['Account Subclass'].unique().tolist()
                    for s in acc_subclassesB:
                        ipt_accSubclassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                    ipt_accSubclassB.select_set(0, END)
            ipt_accClassB.bind('<<ListboxSelect>>', accClassBSelectionChange)
            fbot = frame(srd, TOP)
            def report(master, x, y):
                sel_list_AccA = []
                for i in ipt_accSubclass.curselection():
                    sel_list_AccA.append(ipt_accSubclass.get(i))
                sel_list_AccB = []
                for i in ipt_accSubclassB.curselection():
                    sel_list_AccB.append(ipt_accSubclassB.get(i))
                if sel_list_AccA == [] or sel_list_AccB == []:
                    master.status.set("Select accounts properly!")
                    return
                rw = Toplevel(srd)
                rw.wm_title("Analyze unexpected preparer relationships: "+x+" and "+y)
                tempData = glData.merge(caData, on=['Particulars'])
                tempAData = tempData.loc[(tempData["Account Subclass"].isin(sel_list_AccA))]
                tempADrData = tempAData.loc[(tempAData["Amount"] >= 0)]
                tempACrData = tempAData.loc[(tempAData["Amount"] < 0)]
                tempBData = tempData.loc[(tempData["Account Subclass"].isin(sel_list_AccB))]
                tempBDrData = tempBData.loc[(tempBData["Amount"] >= 0)]
                tempBCrData = tempBData.loc[(tempBData["Amount"] < 0)]
                usersA = set(tempAData["Preparer"].unique().tolist())
                usersB = set(tempBData["Preparer"].unique().tolist())
                common_users = usersA & usersB
                Data = pd.DataFrame(dict([ ["No. of Preparers posting to both accounts in CY", [len(common_users)]],["Primary Account Net Amount in CY",[tempAData['Amount'].sum()]],["Primary Account Debit Amount in CY", [tempADrData['Amount'].sum()]],["Primary Account Credit Amount in CY", [tempACrData['Amount'].sum()]],["Secondary Account Net Amount in CY", [tempBData['Amount'].sum()]],["Secondary Account Debit Amount in CY", [tempBDrData['Amount'].sum()]],["Secondary Account Credit Amount in CY", [tempBCrData['Amount'].sum()]],["No. of Postings to Primary Account in CY", [len(tempAData['JV Number'].unique().tolist())]],["No. of Postings to Secondary Account in CY", [len(tempBData['JV Number'].unique().tolist())]] ]))
                ft = frame(rw, TOP)
                for col in tuple(Data):
                    Data[col] = Data[col].map(master.format)
                t = Table(ft, dataframe=Data, width=1000, showtoolbar=False, showstatusbar=False)
                t.show()
                t.setWrap()
                fbt = frame(rw, TOP)
                Label(fbt, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
                def showDetails(master):
                    col = t.getSelectedColumn()
                    row = t.getSelectedRow()
                    tempD = t.model.df
                    if str(tempD.iloc[row, col]) in ('NaN', 'nan', '', '0'):
                        return
                    jv = True
                    if tempD.columns[col] in ("Primary Account Net Amount in CY", "No. of Postings to Primary Account in CY"):
                        detailsData = tempAData
                    elif tempD.columns[col] in ("Primary Account Debit Amount in CY"):
                        detailsData = tempADrData
                    elif tempD.columns[col] in ("Primary Account Credit Amount in CY"):
                        detailsData = tempACrData
                    elif tempD.columns[col] in ("Secondary Account Net Amount in CY", "No. of Postings to Secondary Account in CY"):
                        detailsData = tempBData
                    elif tempD.columns[col] in ("Secondary Account Debit Amount in CY"):
                        detailsData = tempBDrData
                    elif tempD.columns[col] in ("Secondary Account Credit Amount in CY"):
                        detailsData = tempBCrData
                    else:
                        jv = False
                        detailsData = pd.DataFrame(dict([ ["Common Preparers", list(common_users)] ]))
                    sdw = Toplevel(rw)
                    sdw.wm_title("Analyze unexpected preparer relationships: Details")
                    if jv:
                        detailsData['Amount'] = detailsData['Amount'].map(master.format)
                    #fd1: Top pane
                    fd1 = frame(sdw, TOP)
                    detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=False, showstatusbar=False)
                    detailst.show()
                    fd2 = frame(sdw, TOP)
                    def showJVDetails(master):
                        coli = detailst.getSelectedColumn()
                        rowi = detailst.getSelectedRow()
                        tD = detailst.model.df
                        if not tD.columns[coli] == 'JV Number':
                            return
                        if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                            return
                        sjdw = Toplevel(sdw)
                        sjdw.wm_title("Analyze unexpected preparer relationships: JV Number Details")
                        jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                        jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                        fj1 = frame(sjdw, TOP)
                        pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                        pt.show()
                        fj2 = frame(sjdw, TOP)
                        def tag_jv(master, jvno):
                            tjw = Toplevel(sjdw)
                            ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                            def ok(master, jvno):
                                if ipt_tag.get() == '':
                                    master.status.set("Input Tag comment is mandatory!")
                                    return
                                master.project.addTag(jvno, "Analyze unexpected preparer relationships: "+ipt_tag.get())
                                tjw.destroy()
                            Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                            Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                        Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                        Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                    if jv:
                        Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                    def export_to_excel():
                        savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                        if savefile == '':
                            return
                        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                        detailsData.to_excel(writer, sheet_name='Sheet1')
                        writer.save()
                    Button(fd2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
                    Button(fd2, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=5)
                Button(fbt, text="Details", command=lambda: showDetails(master)).pack(side=LEFT, padx=10, pady=10)
                def export_to_excel():
                    savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                    if savefile == '':
                        return
                    writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                    Data.to_excel(writer, sheet_name='Sheet1')
                    writer.save()
                Button(fbt, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=10, pady=10)
                Button(fbt, text="Done", command=rw.destroy).pack(side=LEFT, padx=10, pady=10)
                Label(fbt, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
            Button(fbot, text="Generate Report", command=lambda: report(master, x, y)).pack(side=TOP, padx=10, pady=10)
            Button(fbot, text="Done", command= srd.destroy).pack(side=TOP, padx=10, pady=10)
        Button(f3, text="Receivables and Payables", command= lambda: sod_relationship(master, "Receivables", "Payables", None)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Cash and Revenue", command= lambda: sod_relationship(master, "Cash", "Revenue", None)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Cash and Other Income", command= lambda: sod_relationship(master, "Cash", "Other Income", None)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Cash and Cost of Sales", command= lambda: sod_relationship(master, "Cash", "Cost of Sales", None)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Sales and Cost of Sales", command= lambda: sod_relationship(master, "Sales", "Cost of Sales", None)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        def custom(master):
            cw = Toplevel(urw)
            fbot = frame(cw, BOTTOM)
            fL = frame(cw, LEFT)
            Label(fL, text="Primary account name:", relief=FLAT, anchor="e").pack(side=TOP, fill=BOTH, expand=YES, pady=10)
            Label(fL, text="Secondary account name:", relief=FLAT, anchor="e").pack(side=TOP, fill=BOTH, expand=YES, pady=10)
            fR = frame(cw, RIGHT)
            ipt_pri = Entry(fR, relief=SUNKEN, width=30)
            ipt_pri.pack(side=TOP, fill=BOTH, expand=YES, pady=10)
            ipt_sec = Entry(fR, relief=SUNKEN, width=30)
            ipt_sec.pack(side=TOP, fill=BOTH, expand=YES, pady=10)
            Button(fbot, text="Next", command= lambda: sod_relationship(master, ipt_pri.get(), ipt_sec.get(), cw)).pack(side=TOP, padx=10, pady=10)
        Button(f4, text="Custom...", command= lambda: custom(master)).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f5 = frame(urw, TOP)
        Button(f5, text="Done", command= urw.destroy).pack(side=TOP, padx=10, pady=10)

    def understand_booking_patterns(self):
        ubp = Toplevel(self)
        ubp.wm_title("Understand Booking Patterns")
        caData = self.project.getCAData()
        glData = self.project.getGLData()
        #ftop: Top Pane
        ftop = frame(ubp, TOP)
        Label(ftop, text="Select GL Accounts:", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmid = frame(ubp, TOP)
        f1 = frame(fmid, LEFT)
        Label(f1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCat = Listbox(f1,selectmode='multiple', exportselection=False)
        scroll_accCat = Scrollbar(f1, orient=VERTICAL, command=ipt_accCat.yview)
        ipt_accCat.config(yscrollcommand=scroll_accCat.set)
        acc_categories = caData['Account Category'].unique().tolist()
        for s in acc_categories:
            if str(s) != 'nan':
                ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCat.pack(side=RIGHT, fill=Y)
        f2 = frame(fmid, LEFT)
        Label(f2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClass = Listbox(f2,selectmode='multiple', exportselection=False)
        scroll_accClass = Scrollbar(f2, orient=VERTICAL, command=ipt_accClass.yview)
        ipt_accClass.config(yscrollcommand=scroll_accClass.set)
        ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClass.pack(side=RIGHT, fill=Y)
        f3 = frame(fmid, LEFT)
        Label(f3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclass = Listbox(f3,selectmode='multiple', exportselection=False)
        scroll_accSubclass = Scrollbar(f3, orient=VERTICAL, command=ipt_accSubclass.yview)
        ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
        ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclass.pack(side=RIGHT, fill=Y)
        f4 = frame(fmid, LEFT)
        Label(f4, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAcc = Listbox(f4,selectmode='multiple', exportselection=False)
        scroll_glAcc = Scrollbar(f4, orient=VERTICAL, command=ipt_glAcc.yview)
        ipt_glAcc.config(yscrollcommand=scroll_glAcc.set)
        ipt_glAcc.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAcc.pack(side=RIGHT, fill=Y)
        def accCatSelectionChange(evt):
            ipt_accClass.delete(0, END)
            w = evt.widget
            sel_list_accCat = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCat.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                acc_classes = tempData['Account Class'].unique().tolist()
                for s in acc_classes:
                    ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
        def accClassSelectionChange(evt):
            ipt_accSubclass.delete(0, END)
            w = evt.widget
            sel_list_accClass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                acc_subclasses = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclasses:
                    ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
        def accSubclassSelectionChange(evt):
            ipt_glAcc.delete(0, END)
            w = evt.widget
            sel_list_accSubclass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclass))]
                particulars = tempData['Particulars'].unique().tolist()
                for s in particulars:
                    ipt_glAcc.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAcc.select_set(0, END)
        ipt_accSubclass.bind('<<ListboxSelect>>', accSubclassSelectionChange)
        fbot0 = frame(ubp, TOP)
        fbot0_1 = frame(fbot0, TOP)
        Label(fbot0_1, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        def day_analysis(master):
            sel_list_glAcc = []
            for i in ipt_glAcc.curselection():
                sel_list_glAcc.append(ipt_glAcc.get(i))
            if sel_list_glAcc == []:
                master.status.set("Select GL Accounts!")
                return
            daw = Toplevel(ubp)
            daw.wm_title("Day Analysis - Day of Week")
            f1 = frame(daw, TOP)
            f2 = frame(daw, TOP)
            tempData = glData.loc[(glData["Particulars"].isin(sel_list_glAcc))]
            tempDrData = tempData.loc[(tempData["Amount"] >= 0)]
            tempCrData = tempData.loc[(tempData["Amount"] < 0)]
            tempData = tempData.groupby(tempData.Date.dt.weekday_name).sum()
            tempData = tempData.reset_index()
            tempData = tempData[['Date', 'Amount']]
            tempData = tempData.rename(columns = {'Date':'DayofWk', 'Amount':'NetAmount'})
            tempDrData = tempDrData.groupby(tempDrData.Date.dt.weekday_name).sum()
            tempDrData = tempDrData.reset_index()
            tempDrData = tempDrData[['Date', 'Amount']]
            tempDrData = tempDrData.rename(columns = {'Date':'DayofWk', 'Amount':'DebitAmount'})
            tempCrData = tempCrData.groupby(tempCrData.Date.dt.weekday_name).sum()
            tempCrData = tempCrData.reset_index()
            tempCrData = tempCrData[['Date', 'Amount']]
            tempCrData = tempCrData.rename(columns = {'Date':'DayofWk', 'Amount':'CreditAmount'})
            Data = pd.merge(tempDrData, tempCrData, on=['DayofWk'])
            Data = pd.merge(Data, tempData, on=['DayofWk'])
            if Data.empty:
                Label(f1, text="No Data to Analyze!", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            else:
                day = {'Monday':1, 'Tuesday':2, 'Wednesday':3, 'Thursday':4, 'Friday':5, 'Saturday':6, 'Sunday':7}
                tData = pd.DataFrame()
                tData['DayofWk'] = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday']
                tData['DebitAmount'] = [0, 0, 0, 0, 0, 0, 0]
                tData['CreditAmount'] = [0, 0, 0, 0, 0, 0, 0]
                tData['NetAmount'] = [0, 0, 0, 0, 0, 0, 0]
                Data = Data.set_index(['DayofWk'])
                tData = tData.set_index(['DayofWk'])
                Data = Data.add(tData, fill_value=0).reset_index()
                Data['SNo'] = Data.apply(lambda row: day[row['DayofWk']], axis=1)
                Data = Data.sort_values(by=['SNo'])
                Data = Data[['DayofWk', 'DebitAmount', 'CreditAmount', 'NetAmount']]
                i = 0
                for col in tuple(Data):
                    if not i == 0:
                        Data[col] = Data[col].map(master.format)
                    i = i+1
                t = Table(f1, dataframe=Data, width=400, height=140, showtoolbar=False, showstatusbar=False)
                t.show()
                def showDetails(master):
                    col = t.getSelectedColumn()
                    row = t.getSelectedRow()
                    tempD = t.model.df
                    if tempD.columns[col] in ('DayofWk') :
                        return
                    if str(tempD.iloc[row, col]) in ('NaN', 'nan', '', '0'):
                        return
                    detailsData = glData.loc[(glData['Particulars'].isin(sel_list_glAcc))]
                    if tempD.columns[col] in ('DebitAmount'):
                        detailsData = detailsData.loc[(detailsData['Amount'] >= 0)]
                    elif tempD.columns[col] in ('CreditAmount'):
                        detailsData = detailsData.loc[(detailsData['Amount'] < 0)]
                    if not 'DayofWk' in tuple(tempD):
                        ind = list(tempD.index)
                        sel_day = ind[row]
                    else:
                        sel_day = tempD['DayofWk'].iloc[row]
                    detailsData['Day'] = detailsData['Date'].dt.day_name()
                    detailsData = detailsData.loc[(detailsData['Day'] == sel_day)]
                    detailsData = detailsData.drop(columns=['Day'])
                    sdw = Toplevel(daw)
                    sdw.wm_title("Day Analysis - Day of Week: Details")
                    detailsData['Amount'] = detailsData['Amount'].map(master.format)
                    #fd1: Top pane
                    fd1 = frame(sdw, TOP)
                    detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=False, showstatusbar=False)
                    detailst.show()
                    fd2 = frame(sdw, TOP)
                    def showJVDetails(master):
                        coli = detailst.getSelectedColumn()
                        rowi = detailst.getSelectedRow()
                        tD = detailst.model.df
                        if not tD.columns[coli] == 'JV Number':
                            return
                        if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                            return
                        sjdw = Toplevel(sdw)
                        sjdw.wm_title("Day Analysis - Day of Week: JV Number Details")
                        jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                        jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                        fj1 = frame(sjdw, TOP)
                        pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                        pt.show()
                        fj2 = frame(sjdw, TOP)
                        def tag_jv(master, jvno):
                            tjw = Toplevel(sjdw)
                            ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                            def ok(master, jvno):
                                if ipt_tag.get() == '':
                                    master.status.set("Input Tag comment is mandatory!")
                                    return
                                master.project.addTag(jvno, "Day Analysis - Day of Week: "+ipt_tag.get())
                                tjw.destroy()
                            Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                            Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            return
                        Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                        Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                    Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                    def export_to_excel():
                        savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                        if savefile == '':
                            return
                        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                        detailsData.to_excel(writer, sheet_name='Sheet1')
                        writer.save()
                    Button(fd2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
                    Button(fd2, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=5)
                Button(f2, text="Details", command=lambda: showDetails(master)).pack(side=TOP, padx=10, pady=5)
                def export_to_excel():
                    savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                    if savefile == '':
                        return
                    writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                    Data.to_excel(writer, sheet_name='Sheet1')
                    writer.save()
                Button(f2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
            Button(f2, text="Done", command=daw.destroy).pack(side=TOP, padx=10, pady=5)
        Button(fbot0_1, text="Day Analysis - Day of Week", bg="white", fg="RoyalBlue4", command= lambda: day_analysis(self)).pack(side=LEFT, padx=10, pady=5)
        def day_lag_analysis(master):
            sel_list_glAcc = []
            for i in ipt_glAcc.curselection():
                sel_list_glAcc.append(ipt_glAcc.get(i))
            if sel_list_glAcc == []:
                master.status.set("Select GL Accounts!")
                return
            dlaw = Toplevel(ubp)
            dlaw.wm_title("Date Analysis - Day Lag")
            f1 = frame(dlaw, TOP)
            f2 = frame(dlaw, TOP)
            temp0Data = glData.loc[(glData["Particulars"].isin(sel_list_glAcc))]
            temp0Data['DaysLag'] = temp0Data.apply(lambda row: (row['Date'] - row['Effective Date']).days, axis=1)
            temp0Data['LineCount'] = 1
            tempData = temp0Data
            tempDrData = tempData.loc[(tempData["Amount"] >= 0)]
            tempCrData = tempData.loc[(tempData["Amount"] < 0)]
            tempData = tempData.groupby(tempData['DaysLag']).sum()
            tempData = tempData.reset_index()
            tempData = tempData[['DaysLag', 'Amount', 'LineCount']]
            tempData = tempData.rename(columns = {'Amount':'NetAmount'})
            tempDrData = tempDrData.groupby(tempDrData['DaysLag']).sum()
            tempDrData = tempDrData.reset_index()
            tempDrData = tempDrData[['DaysLag', 'Amount']]
            tempDrData = tempDrData.rename(columns = {'Amount':'DebitAmount'})
            tempCrData = tempCrData.groupby(tempCrData['DaysLag']).sum()
            tempCrData = tempCrData.reset_index()
            tempCrData = tempCrData[['DaysLag', 'Amount']]
            tempCrData = tempCrData.rename(columns = {'Amount':'CreditAmount'})
            Data = pd.merge(tempDrData, tempCrData, how='outer', on=['DaysLag'])
            Data = pd.merge(Data, tempData, how='outer', on=['DaysLag'])
            if Data.empty:
                Label(f1, text="No Data to Analyze!", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            else:
                i = 0
                for col in tuple(Data):
                    if not i == 0:
                        Data[col] = Data[col].map(master.format)
                    i = i+1
                Data = Data.sort_values(by=['DaysLag'], ascending=True)
                t = Table(f1, dataframe=Data, width=400, height=140, showtoolbar=False, showstatusbar=False)
                t.show()
                def showDetails(master):
                    col = t.getSelectedColumn()
                    row = t.getSelectedRow()
                    tempD = t.model.df
                    if tempD.columns[col] in ('DaysLag') :
                        return
                    if str(tempD.iloc[row, col]) in ('NaN', 'nan', '', '0'):
                        return
                    detailsData = temp0Data
                    if tempD.columns[col] in ('DebitAmount'):
                        detailsData = detailsData.loc[(detailsData['Amount'] >= 0)]
                    elif tempD.columns[col] in ('CreditAmount'):
                        detailsData = detailsData.loc[(detailsData['Amount'] < 0)]
                    if not 'DaysLag' in tuple(tempD):
                        ind = list(tempD.index)
                        sel_day_lag = int(ind[row])
                    else:
                        sel_day_lag = int(tempD['DaysLag'].iloc[row])
                    detailsData = detailsData.loc[(detailsData['DaysLag'] == sel_day_lag)]
                    detailsData = detailsData.drop(columns=['DaysLag', 'LineCount'])
                    sdw = Toplevel(dlaw)
                    sdw.wm_title("Date Analysis - Day Lag: Details")
                    detailsData['Amount'] = detailsData['Amount'].map(master.format)
                    #fd1: Top pane
                    fd1 = frame(sdw, TOP)
                    detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=False, showstatusbar=False)
                    detailst.show()
                    fd2 = frame(sdw, TOP)
                    def showJVDetails(master):
                        coli = detailst.getSelectedColumn()
                        rowi = detailst.getSelectedRow()
                        tD = detailst.model.df
                        if not tD.columns[coli] == 'JV Number':
                            return
                        if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                            return
                        sjdw = Toplevel(sdw)
                        sjdw.wm_title("Date Analysis - Day Lag: JV Number Details")
                        jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                        jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                        fj1 = frame(sjdw, TOP)
                        pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                        pt.show()
                        fj2 = frame(sjdw, TOP)
                        def tag_jv(master, jvno):
                            tjw = Toplevel(sjdw)
                            ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                            def ok(master, jvno):
                                if ipt_tag.get() == '':
                                    master.status.set("Input Tag comment is mandatory!")
                                    return
                                master.project.addTag(jvno, "Date Analysis - Day Lag: "+ipt_tag.get())
                                tjw.destroy()
                            Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                            Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            return
                        Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                        Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                    Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                    def export_to_excel():
                        savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                        if savefile == '':
                            return
                        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                        detailsData.to_excel(writer, sheet_name='Sheet1')
                        writer.save()
                    Button(fd2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
                    Button(fd2, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=5)
                Button(f2, text="Details", command=lambda: showDetails(master)).pack(side=TOP, padx=10, pady=5)
                def export_to_excel():
                    savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                    if savefile == '':
                        return
                    writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                    Data.to_excel(writer, sheet_name='Sheet1')
                    writer.save()
                Button(f2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
            Button(f2, text="Done", command=dlaw.destroy).pack(side=TOP, padx=10, pady=5)
        Button(fbot0_1, text="Date Analysis - Day Lag", bg="white", fg="RoyalBlue4", command=lambda: day_lag_analysis(self)).pack(side=LEFT, padx=10, pady=5)
        Label(fbot0_1, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        fbot0_2 = frame(fbot0, TOP)
        months = {1:'01.Jan', 2:'02.Feb', 3:'03.Mar', 4:'04.Apr', 5:'05.May', 6:'06.Jun', 7:'07.Jul', 8:'08.Aug', 9:'09.Sep', 10:'10.Oct', 11:'11.Nov', 12:'12.Dec'}
        def netActivity_analysis(master):
            sel_list_glAcc = []
            for i in ipt_glAcc.curselection():
                sel_list_glAcc.append(ipt_glAcc.get(i))
            if sel_list_glAcc == []:
                master.status.set("Select GL Accounts!")
                return
            naw = Toplevel(ubp)
            naw.wm_title("Net Activity Analysis by Month")
            f1 = frame(naw, TOP)
            f2 = frame(naw, TOP)
            tempData = glData.loc[(glData["Particulars"].isin(sel_list_glAcc))]
            tempData['Month'] = tempData.apply(lambda row: months[row.Date.month], axis=1)
            tempData = pd.pivot_table(tempData, values='Amount', index=tempData.Date.dt.day, columns='Month', aggfunc=np.sum).reset_index()
            if tempData.empty:
                Label(f1, text="No Data to Analyze!", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            else:
                i = 0
                for col in tuple(tempData):
                    if not i == 0:
                        tempData[col] = tempData[col].map(master.format)
                    i = i+1
                tempData = tempData.sort_values(by=['Date'], ascending=True)
                t = Table(f1, dataframe=tempData, width=700, height=620, showtoolbar=False, showstatusbar=False)
                t.show()
                Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
                def showDetails(master):
                    col = t.getSelectedColumn()
                    row = t.getSelectedRow()
                    tempD = t.model.df
                    if str(tempD.columns[col]) in ('Date') :
                        return
                    if str(tempD.iloc[row, col]) in ('NaN', 'nan', '', '0'):
                        return
                    sel_month = tempD.columns[col][:2]
                    if not 'Date' in tuple(tempD):
                        ind = list(tempD.index)
                        sel_day = ind[row]
                    else:
                        sel_day = tempD['Date'].iloc[row]
                    detailsData = glData.loc[(glData['Particulars'].isin(sel_list_glAcc)) & (glData.Date.dt.day == int(sel_day)) & (glData.Date.dt.month == int(sel_month))]
                    sdw = Toplevel(naw)
                    sdw.wm_title("Net Activity Analysis by Month: Details")
                    detailsData['Amount'] = detailsData['Amount'].map(master.format)
                    #fd1: Top pane
                    fd1 = frame(sdw, TOP)
                    detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=False, showstatusbar=False)
                    detailst.show()
                    fd2 = frame(sdw, TOP)
                    def showJVDetails(master):
                        coli = detailst.getSelectedColumn()
                        rowi = detailst.getSelectedRow()
                        tD = detailst.model.df
                        if not tD.columns[coli] == 'JV Number':
                            return
                        if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                            return
                        sjdw = Toplevel(sdw)
                        sjdw.wm_title("Net Activity Analysis by Month: JV Number Details")
                        jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                        jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                        fj1 = frame(sjdw, TOP)
                        pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                        pt.show()
                        fj2 = frame(sjdw, TOP)
                        def tag_jv(master, jvno):
                            tjw = Toplevel(sjdw)
                            ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                            def ok(master, jvno):
                                if ipt_tag.get() == '':
                                    master.status.set("Input Tag comment is mandatory!")
                                    return
                                master.project.addTag(jvno, "Net Activity Analysis by Month: "+ipt_tag.get())
                                tjw.destroy()
                            Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                            Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            return
                        Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                        Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                    Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                    def export_to_excel():
                        savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                        if savefile == '':
                            return
                        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                        detailsData.to_excel(writer, sheet_name='Sheet1')
                        writer.save()
                    Button(fd2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
                    Button(fd2, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=5)
                Button(f2, text="Details", command=lambda: showDetails(master)).pack(side=LEFT, padx=10)
                def export_to_excel():
                    savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                    if savefile == '':
                        return
                    writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                    Data.to_excel(writer, sheet_name='Sheet1')
                    writer.save()
                Button(f2, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=10)
            Button(f2, text="Done", command=naw.destroy).pack(side=LEFT, padx=10)
            Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        Button(fbot0_2, text="Net Activity Analysis by Month", bg="white", fg="RoyalBlue4", command=lambda: netActivity_analysis(self)).pack(side=LEFT, padx=10, pady=5)
        def debitActivity_analysis(master):
            sel_list_glAcc = []
            for i in ipt_glAcc.curselection():
                sel_list_glAcc.append(ipt_glAcc.get(i))
            if sel_list_glAcc == []:
                master.status.set("Select GL Accounts!")
                return
            naw = Toplevel(ubp)
            naw.wm_title("Debit Activity Analysis by Month")
            f1 = frame(naw, TOP)
            f2 = frame(naw, TOP)
            tempData = glData.loc[(glData["Particulars"].isin(sel_list_glAcc)) & (glData['Amount'] >= 0)]
            tempData['Month'] = tempData.apply(lambda row: months[row.Date.month], axis=1)
            tempData = pd.pivot_table(tempData, values='Amount', index=tempData.Date.dt.day, columns='Month', aggfunc=np.sum).reset_index()
            if tempData.empty:
                Label(f1, text="No Data to Analyze!", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            else:
                i = 0
                for col in tuple(tempData):
                    if not i == 0:
                        tempData[col] = tempData[col].map(master.format)
                    i = i+1
                tempData = tempData.sort_values(by=['Date'], ascending=True)
                t = Table(f1, dataframe=tempData, width=700, height=620, showtoolbar=False, showstatusbar=False)
                t.show()
                def showDetails(master):
                    col = t.getSelectedColumn()
                    row = t.getSelectedRow()
                    tempD = t.model.df
                    if str(tempD.columns[col]) in ('Date') :
                        return
                    if str(tempD.iloc[row, col]) in ('NaN', 'nan', '', '0'):
                        return
                    sel_month = tempD.columns[col][:2]
                    if not 'Date' in tuple(tempD):
                        ind = list(tempD.index)
                        sel_day = ind[row]
                    else:
                        sel_day = tempD['Date'].iloc[row]
                    detailsData = glData.loc[(glData['Particulars'].isin(sel_list_glAcc)) & (glData['Amount'] >= 0) & (glData.Date.dt.day == int(sel_day)) & (glData.Date.dt.month == int(sel_month))]
                    sdw = Toplevel(naw)
                    sdw.wm_title("Debit Activity Analysis by Month: Details")
                    detailsData['Amount'] = detailsData['Amount'].map(master.format)
                    #fd1: Top pane
                    fd1 = frame(sdw, TOP)
                    detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=False, showstatusbar=False)
                    detailst.show()
                    fd2 = frame(sdw, TOP)
                    def showJVDetails(master):
                        coli = detailst.getSelectedColumn()
                        rowi = detailst.getSelectedRow()
                        tD = detailst.model.df
                        if not tD.columns[coli] == 'JV Number':
                            return
                        if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                            return
                        sjdw = Toplevel(sdw)
                        sjdw.wm_title("Debit Activity Analysis by Month: JV Number Details")
                        jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                        jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                        fj1 = frame(sjdw, TOP)
                        pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                        pt.show()
                        fj2 = frame(sjdw, TOP)
                        def tag_jv(master, jvno):
                            tjw = Toplevel(sjdw)
                            ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                            def ok(master, jvno):
                                if ipt_tag.get() == '':
                                    master.status.set("Input Tag comment is mandatory!")
                                    return
                                master.project.addTag(jvno, "Debit Activity Analysis by Month: "+ipt_tag.get())
                                tjw.destroy()
                            Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                            Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            return
                        Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                        Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                    Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                    def export_to_excel():
                        savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                        if savefile == '':
                            return
                        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                        detailsData.to_excel(writer, sheet_name='Sheet1')
                        writer.save()
                    Button(fd2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
                    Button(fd2, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=5)
                Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
                Button(f2, text="Details", command=lambda: showDetails(master)).pack(side=LEFT, padx=10)
                def export_to_excel():
                    savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                    if savefile == '':
                        return
                    writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                    Data.to_excel(writer, sheet_name='Sheet1')
                    writer.save()
                Button(f2, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=10)
            Button(f2, text="Done", command=naw.destroy).pack(side=LEFT, padx=10)
            Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        Button(fbot0_2, text="Debit Activity Analysis by Month", bg="white", fg="RoyalBlue4", command=lambda: debitActivity_analysis(self)).pack(side=LEFT, padx=10, pady=5)
        def creditActivity_analysis(master):
            sel_list_glAcc = []
            for i in ipt_glAcc.curselection():
                sel_list_glAcc.append(ipt_glAcc.get(i))
            if sel_list_glAcc == []:
                master.status.set("Select GL Accounts!")
                return
            naw = Toplevel(ubp)
            naw.wm_title("Credit Activity Analysis by Month")
            f1 = frame(naw, TOP)
            f2 = frame(naw, TOP)
            tempData = glData.loc[(glData["Particulars"].isin(sel_list_glAcc)) & (glData['Amount'] < 0)]
            tempData['Month'] = tempData.apply(lambda row: months[row.Date.month], axis=1)
            tempData = pd.pivot_table(tempData, values='Amount', index=tempData.Date.dt.day, columns='Month', aggfunc=np.sum).reset_index()
            if tempData.empty:
                Label(f1, text="No Data to Analyze!", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
            else:
                i = 0
                for col in tuple(tempData):
                    if not i == 0:
                        tempData[col] = tempData[col].map(master.format)
                    i = i+1
                tempData = tempData.sort_values(by=['Date'], ascending=True)
                t = Table(f1, dataframe=tempData, width=700, height=620, showtoolbar=False, showstatusbar=False)
                t.show()
                def showDetails(master):
                    col = t.getSelectedColumn()
                    row = t.getSelectedRow()
                    tempD = t.model.df
                    if str(tempD.columns[col]) in ('Date') :
                        return
                    if str(tempD.iloc[row, col]) in ('NaN', 'nan', '', '0'):
                        return
                    sel_month = tempD.columns[col][:2]
                    if not 'Date' in tuple(tempD):
                        ind = list(tempD.index)
                        sel_day = ind[row]
                    else:
                        sel_day = tempD['Date'].iloc[row]
                    detailsData = glData.loc[(glData['Particulars'].isin(sel_list_glAcc)) & (glData['Amount'] < 0) & (glData.Date.dt.day == int(sel_day)) & (glData.Date.dt.month == int(sel_month))]
                    sdw = Toplevel(naw)
                    sdw.wm_title("Credit Activity Analysis by Month: Details")
                    detailsData['Amount'] = detailsData['Amount'].map(master.format)
                    #fd1: Top pane
                    fd1 = frame(sdw, TOP)
                    detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=False, showstatusbar=False)
                    detailst.show()
                    fd2 = frame(sdw, TOP)
                    def showJVDetails(master):
                        coli = detailst.getSelectedColumn()
                        rowi = detailst.getSelectedRow()
                        tD = detailst.model.df
                        if not tD.columns[coli] == 'JV Number':
                            return
                        if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                            return
                        sjdw = Toplevel(sdw)
                        sjdw.wm_title("Credit Activity Analysis by Month: JV Number Details")
                        jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                        jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                        fj1 = frame(sjdw, TOP)
                        pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                        pt.show()
                        fj2 = frame(sjdw, TOP)
                        def tag_jv(master, jvno):
                            tjw = Toplevel(sjdw)
                            ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                            def ok(master, jvno):
                                if ipt_tag.get() == '':
                                    master.status.set("Input Tag comment is mandatory!")
                                    return
                                master.project.addTag(jvno, "Credit Activity Analysis by Month: "+ipt_tag.get())
                                tjw.destroy()
                            Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                            Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                            return
                        Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                        Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                    Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                    def export_to_excel():
                        savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                        if savefile == '':
                            return
                        writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                        detailsData.to_excel(writer, sheet_name='Sheet1')
                        writer.save()
                    Button(fd2, text="Export to Excel", command=export_to_excel).pack(side=TOP, padx=10, pady=5)
                    Button(fd2, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=5)
                Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
                Button(f2, text="Details", command=lambda: showDetails(master)).pack(side=LEFT, padx=10)
                def export_to_excel():
                    savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                    if savefile == '':
                        return
                    writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                    Data.to_excel(writer, sheet_name='Sheet1')
                    writer.save()
                Button(f2, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=10)
            Button(f2, text="Done", command=naw.destroy).pack(side=LEFT, padx=10)
            Label(f2, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        Button(fbot0_2, text="Credit Activity Analysis by Month", bg="white", fg="RoyalBlue4", command=lambda: creditActivity_analysis(self)).pack(side=LEFT, padx=10, pady=5)
        fbot1 = frame(ubp, TOP)
        Button(fbot1, text="Done", command=ubp.destroy).pack(side=RIGHT, padx=10)
        Button(fbot1, text="Cancel", command=ubp.destroy).pack(side=RIGHT, padx=10)

    def correlation_2acc(self):
        c2aw = Toplevel(self)
        c2aw.wm_title("Correlation Analysis of 2 Accounts")
        caData = self.project.getCAData()
        glData = self.project.getGLData()
        tbData = self.project.getTBData()
        #ftop: Top Pane
        ftop = frame(c2aw, TOP)
        Label(ftop, text="Set name of Group 'A' Account:", relief=FLAT, anchor='e').pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_accA_name = Entry(ftop, relief=SUNKEN)
        ipt_accA_name.pack(side=LEFT, padx=10, pady=10)
        Label(ftop, text="", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmid = frame(c2aw, TOP)
        f1 = frame(fmid, LEFT)
        Label(f1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCat = Listbox(f1,selectmode='multiple', exportselection=False)
        scroll_accCat = Scrollbar(f1, orient=VERTICAL, command=ipt_accCat.yview)
        ipt_accCat.config(yscrollcommand=scroll_accCat.set)
        acc_categories = caData['Account Category'].unique().tolist()
        for s in acc_categories:
            if str(s) != 'nan':
                ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCat.pack(side=RIGHT, fill=Y)
        f2 = frame(fmid, LEFT)
        Label(f2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClass = Listbox(f2,selectmode='multiple', exportselection=False)
        scroll_accClass = Scrollbar(f2, orient=VERTICAL, command=ipt_accClass.yview)
        ipt_accClass.config(yscrollcommand=scroll_accClass.set)
        ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClass.pack(side=RIGHT, fill=Y)
        f3 = frame(fmid, LEFT)
        Label(f3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclass = Listbox(f3,selectmode='multiple', exportselection=False)
        scroll_accSubclass = Scrollbar(f3, orient=VERTICAL, command=ipt_accSubclass.yview)
        ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
        ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclass.pack(side=RIGHT, fill=Y)
        f4 = frame(fmid, LEFT)
        Label(f4, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAcc = Listbox(f4,selectmode='multiple', exportselection=False)
        scroll_glAcc = Scrollbar(f4, orient=VERTICAL, command=ipt_glAcc.yview)
        ipt_glAcc.config(yscrollcommand=scroll_glAcc.set)
        ipt_glAcc.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAcc.pack(side=RIGHT, fill=Y)
        def accCatSelectionChange(evt):
            ipt_accClass.delete(0, END)
            w = evt.widget
            sel_list_accCat = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCat.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                acc_classes = tempData['Account Class'].unique().tolist()
                for s in acc_classes:
                    ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
        def accClassSelectionChange(evt):
            ipt_accSubclass.delete(0, END)
            w = evt.widget
            sel_list_accClass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                acc_subclasses = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclasses:
                    ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
        def accSubclassSelectionChange(evt):
            ipt_glAcc.delete(0, END)
            w = evt.widget
            sel_list_accSubclass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclass))]
                particulars = tempData['Particulars'].unique().tolist()
                for s in particulars:
                    ipt_glAcc.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAcc.select_set(0, END)
        ipt_accSubclass.bind('<<ListboxSelect>>', accSubclassSelectionChange)
        #Account B
        ftopB = frame(c2aw, TOP)
        Label(ftopB, text="Set name of Group 'B' Account:", relief=FLAT, anchor='e').pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_accB_name = Entry(ftopB, relief=SUNKEN)
        ipt_accB_name.pack(side=LEFT, padx=10, pady=10)
        Label(ftopB, text="", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmidB = frame(c2aw, TOP)
        f1B = frame(fmidB, LEFT)
        Label(f1B, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCatB = Listbox(f1B,selectmode='multiple', exportselection=False)
        scroll_accCatB = Scrollbar(f1B, orient=VERTICAL, command=ipt_accCatB.yview)
        ipt_accCatB.config(yscrollcommand=scroll_accCatB.set)
        acc_categoriesB = caData['Account Category'].unique().tolist()
        for s in acc_categoriesB:
            if str(s) != 'nan':
                ipt_accCatB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCatB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCatB.pack(side=RIGHT, fill=Y)
        f2B = frame(fmidB, LEFT)
        Label(f2B, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClassB = Listbox(f2B,selectmode='multiple', exportselection=False)
        scroll_accClassB = Scrollbar(f2B, orient=VERTICAL, command=ipt_accClassB.yview)
        ipt_accClassB.config(yscrollcommand=scroll_accClassB.set)
        ipt_accClassB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClassB.pack(side=RIGHT, fill=Y)
        f3B = frame(fmidB, LEFT)
        Label(f3B, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclassB = Listbox(f3B,selectmode='multiple', exportselection=False)
        scroll_accSubclassB = Scrollbar(f3B, orient=VERTICAL, command=ipt_accSubclassB.yview)
        ipt_accSubclassB.config(yscrollcommand=scroll_accSubclassB.set)
        ipt_accSubclassB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclassB.pack(side=RIGHT, fill=Y)
        f4B = frame(fmidB, LEFT)
        Label(f4B, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAccB = Listbox(f4B,selectmode='multiple', exportselection=False)
        scroll_glAccB = Scrollbar(f4B, orient=VERTICAL, command=ipt_glAccB.yview)
        ipt_glAccB.config(yscrollcommand=scroll_glAccB.set)
        ipt_glAccB.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAccB.pack(side=RIGHT, fill=Y)
        def accCatBSelectionChange(evt):
            ipt_accClassB.delete(0, END)
            w = evt.widget
            sel_list_accCatB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCatB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCatB))]
                acc_classesB = tempData['Account Class'].unique().tolist()
                for s in acc_classesB:
                    ipt_accClassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCatB.bind('<<ListboxSelect>>', accCatBSelectionChange)
        def accClassBSelectionChange(evt):
            ipt_accSubclassB.delete(0, END)
            w = evt.widget
            sel_list_accClassB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClassB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClassB))]
                acc_subclassesB = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclassesB:
                    ipt_accSubclassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClassB.bind('<<ListboxSelect>>', accClassBSelectionChange)
        def accSubclassBSelectionChange(evt):
            ipt_glAccB.delete(0, END)
            w = evt.widget
            sel_list_accSubclassB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclassB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclassB))]
                particularsB = tempData['Particulars'].unique().tolist()
                for s in particularsB:
                    ipt_glAccB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAccB.select_set(0, END)
        ipt_accSubclassB.bind('<<ListboxSelect>>', accSubclassBSelectionChange)
        def generateCorrAnalysis(master):
            sel_list_glAccA = []
            for i in ipt_glAcc.curselection():
                sel_list_glAccA.append(ipt_glAcc.get(i))
            sel_list_glAccB = []
            for i in ipt_glAccB.curselection():
                sel_list_glAccB.append(ipt_glAccB.get(i))
            if sel_list_glAccA == [] or sel_list_glAccB == [] or ipt_accB_name.get() == "" or ipt_accA_name.get() == "":
                master.status.set("Select Account A and Account B, and set their names properly!")
                return
            def value(row, l, acc):
                if row['Particulars'] in l:
                    return "in account "+acc
                else:
                    return "not in account "+acc
            gw = Toplevel(master)
            gw.wm_title("Correlation Analysis")
            tbAData = tbData.loc[(tbData['Particulars'].isin(sel_list_glAccA))]
            tempAData = glData.loc[(glData["Particulars"].isin(sel_list_glAccA))]
            jvnoA = tempAData['JV Number'].unique().tolist()
            tempAData = glData.loc[(glData["JV Number"].isin(jvnoA))]
            tempAData = tempAData.loc[(~tempAData["Particulars"].isin(sel_list_glAccA))]
            tAData = tempAData
            tempAData[ipt_accA_name.get()] = tempAData.apply(lambda row: value(row, sel_list_glAccB, "B"), axis = 1)
            tempAData = tempAData.groupby([ipt_accA_name.get(), 'Date']).aggregate({'Amount': np.sum}).reset_index()
            tempAData = pd.pivot_table(tempAData, values='Amount', index=['Date'], columns=ipt_accA_name.get(), aggfunc=np.sum).reset_index()
            tAData = tAData.groupby(['Date']).aggregate({'Amount': np.sum}).reset_index()
            tbBData = tbData.loc[(tbData['Particulars'].isin(sel_list_glAccB))]
            tempBData = glData.loc[(glData["Particulars"].isin(sel_list_glAccB))]
            jvnoB = tempBData['JV Number'].unique().tolist()
            tempBData = glData.loc[(glData["JV Number"].isin(jvnoB))]
            tempBData = tempBData.loc[(~tempBData["Particulars"].isin(sel_list_glAccB))]
            tBData = tempBData
            tempBData[ipt_accB_name.get()] = tempBData.apply(lambda row: value(row, sel_list_glAccA, "A"), axis = 1)
            tempBData = tempBData.groupby([ipt_accB_name.get(), 'Date']).aggregate({'Amount': np.sum}).reset_index()
            tempBData = pd.pivot_table(tempBData, values='Amount', index=['Date'], columns=ipt_accB_name.get(), aggfunc=np.sum).reset_index()
            tBData = tBData.groupby(['Date']).aggregate({'Amount': np.sum}).reset_index()
            ftop = frame(gw, TOP)
            fg1 = frame(ftop, LEFT)
            Label(fg1, text="Composition of "+ipt_accA_name.get()+" activity (Primary) \n\n", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="Opening Balance", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="B - activity posting to "+ipt_accB_name.get()+" >>", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="A - activity not posting to "+ipt_accB_name.get()+" >>", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="Closing Balance", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg1, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            fg2 = frame(ftop, LEFT)
            Label(fg2, text="Audit Period\n"+master.project.getFYend()+"\n------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text='{:,.0f}'.format(tbAData['Opening Balance'].sum()), relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text='{:,.0f}'.format(tempAData['in account B'].sum()), relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text='{:,.0f}'.format(tempAData['not in account B'].sum()), relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text='{:,.0f}'.format(tbAData['Closing Balance'].sum()), relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg2, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            fg3 = frame(ftop, LEFT)
            Label(fg3, text=" \n \n", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg3, text="Correlation", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg3, text="difference", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg3, text='{:,.0f}'.format((abs(tempAData['in account B'].sum()) - abs(tempBData['in account A'].sum()))*100/abs(tbAData['Closing Balance'].sum()))+"%", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg3, text=" ", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg3, text=" ", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg3, text=" ", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            fg4 = frame(ftop, LEFT)
            Label(fg4, text="Composition of "+ipt_accB_name.get()+" activity (Secondary)\n\n", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg4, text="Opening Balance", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg4, text="B - activity posting to "+ipt_accA_name.get()+" >>", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg4, text="C - activity not posting to "+ipt_accA_name.get()+" >>", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg4, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg4, text="Closing Balance", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg4, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            fg5 = frame(ftop, LEFT)
            Label(fg5, text="Audit Period\n"+master.project.getFYend()+"\n------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg5, text='{:,.0f}'.format(tbBData['Opening Balance'].sum()), relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg5, text='{:,.0f}'.format(tempBData['in account A'].sum()), relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg5, text='{:,.0f}'.format(tempBData['not in account A'].sum()), relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg5, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg5, text='{:,.0f}'.format(tbBData['Closing Balance'].sum()), relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            Label(fg5, text="---------------", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
            graphF = frame(gw, TOP)
            Data0 = tempAData.groupby(tempAData.Date.dt.to_period('M')).sum()
            Data1 = tAData.groupby(tAData.Date.dt.to_period('M')).sum()
            Data1 = Data1.rename(columns = {'Amount':ipt_accA_name.get()})
            Data2 = tBData.groupby(tBData.Date.dt.to_period('M')).sum()
            Data2 = Data2.rename(columns = {'Amount':ipt_accB_name.get()})
            figure = plt.Figure(figsize=(5,3), dpi=100)
            bar = FigureCanvasTkAgg(figure, graphF)
            bar.get_tk_widget().pack(side=TOP, fill=BOTH)
            ax1 = figure.add_subplot(121)
            Data0.plot.bar(stacked=True, legend=True, ax=ax1)
            ax1.xaxis.set_tick_params(rotation=0)
            ax2 = figure.add_subplot(122)
            Data1.plot.line(legend=True, ax=ax2)
            Data2.plot.line(legend=True, ax=ax2)
            os.chdir('images')
            figure.savefig('myplot.png')
            os.chdir('..')
            tableF = frame(gw, TOP)
            Data0 = Data0.reset_index()
            Data1 = Data1.reset_index()
            Data2 = Data2.reset_index()
            Data = pd.merge(Data2, Data1, on=["Date"])
            Data = pd.merge(Data, Data0, on=["Date"])
            Data = Data.rename(columns = {'in account B':ipt_accA_name.get()+" attributed to "+ipt_accB_name.get(), 'not in account B': ipt_accA_name.get()+" not attributed to "+ipt_accB_name.get()})
            i=0
            for col in tuple(Data):
                if not i == 0:
                    Data[col] = Data[col].map(master.format)
                i = i+1
            pt = Table(tableF, dataframe=Data, width=400, height=100, showtoolbar=False, showstatusbar=False)
            pt.show()
            buttonF = frame(gw, BOTTOM)
            def export_to_excel(df):
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                os.chdir('images')
                worksheet.insert_image('B6', 'myplot.png')
                writer.save()
                os.chdir('..')
            Button(buttonF, text="Export to Excel", command=lambda: export_to_excel(Data0)).pack(side=TOP, padx=10, pady=10)
        fbot0 = frame(c2aw, TOP)
        Button(fbot0, text="Generate Correlation", command=lambda: generateCorrAnalysis(self)).pack(side=TOP)
        fbot1 = frame(c2aw, TOP)
        Button(fbot1, text="Done", command=c2aw.destroy).pack(side=RIGHT, padx=10)
        Button(fbot1, text="Cancel", command=c2aw.destroy).pack(side=RIGHT, padx=10)

    def significant_acc_window(self):
        saw = Toplevel(self)
        saw.wm_title("Significant Accounts Identification")
        tbData = self.project.getTBData()
        self.sigDatabySize = None
        self.threshold = None
        f0 = frame(saw, TOP)
        Label(f0, text="Select Significant Accounts by size", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1 = frame(saw, TOP)
        Label(f1, text="Enter Threshold Amount:", relief=FLAT, anchor='e').pack(side=LEFT, fill=BOTH, expand=YES, padx=10)
        ipt_threshold_amt = Entry(f1, relief=SUNKEN)
        ipt_threshold_amt.pack(side=LEFT, padx=10)
        Label(f1, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10)
        f2 = frame(saw, TOP)
        self.f3 = frame(saw, TOP)
        Canvas(self.f3, width=800, height=250).pack()
        def getSignificantBySize(master):
            if ipt_threshold_amt.get() == '':
                return
            try:
                master.threshold = int(ipt_threshold_amt.get())
            except Exception as e:
                master.status.set("Please input a number in threshold value!")
                return
            master.sigDatabySize = tbData.loc[(abs(tbData['Closing Balance']) >= master.threshold)]
            master.f3.destroy()
            master.f3 = frame(saw, TOP)
            st = Table(master.f3, dataframe=master.sigDatabySize, width=800, height=21, showtoolbar=True, showstatusbar=True)
            st.show()
        Button(f2, text="Get Significant Accounts", command=lambda: getSignificantBySize(self)).pack(side=TOP, padx=10)
        f4 = frame(saw, BOTTOM)
        Button(f4, text="Ok and Next", command=lambda: self.significant_byrisk_window(saw)).pack(side=RIGHT, padx=10)
        Button(f4, text="Cancel", command=saw.destroy).pack(side=RIGHT, padx=10)

    def significant_byrisk_window(self, saw):
        if self.sigDatabySize is None:
            self.status.set("Must select significant accounts by size!")
            return
        else:
            self.status.set('')
        saw.destroy()
        sbw = Toplevel(self)
        sbw.wm_title("Significant Accounts Identification")
        tbData = self.project.getTBData()
        selected = {}
        tempData = tbData.loc[(abs(tbData['Closing Balance']) < self.threshold)]
        f0 = frame(sbw, TOP)
        Label(f0, text="Select Significant Accounts by risk", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1 = frame(sbw, TOP)
        tTree = ttk.Treeview(f1)
        T_scroll = Scrollbar(f1, command= tTree.yview)
        tTree.configure(yscrollcommand=T_scroll.set)
        tTree["columns"]=("A")
        tTree.column("A", width=200)
        tTree.heading("A", text="Closing Balance")
        for index, row in tempData.iterrows():
            tTree.insert('', 'end', row['Particulars'], text=row['Particulars'], values=('{:,.0f}'.format(row['Closing Balance'])))
        tTree.pack(side=LEFT)
        T_scroll.pack(side=LEFT, fill=Y)
        ipt_accSel = Listbox(f1, selectmode='multiple', exportselection=False)
        def select(master):
            selection = tTree.selection()
            if selection == ():
                return
            elif len(selection) > 1:
                master.status.set("Select only one account at a time.")
                return
            for iid in selection:
                if iid in selected:
                    master.status.set(iid+" already selected!")
                    return
                m = Toplevel(sbw)
                Label(m, text="Document Rationale:", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                ipt_rationale = Entry(m, relief=SUNKEN, width=40)
                ipt_rationale.pack(side=TOP)
                def add(p):
                    selected[p] = ipt_rationale.get()
                    ipt_accSel.delete(0, END)
                    for d in list(selected.keys()):
                        ipt_accSel.insert(END, d)
                    m.destroy()
                Button(m, text="Ok", command=lambda: add(iid)).pack(side=TOP, padx=10)
                break
        f1_1 = frame(f1, LEFT)
        Button(f1_1, text="Add >>", command=lambda: select(self)).pack(side=TOP, padx=10, pady=10)
        def delet(evt):
            for i in ipt_accSel.curselection():
                del(selected[ipt_accSel.get(i)])
            ipt_accSel.delete(0, END)
            for item in list(selected.keys()):
                ipt_accSel.insert(END, item)            
        Button(f1_1, text="Remove --", command=lambda: delet(0)).pack(side=TOP, padx=10, pady=10)
        ipt_accSel.bind("<Delete>", delet)
        scroll_accSel = Scrollbar(f1, orient=VERTICAL, command=ipt_accSel.yview)
        ipt_accSel.config(yscrollcommand=scroll_accSel.set)
        ipt_accSel.pack(side=LEFT, fill=BOTH, expand=YES)
        scroll_accSel.pack(side=LEFT, fill=Y)
        f2 = frame(sbw, TOP)
        def save(master, savefilename):
            tsigDatabySize = master.sigDatabySize[['Particulars', 'Closing Balance']]
            tsigDatabySize['Significance'] = "By Size"
            #sigDatabyRisk = {'Particulars': list(selected.keys())}
            #tsigDatabyRisk = pd.DataFrame.from_dict(sigDatabyRisk)
            sigDatabyRisk = tbData.loc[(tbData['Particulars'].isin(list(selected.keys())))]
            tsigDatabyRisk = sigDatabyRisk[['Particulars', 'Closing Balance']]
            tsigDatabyRisk['Significance'] = tsigDatabyRisk.apply(lambda row: "By Risk - "+selected[row['Particulars']], axis=1)
            sigData = pd.concat([tsigDatabySize, tsigDatabyRisk]).reset_index()
            sigData = sigData[['Particulars', 'Closing Balance', 'Significance']]
            writer = pd.ExcelWriter(savefilename, engine='xlsxwriter')
            sigData.to_excel(writer, sheet_name='Sheet1')
            writer.save()
        def exportSignificantAccs(master):
            savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
            if savefile == '':
                return
            save(master, savefile)
        Button(f2, text="Export Significant Accounts", command=lambda: exportSignificantAccs(self)).pack(side=TOP, padx=10)
        f3 = frame(sbw, TOP)
        def dne(master, sbw):
            os.chdir("Data")
            filename = os.path.abspath(master.project.getProjectName()+"_sig")
            os.chdir("..")
            save(master, filename)
            sbw.destroy()
        Button(f3, text="Done", command=lambda: dne(self, sbw)).pack(side=RIGHT, padx=10)
        Button(f3, text="Cancel", command=sbw.destroy).pack(side=RIGHT, padx=10)

    def relation_2acc(self):
        c2aw = Toplevel(self)
        c2aw.wm_title("Relationship Analysis of 2 Accounts")
        caData = self.project.getCAData()
        glData = self.project.getGLData()
        tbData = self.project.getTBData()
        #ftop: Top Pane
        ftop = frame(c2aw, TOP)
        Label(ftop, text="Set name of Group 'A' Account:", relief=FLAT, anchor='e').pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_accA_name = Entry(ftop, relief=SUNKEN)
        ipt_accA_name.pack(side=LEFT, padx=10, pady=10)
        Label(ftop, text="", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmid = frame(c2aw, TOP)
        f1 = frame(fmid, LEFT)
        Label(f1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCat = Listbox(f1,selectmode='multiple', exportselection=False)
        scroll_accCat = Scrollbar(f1, orient=VERTICAL, command=ipt_accCat.yview)
        ipt_accCat.config(yscrollcommand=scroll_accCat.set)
        acc_categories = caData['Account Category'].unique().tolist()
        for s in acc_categories:
            if str(s) != 'nan':
                ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCat.pack(side=RIGHT, fill=Y)
        f2 = frame(fmid, LEFT)
        Label(f2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClass = Listbox(f2,selectmode='multiple', exportselection=False)
        scroll_accClass = Scrollbar(f2, orient=VERTICAL, command=ipt_accClass.yview)
        ipt_accClass.config(yscrollcommand=scroll_accClass.set)
        ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClass.pack(side=RIGHT, fill=Y)
        f3 = frame(fmid, LEFT)
        Label(f3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclass = Listbox(f3,selectmode='multiple', exportselection=False)
        scroll_accSubclass = Scrollbar(f3, orient=VERTICAL, command=ipt_accSubclass.yview)
        ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
        ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclass.pack(side=RIGHT, fill=Y)
        f4 = frame(fmid, LEFT)
        Label(f4, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAcc = Listbox(f4,selectmode='multiple', exportselection=False)
        scroll_glAcc = Scrollbar(f4, orient=VERTICAL, command=ipt_glAcc.yview)
        ipt_glAcc.config(yscrollcommand=scroll_glAcc.set)
        ipt_glAcc.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAcc.pack(side=RIGHT, fill=Y)
        def accCatSelectionChange(evt):
            ipt_accClass.delete(0, END)
            w = evt.widget
            sel_list_accCat = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCat.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                acc_classes = tempData['Account Class'].unique().tolist()
                for s in acc_classes:
                    ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
        def accClassSelectionChange(evt):
            ipt_accSubclass.delete(0, END)
            w = evt.widget
            sel_list_accClass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                acc_subclasses = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclasses:
                    ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
        def accSubclassSelectionChange(evt):
            ipt_glAcc.delete(0, END)
            w = evt.widget
            sel_list_accSubclass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclass))]
                particulars = tempData['Particulars'].unique().tolist()
                for s in particulars:
                    ipt_glAcc.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAcc.select_set(0, END)
        ipt_accSubclass.bind('<<ListboxSelect>>', accSubclassSelectionChange)
        #Account B
        ftopB = frame(c2aw, TOP)
        Label(ftopB, text="Set name of Group 'B' Account:", relief=FLAT, anchor='e').pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_accB_name = Entry(ftopB, relief=SUNKEN)
        ipt_accB_name.pack(side=LEFT, padx=10, pady=10)
        Label(ftopB, text="", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        #fmid: Listboxes
        fmidB = frame(c2aw, TOP)
        f1B = frame(fmidB, LEFT)
        Label(f1B, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCatB = Listbox(f1B,selectmode='multiple', exportselection=False)
        scroll_accCatB = Scrollbar(f1B, orient=VERTICAL, command=ipt_accCatB.yview)
        ipt_accCatB.config(yscrollcommand=scroll_accCatB.set)
        acc_categoriesB = caData['Account Category'].unique().tolist()
        for s in acc_categoriesB:
            if str(s) != 'nan':
                ipt_accCatB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCatB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCatB.pack(side=RIGHT, fill=Y)
        f2B = frame(fmidB, LEFT)
        Label(f2B, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClassB = Listbox(f2B,selectmode='multiple', exportselection=False)
        scroll_accClassB = Scrollbar(f2B, orient=VERTICAL, command=ipt_accClassB.yview)
        ipt_accClassB.config(yscrollcommand=scroll_accClassB.set)
        ipt_accClassB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClassB.pack(side=RIGHT, fill=Y)
        f3B = frame(fmidB, LEFT)
        Label(f3B, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclassB = Listbox(f3B,selectmode='multiple', exportselection=False)
        scroll_accSubclassB = Scrollbar(f3B, orient=VERTICAL, command=ipt_accSubclassB.yview)
        ipt_accSubclassB.config(yscrollcommand=scroll_accSubclassB.set)
        ipt_accSubclassB.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclassB.pack(side=RIGHT, fill=Y)
        f4B = frame(fmidB, LEFT)
        Label(f4B, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAccB = Listbox(f4B,selectmode='multiple', exportselection=False)
        scroll_glAccB = Scrollbar(f4B, orient=VERTICAL, command=ipt_glAccB.yview)
        ipt_glAccB.config(yscrollcommand=scroll_glAccB.set)
        ipt_glAccB.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAccB.pack(side=RIGHT, fill=Y)
        def accCatBSelectionChange(evt):
            ipt_accClassB.delete(0, END)
            w = evt.widget
            sel_list_accCatB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCatB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCatB))]
                acc_classesB = tempData['Account Class'].unique().tolist()
                for s in acc_classesB:
                    ipt_accClassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCatB.bind('<<ListboxSelect>>', accCatBSelectionChange)
        def accClassBSelectionChange(evt):
            ipt_accSubclassB.delete(0, END)
            w = evt.widget
            sel_list_accClassB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClassB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClassB))]
                acc_subclassesB = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclassesB:
                    ipt_accSubclassB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClassB.bind('<<ListboxSelect>>', accClassBSelectionChange)
        def accSubclassBSelectionChange(evt):
            ipt_glAccB.delete(0, END)
            w = evt.widget
            sel_list_accSubclassB = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclassB.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclassB))]
                particularsB = tempData['Particulars'].unique().tolist()
                for s in particularsB:
                    ipt_glAccB.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAccB.select_set(0, END)
        ipt_accSubclassB.bind('<<ListboxSelect>>', accSubclassBSelectionChange)
        def fetch(master):
            sel_list_glAccA = []
            for i in ipt_glAcc.curselection():
                sel_list_glAccA.append(ipt_glAcc.get(i))
            sel_list_glAccB = []
            for i in ipt_glAccB.curselection():
                sel_list_glAccB.append(ipt_glAccB.get(i))
            if sel_list_glAccA == [] or sel_list_glAccB == [] or ipt_accB_name.get() == "" or ipt_accA_name.get() == "":
                master.status.set("Select Account A and Account B and name them!")
                return
            gw = Toplevel(c2aw)
            gw.wm_title("Relationship Analysis Graph")
            graphF = frame(gw, TOP)
            accAData = glData.loc[(glData["Particulars"].isin(sel_list_glAccA))]
            Data = accAData[['Date', 'Amount']]
            Data = Data.groupby(Data.Date.dt.to_period("M")).sum()
            #Data['Amount'] = Data['Amount'].map(lambda x: abs(x))
            Data = Data.rename(columns = {'Amount':ipt_accA_name.get()})
            accBData = glData.loc[(glData["Particulars"].isin(sel_list_glAccB))]
            Data1 = accBData[['Date', 'Amount']]
            Data1 = Data1.groupby(Data1.Date.dt.to_period("M")).sum()
            #Data1['Amount'] = Data1['Amount'].map(lambda x: abs(x))
            Data1 = Data1.rename(columns = {'Amount':ipt_accB_name.get()})
            df = pd.merge(Data, Data1, on=['Date'], how="outer").reset_index()
            df = df.rename(columns = {'Date':'Month'})
            df["Primary Account as a % of Secondary Account (in %)"] = df[ipt_accA_name.get()]*100/df[ipt_accB_name.get()]
            dfg = df[['Month', "Primary Account as a % of Secondary Account (in %)"]]
            figure = plt.Figure(figsize=(5,4), dpi=100)
            ax = figure.add_subplot(111)
            line = FigureCanvasTkAgg(figure, graphF)
            line.get_tk_widget().pack(side=TOP, fill=BOTH)
            dfg.plot(kind='line', legend=True, ax=ax, color='red', marker='o', fontsize=10)
            #Data1.plot(kind='line', legend=True, ax=ax, color='blue', marker='o', fontsize=10)
            ax.set_title(ipt_accA_name.get()+' as a % of '+ipt_accB_name.get())
            os.chdir('images')
            figure.savefig('myplot.png')
            os.chdir('..')
            for col in tuple(df):
                if not col in ('Month'): 
                    df[col] = df[col].map(master.format)
            tableF = frame(gw, TOP)
            t = Table(tableF, dataframe=df, width=700, height=60, showtoolbar=False, showstatusbar=False)
            t.show()
            t.setWrap()
            buttonF = frame(gw, BOTTOM)
            def export_to_excel(df):
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                df.to_excel(writer, sheet_name='Sheet1')
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                os.chdir('images')
                worksheet.insert_image('E2', 'myplot.png')
                writer.save()
                os.chdir('..')
            Label(buttonF, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
            def showDetails(master):
                col = t.getSelectedColumn()
                row = t.getSelectedRow()
                tempD = t.model.df
                sel_gl_list = []
                if tempD.columns[col] == ipt_accA_name.get():
                    sel_gl_list = sel_list_glAccA
                elif tempD.columns[col] == ipt_accB_name.get():
                    sel_gl_list = sel_list_glAccB
                else:
                    return
                if str(tempD.iloc[row, col]) in ('NaN', 'nan', ''):
                    return
                i = 0
                for coln in tuple(tempD):
                    if coln in ('Month'):
                        period = str(tempD.iloc[row, i])
                    i = i + 1
                year = int(period[:period.find('-')])
                month = int(period[(period.find('-')+1):])
                sdw = Toplevel(gw)
                sdw.wm_title("Relationship Analysis of 2 accounts: Details")
                detailsData = glData.loc[(glData['Particulars'].isin(sel_gl_list)) & (glData.Date.dt.month == month) & (glData.Date.dt.year == year)]
                detailsData['Amount'] = detailsData['Amount'].map(master.format)
                #fd1: Top pane
                fd1 = frame(sdw, TOP)
                detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=True, showstatusbar=True)
                detailst.show()
                fd1.pack(expand=YES, fill=BOTH)
                fd2 = frame(sdw, TOP)
                def showJVDetails(master):
                    coli = detailst.getSelectedColumn()
                    rowi = detailst.getSelectedRow()
                    tD = detailst.model.df
                    if not tD.columns[coli] == 'JV Number':
                        return
                    if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                        return
                    sjdw = Toplevel(sdw)
                    sjdw.wm_title("Preparer Map Analysis: JV Number Details")
                    jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                    jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                    fj1 = frame(sjdw, TOP)
                    pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                    pt.show()
                    fj1.pack(expand=YES, fill=BOTH)            
                    fj2 = frame(sjdw, TOP)
                    def tag_jv(master, jvno):
                        tjw = Toplevel(sjdw)
                        ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                        def ok(master, jvno):
                            if ipt_tag.get() == '':
                                master.status.set("Input Tag comment is mandatory!")
                                return
                            master.project.addTag(jvno, "Process Map: "+ipt_tag.get())
                            tjw.destroy()
                        Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                        Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                        ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                        return
                    Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                    Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
                fd3 = frame(sdw, TOP)
                Button(fd3, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=10)
            Button(buttonF, text="Details", command=lambda: showDetails(master)).pack(side=LEFT, padx=5)
            Button(buttonF, text="Export to Excel", command=lambda: export_to_excel(df)).pack(side=LEFT, padx=5)
            Button(buttonF, text="Done", command=gw.destroy).pack(side=LEFT, padx=5)
            Label(buttonF, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        fbot0 = frame(c2aw, TOP)
        Button(fbot0, text="Generate Relationship Graph", command=lambda: fetch(self)).pack(side=TOP)
        fbot1 = frame(c2aw, TOP)
        Button(fbot1, text="Done", command=c2aw.destroy).pack(side=RIGHT, padx=10)
        Button(fbot1, text="Cancel", command=c2aw.destroy).pack(side=RIGHT, padx=10)

    def cutoff_analysis(self):
        caw = Toplevel(self)
        caw.wm_title("Cut-Off Analysis")
        #f1: Top pane
        f1 = frame(caw, TOP)
        Label(f1, text="Specify 'Entry Date' Range: ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        Label(f1, text="From ", relief=FLAT, anchor="e").pack(side=LEFT, padx=5)
        glData = self.project.getGLData()
        self.fetchData = glData
        from_dt = glData['Date'].min().strftime('%d/%m/%Y')
        to_dt = glData['Date'].max().strftime('%d/%m/%Y')
        #from-to Date Combobox
        ipt_from_dt = DateEntry(f1, relief=SUNKEN, year=int(from_dt[6:10]), month=int(from_dt[3:5]), day=int(from_dt[:2]))
        ipt_from_dt.pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        Label(f1, text="To ", relief=FLAT, anchor="e").pack(side=LEFT, padx=5)
        ipt_to_dt = DateEntry(f1, relief=SUNKEN, year=int(to_dt[6:10]), month=int(to_dt[3:5]), day=int(to_dt[:2]))
        ipt_to_dt.pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        f1.pack(expand=YES, fill=BOTH)
        #fmid: Mid pane
        fmid = frame(caw, TOP)
        caData = self.project.getCAData()
        fmid_1 = frame(fmid, LEFT)
        Label(fmid_1, text="Account Category", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accCat = Listbox(fmid_1,selectmode='multiple', exportselection=False)
        scroll_accCat = Scrollbar(fmid_1, orient=VERTICAL, command=ipt_accCat.yview)
        ipt_accCat.config(yscrollcommand=scroll_accCat.set)
        acc_categories = caData['Account Category'].unique().tolist()
        for s in acc_categories:
            if str(s) != 'nan':
                ipt_accCat.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.pack(side=LEFT, fill=X, expand=YES)
        scroll_accCat.pack(side=RIGHT, fill=Y)
        fmid_2 = frame(fmid, LEFT)
        Label(fmid_2, text="Account Class", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accClass = Listbox(fmid_2,selectmode='multiple', exportselection=False)
        scroll_accClass = Scrollbar(fmid_2, orient=VERTICAL, command=ipt_accClass.yview)
        ipt_accClass.config(yscrollcommand=scroll_accClass.set)
        ipt_accClass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accClass.pack(side=RIGHT, fill=Y)
        fmid_3 = frame(fmid, LEFT)
        Label(fmid_3, text="Account Subclass", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_accSubclass = Listbox(fmid_3,selectmode='multiple', exportselection=False)
        scroll_accSubclass = Scrollbar(fmid_3, orient=VERTICAL, command=ipt_accSubclass.yview)
        ipt_accSubclass.config(yscrollcommand=scroll_accSubclass.set)
        ipt_accSubclass.pack(side=LEFT, fill=X, expand=YES)
        scroll_accSubclass.pack(side=RIGHT, fill=Y)
        fmid_4 = frame(fmid, LEFT)
        Label(fmid_4, text="GL Accounts", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES)
        ipt_glAcc = Listbox(fmid_4,selectmode='multiple', exportselection=False)
        scroll_glAcc = Scrollbar(fmid_4, orient=VERTICAL, command=ipt_glAcc.yview)
        ipt_glAcc.config(yscrollcommand=scroll_glAcc.set)
        ipt_glAcc.pack(side=LEFT, fill=X, expand=YES)
        scroll_glAcc.pack(side=RIGHT, fill=Y)
        def accCatSelectionChange(evt):
            ipt_accClass.delete(0, END)
            w = evt.widget
            sel_list_accCat = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accCat.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Category"].isin(sel_list_accCat))]
                acc_classes = tempData['Account Class'].unique().tolist()
                for s in acc_classes:
                    ipt_accClass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accCat.bind('<<ListboxSelect>>', accCatSelectionChange)
        def accClassSelectionChange(evt):
            ipt_accSubclass.delete(0, END)
            w = evt.widget
            sel_list_accClass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accClass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Class"].isin(sel_list_accClass))]
                acc_subclasses = tempData['Account Subclass'].unique().tolist()
                for s in acc_subclasses:
                    ipt_accSubclass.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
        ipt_accClass.bind('<<ListboxSelect>>', accClassSelectionChange)
        def accSubclassSelectionChange(evt):
            ipt_glAcc.delete(0, END)
            w = evt.widget
            sel_list_accSubclass = []
            selected = False
            for i in w.curselection():
                selected = True
                sel_list_accSubclass.append(w.get(i))
            if selected:
                tempData = caData.loc[(caData["Account Subclass"].isin(sel_list_accSubclass))]
                particulars = tempData['Particulars'].unique().tolist()
                for s in particulars:
                    ipt_glAcc.insert(END, uni.normalize('NFKD', s).encode('ascii','ignore'))
                ipt_glAcc.select_set(0, END)
        ipt_accSubclass.bind('<<ListboxSelect>>', accSubclassSelectionChange)
        #f3: Mid pane
        f3 = frame(caw, TOP)
        Label(f3, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10)
        Label(f3, text="Threshold Amount:", relief=FLAT).pack(side=LEFT, padx=5)
        ipt_thresholdAmt = Entry(f3, relief=SUNKEN)
        ipt_thresholdAmt.pack(side=LEFT, padx=5)       
        Label(f3, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10)
        #f4: Mid pane
        f4 = frame(caw, TOP)
        f5 = frame(caw, TOP)
        self.pt = Label(f5, text="", relief=FLAT)
        self.pt.pack(side=LEFT, fill=BOTH, expand=YES, padx=10)
        def fetch(master):
            master.fetchData = glData.loc[(glData["Date"] >= pd.Timestamp(ipt_from_dt.get_date())) & (glData["Date"] <= pd.Timestamp(ipt_to_dt.get_date()))]
            sel_list_glAcc = []
            for i in ipt_glAcc.curselection():
                sel_list_glAcc.append(ipt_glAcc.get(i))
            master.fetchData = master.fetchData.loc[(master.fetchData["Particulars"].isin(sel_list_glAcc))]
            master.fetchData = master.fetchData.loc[(abs(master.fetchData["Amount"]) >= int(ipt_thresholdAmt.get()))]
            master.pt.destroy()
            master.pt = Table(f5, dataframe=master.fetchData, width=700, showtoolbar=False, showstatusbar=False)
            master.pt.show()
            master.details.config(state="normal")
        Button(f4, text="Generate", command=lambda: fetch(self)).pack(side=TOP, padx=2, pady=2)
        #f6: Bottom pane
        f6 = frame(caw, BOTTOM)
        Label(f6, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        def showJVDetails(master):
            coli = master.pt.getSelectedColumn()
            rowi = master.pt.getSelectedRow()
            tD = master.pt.model.df
            if not tD.columns[coli] == 'JV Number':
                return
            if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                return
            sjdw = Toplevel(caw)
            sjdw.wm_title("Cut-off Analysis: JV Number Details")
            jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
            jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
            fj1 = frame(sjdw, TOP)
            pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
            pt.show()
            fj2 = frame(sjdw, TOP)
            def tag_jv(master, jvno):
                tjw = Toplevel(sjdw)
                ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                def ok(master, jvno):
                    if ipt_tag.get() == '':
                        master.status.set("Input Tag comment is mandatory!")
                        return
                    master.project.addTag(jvno, "Cut-off Analysis: "+ipt_tag.get())
                    tjw.destroy()
                Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                return
            Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
            Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
        self.details = Button(f6, text="Details", command=lambda: showJVDetails(self), state=DISABLED)
        self.details.pack(side=LEFT, padx=5)
        def export_to_excel(master):
            savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
            if savefile == '':
                return
            writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
            master.fetchData.to_excel(writer)
            writer.save()            
        Button(f6, text="Export to Excel", command=lambda: export_to_excel(self)).pack(side=LEFT, padx=5)
        Button(f6, text="Done", command=caw.destroy).pack(side=LEFT, padx=5)
        Label(f6, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10)

    def init_dashboard(self):
        self.l1.destroy()
        self.f0.destroy()
        self.f0 = frame(self.w, TOP)
        self.status.set("")
        #f1: left pane
        f1 = frame(self.f0, LEFT)
        Label(f1, text="Financial Statement Profiling", bg="SkyBlue4", fg="white", font='Helvetica 12 bold').pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Analyze Balance Sheet", bg="white", fg="RoyalBlue4", command=self.balance_sheet_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Analyze Income Statement", bg="white", fg="RoyalBlue4", command= self.income_statement_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f1, text="Business Unit Map", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f1, text="Financial Statement Tie-out", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f1, text="Significant Accounts Identification", bg="white", fg="RoyalBlue4", command= self.significant_acc_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f1, text="Income Analysis", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: second pane
        f2 = frame(self.f0, LEFT)
        Label(f2, text="Validation", bg="SkyBlue4", fg="white", font='Helvetica 12 bold').pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f2, text="JE Validation", bg="white", fg="RoyalBlue4", command=self.JEvalidate_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f2, text="Date Validation", bg="white", fg="RoyalBlue4", command=self.date_validation_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f2, text="Trial Balance Validation", bg="white", fg="RoyalBlue4", command=self.tb_validation_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f2, text="Validation Results Summary", command=self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Label(f2, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f2, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f2, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #f3: third pane
        f3 = frame(self.f0, LEFT)
        Label(f3, text="Process Analysis", bg="SkyBlue4", fg="white", font='Helvetica 12 bold').pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Process Map", bg="white", fg="RoyalBlue4", command=self.process_map_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f3, text="Preparer Map", bg="white", fg="RoyalBlue4", command=self.preparer_map_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Analyze preparers, approvers and segregation of duties", bg="white", fg="RoyalBlue4", command= self.analyze_sod).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Identify and Understand Booking Patterns", bg="white", fg="RoyalBlue4", command= self.understand_booking_patterns).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f3, text="Tagging Analysis - Journals", bg="white", fg="RoyalBlue4", command= self.tagging_analysis_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Label(f3, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f3, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f3.pack(expand=YES, fill=BOTH)
        #f4: Last pane
        f4 = frame(self.f0, LEFT)
        Label(f4, text="Account and Journal Entry Analysis", bg="SkyBlue4", fg="white", font='Helvetica 12 bold').pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Analyze Correlation b/w 2 accounts", bg="white", fg="RoyalBlue4", command= self.correlation_2acc).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Analyze Correlation b/w 3 accounts", command= self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Button(f4, text="Analyze Relationship of 2 accounts", bg="white", fg="RoyalBlue4", command= self.relation_2acc).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Gross Margin Analysis", bg="white", fg="RoyalBlue4", command= self.gross_margin_window).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Cutoff Analysis of GL accounts", bg="white", fg="RoyalBlue4", command= self.cutoff_analysis).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Additional Reports", command= self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        Button(f4, text="Custom Analytics - visualization", command= self.destroy).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)        
        f4.pack(expand=YES, fill=BOTH)

    def tagging_analysis_window(master):
        taw = Toplevel(master)
        taw.wm_title("Tagging Analysis - Journals")
        tgs = master.project.getTags()
        glData = master.project.getGLData()
        f0 = frame(taw, TOP)
        Label(f0, text="JVs selected for sampling:", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1 = frame(taw, TOP)
        tTree = ttk.Treeview(f1)
        T_scroll = Scrollbar(f1, command= tTree.yview)
        tTree.configure(yscrollcommand=T_scroll.set)
        tTree["columns"]=("A")
        tTree.column("A", width=200)
        tTree.heading("A", text="Rationale")
        for jvno in list(tgs.keys()):
            tTree.insert('', 'end', jvno, text=jvno, values=[tgs[jvno]])
        tTree.pack(side=LEFT)
        T_scroll.pack(side=LEFT, fill=Y)
        fmid = frame(taw, TOP)
        Label(fmid, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        def showJVDetails(master):
            selection = tTree.selection()
            if selection == (): #no selection
                return
            if len(selection) > 1: #more than one selection
                master.status.set("Select one JV at a time for details")
                return
            sjdw = Toplevel(taw)
            sjdw.wm_title("Tagging Analysis: JV Number Details")
            jvdetailsData = glData.loc[(glData['JV Number'] == int(selection[0]))]
            jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
            fj1 = frame(sjdw, TOP)
            pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
            pt.show()
            fj2 = frame(sjdw, TOP)
            Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
        Button(fmid, text="Details", command=lambda: showJVDetails(master)).pack(side=LEFT, padx=10, pady=10)        
        def remTag(master):
            selection = tTree.selection()
            if selection == (): #no selection
                return
            for jv in selection:
                master.project.removeTag(int(jv))
                tTree.delete(jv)
        Button(fmid, text="Remove", command=lambda: remTag(master)).pack(side=LEFT, padx=10, pady=10)        
        Label(fmid, text=" ", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES)
        Button(taw, text="Done", command=taw.destroy).pack(side=TOP, padx=10, pady=10)        

    def tb_validation_window(master):
        tvw = Toplevel(master)
        tvw.wm_title("Trial Balance Validation: Total of Trial Balance")
        tbData = master.project.getTBData()
        total = tbData['Closing Balance'].sum()
        tbData['Opening Balance'] = tbData['Opening Balance'].map(master.format)
        tbData['Debit'] = tbData['Debit'].map(master.format)
        tbData['Credit'] = tbData['Credit'].map(master.format)
        tbData['Closing Balance'] = tbData['Closing Balance'].map(master.format)
        ftop = frame(tvw, TOP)
        tbt = Table(ftop, dataframe=tbData, width=800, height=21, showtoolbar=True, showstatusbar=True)
        tbt.show()
        fmid = frame(tvw, TOP)
        Label(fmid, text="Total of Trial Balance: "+'{:,.0f}'.format(total), font='Helvetica 12 bold').pack(side=TOP, padx=10, pady=10)
        fbot = frame(tvw, TOP)
        Button(fbot, text="Ok and Next", command=lambda: master.tb_rollforward_window(tvw)).pack(side=RIGHT, padx=10, pady=10)        
        Button(fbot, text="Cancel", command=tvw.destroy).pack(side=RIGHT, padx=10, pady=10)

    def tb_rollforward_window(master, tvw):
        tvw.destroy()
        trw = Toplevel(master)
        trw.wm_title("Trial Balance Validation: Trial Balance rollforward")
        tData = master.project.getTBData()
        gData = master.project.getGLData()
        gData = gData[['Particulars','Amount']]
        gData = gData.groupby(gData.Particulars).sum()
        jData = tData.merge(gData, on=['Particulars'], how='left')
        #jData['Opening Balance'] = jData['Opening Balance'].map(master.format)
        #jData['Closing Balance'] = jData['Closing Balance'].map(master.format)
        jData.fillna(value=0, inplace=True)
        #jData['Opening Balance'] = jData['Opening Balance'].str.replace(',', '').astype(float)
        #jData['Closing Balance'] = jData['Closing Balance'].str.replace(',', '').astype(float)
        jData['Difference'] = jData.apply(lambda row: row['Opening Balance'] + row['Amount'] - row['Closing Balance'], axis=1)
        jData_dsp = jData[['Particulars', 'Opening Balance', 'Amount', 'Closing Balance', 'Difference']]
        jData_dsp = jData_dsp.rename(columns = {'Amount':'Movement'})
        i = 0
        for col in tuple(jData_dsp):
            if i > 0:
                jData_dsp[col] = jData_dsp[col].map(master.format)
            i=i+1
        ftop = frame(trw, TOP)
        tbt = Table(ftop, dataframe=jData_dsp, width=600, height=21, showtoolbar=True, showstatusbar=True)
        tbt.show()
        fmid = frame(trw, TOP)
        Label(fmid, text="Guidance: Audit team to analyze differences noted in the roll forward report.", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        fbot = frame(trw, TOP)
        Button(fbot, text="Done", command=trw.destroy).pack(side=RIGHT, padx=10, pady=10)        
        Button(fbot, text="Export Report", command=trw.destroy).pack(side=RIGHT, padx=10, pady=10)

    def balance_sheet_window(master):
        bsw = Toplevel(master)
        bsw.geometry('900x450')
        bsw.wm_title("Analyze Balance Sheet")
        #create pivot
        tbData = master.project.getTBData()
        caData = master.project.getCAData()
        jData = tbData.merge(caData, on=['Particulars'])
        jData = jData.loc[(jData["Financial Statement Category"] == "Balance Sheet")]
        pivotData = pd.pivot_table(jData, values=['Closing Balance','Opening Balance'], index=['Account Type', 'Account Category', 'Account Class', 'Account Subclass', 'Particulars'], aggfunc=np.sum).reset_index()
        #fmid: Middle pane
        fmid = frame(bsw, TOP)
        bsTree = ttk.Treeview(fmid)
        bsT_scroll = Scrollbar(fmid, command= bsTree.yview)
        bsTree.configure(yscrollcommand=bsT_scroll.set)
        bsTree["columns"]=("A", "B", "C", "D")
        bsTree.column("A", width=150)
        bsTree.heading("A", text="Closing Balance")
        bsTree.column("B", width=150)
        bsTree.heading("B", text="Opening Balance")
        bsTree.column("C", width=150)
        bsTree.heading("C", text="Variance")
        bsTree.column("D", width=150)
        bsTree.heading("D", text="Variance %")
        accType = pd.pivot_table(pivotData, values=['Closing Balance','Opening Balance'], index=['Account Type'], aggfunc=np.sum).reset_index()
        accCategory = pd.pivot_table(pivotData, values=['Closing Balance','Opening Balance'], index=['Account Type', 'Account Category'], aggfunc=np.sum).reset_index()
        accClass = pd.pivot_table(pivotData, values=['Closing Balance','Opening Balance'], index=['Account Category', 'Account Class'], aggfunc=np.sum).reset_index()
        accSubclass = pd.pivot_table(pivotData, values=['Closing Balance','Opening Balance'], index=['Account Class', 'Account Subclass'], aggfunc=np.sum).reset_index()
        particulars = pd.pivot_table(pivotData, values=['Closing Balance','Opening Balance'], index=['Account Subclass', 'Particulars'], aggfunc=np.sum).reset_index()
        aset_ob = 0
        aset_cb = 0
        liab_ob = 0
        liab_cb = 0
        df = pd.DataFrame(columns=['Account Type', 'Account Category', 'Account Class', 'Account Subclass', 'Closing Balance', 'Opening Balance', 'Variance', 'Variance%'])
        i = 0
        for index, row in accType.iterrows():
            df.loc[i] = [row['Account Type'], None, None, None, row['Closing Balance'], row['Opening Balance'], row['Closing Balance'] - row['Opening Balance'], (row['Closing Balance'] - row['Opening Balance'])*100/row['Opening Balance'] if not row['Opening Balance'] == 0 else np.nan]
            i = i+1
            bsTree.insert('', 'end', 'AccType-'+row['Account Type'], text=row['Account Type'], values=('{:,.0f}'.format(row['Closing Balance']), '{:,.0f}'.format(row['Opening Balance']), '{:,.0f}'.format(row['Closing Balance'] - row['Opening Balance']), '{:,.1f}%'.format((row['Closing Balance'] - row['Opening Balance'])*100/row['Opening Balance']) if not row['Opening Balance'] == 0 else np.nan), open=True)
            for indexCat, rowCat in accCategory.loc[(accCategory['Account Type'] == row['Account Type'])].iterrows():
                df.loc[i] = [None, rowCat['Account Category'], None, None, rowCat['Closing Balance'], rowCat['Opening Balance'], rowCat['Closing Balance'] - rowCat['Opening Balance'], (rowCat['Closing Balance'] - rowCat['Opening Balance'])*100/rowCat['Opening Balance'] if not rowCat['Opening Balance'] == 0 else np.nan]
                i = i+1
                bsTree.insert('AccType-'+row['Account Type'], 'end', 'AccCategory-'+rowCat['Account Category'], text=rowCat['Account Category'], values=('{:,.0f}'.format(rowCat['Closing Balance']), '{:,.0f}'.format(rowCat['Opening Balance']), '{:,.0f}'.format(rowCat['Closing Balance'] - rowCat['Opening Balance']), '{:,.1f}%'.format((rowCat['Closing Balance'] - rowCat['Opening Balance'])*100/rowCat['Opening Balance']) if not rowCat['Opening Balance'] == 0 else np.nan), open=True)
                for indexClass, rowClass in accClass.loc[(accClass['Account Category'] == rowCat['Account Category'])].iterrows():
                    df.loc[i] = [None, None, rowClass['Account Class'], None, rowClass['Closing Balance'], rowClass['Opening Balance'], rowClass['Closing Balance'] - rowClass['Opening Balance'], (rowClass['Closing Balance'] - rowClass['Opening Balance'])*100/rowClass['Opening Balance'] if not rowClass['Opening Balance'] == 0 else np.nan]
                    i = i+1
                    bsTree.insert('AccCategory-'+rowCat['Account Category'], 'end', 'AccClass-'+rowClass['Account Class'], text=rowClass['Account Class'], values=('{:,.0f}'.format(rowClass['Closing Balance']), '{:,.0f}'.format(rowClass['Opening Balance']), '{:,.0f}'.format(rowClass['Closing Balance'] - rowClass['Opening Balance']), '{:,.1f}%'.format((rowClass['Closing Balance'] - rowClass['Opening Balance'])*100/rowClass['Opening Balance']) if not rowClass['Opening Balance'] == 0 else np.nan), open=True)
                    for indexSubClass, rowSubClass in accSubclass.loc[(accSubclass['Account Class'] == rowClass['Account Class'])].iterrows():
                        df.loc[i] = [None, None, None, rowSubClass['Account Subclass'], rowSubClass['Closing Balance'], rowSubClass['Opening Balance'], rowSubClass['Closing Balance'] - rowSubClass['Opening Balance'], (rowSubClass['Closing Balance'] - rowSubClass['Opening Balance'])*100/rowSubClass['Opening Balance'] if not rowSubClass['Opening Balance'] == 0 else np.nan]
                        i = i+1
                        bsTree.insert('AccClass-'+rowClass['Account Class'], 'end', 'AccSubclass-'+rowSubClass['Account Subclass'], text=rowSubClass['Account Subclass'], values=('{:,.0f}'.format(rowSubClass['Closing Balance']), '{:,.0f}'.format(rowSubClass['Opening Balance']), '{:,.0f}'.format(rowSubClass['Closing Balance'] - rowSubClass['Opening Balance']), '{:,.1f}%'.format((rowSubClass['Closing Balance'] - rowSubClass['Opening Balance'])*100/rowSubClass['Opening Balance']) if not rowSubClass['Opening Balance'] == 0 else np.nan))
                        for indexPart, rowPart in particulars.loc[(particulars['Account Subclass'] == rowSubClass['Account Subclass'])].iterrows():
                            bsTree.insert('AccSubclass-'+rowSubClass['Account Subclass'], 'end', 'Particulars-'+rowPart['Particulars'], text=rowPart['Particulars'], values=('{:,.0f}'.format(rowPart['Closing Balance']), '{:,.0f}'.format(rowPart['Opening Balance']), '{:,.0f}'.format(rowPart['Closing Balance'] - rowPart['Opening Balance']), '{:,.1f}%'.format((rowPart['Closing Balance'] - rowPart['Opening Balance'])*100/rowPart['Opening Balance']) if not rowPart['Opening Balance'] == 0 else np.nan))
            if row['Account Type'] == 'Assets':
                aset_ob = row['Opening Balance']
                aset_cb = row['Closing Balance']
                df.loc[i] = ["Total of Assets", None, None, None, row['Closing Balance'], row['Opening Balance'], row['Closing Balance'] - row['Opening Balance'], (row['Closing Balance'] - row['Opening Balance'])*100/row['Opening Balance'] if not row['Opening Balance'] == 0 else np.nan]
                i = i+1
                bsTree.insert('', 'end', 'AccType- Total Assets', text="Total of Assets", values=('{:,.0f}'.format(row['Closing Balance']), '{:,.0f}'.format(row['Opening Balance']), '{:,.0f}'.format(row['Closing Balance'] - row['Opening Balance']), '{:,.1f}%'.format((row['Closing Balance'] - row['Opening Balance'])*100/row['Opening Balance']) if not row['Opening Balance'] == 0 else np.nan))
            else:
                liab_ob = liab_ob + row['Opening Balance']
                liab_cb = liab_cb + row['Closing Balance']
        df.loc[i] = ["Total of Liabilities", None, None, None, liab_cb, liab_ob, liab_cb - liab_ob, (liab_cb - liab_ob)*100/liab_ob if not liab_ob == 0 else np.nan]
        i = i + 1
        bsTree.insert('', 'end', 'AccType- Total Liabilities', text="Total of Liabilities", values=('{:,.0f}'.format(liab_cb), '{:,.0f}'.format(liab_ob), '{:,.0f}'.format(liab_cb - liab_ob), '{:,.1f}%'.format((liab_cb - liab_ob)*100/liab_ob) if not liab_ob == 0 else np.nan))
        df.loc[i] = ["Balance Sheet check", None, None, None, aset_cb+liab_cb, aset_ob+liab_ob, None, None]
        i = i + 1
        bsTree.insert('', 'end', 'BS check', text="Balance Sheet check", values=('{:,.0f}'.format(aset_cb+liab_cb), '{:,.0f}'.format(aset_ob+liab_ob), 'NA', 'NA'))  
        bsTree.pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
        bsT_scroll.pack(side=RIGHT, fill=Y)
        fmid.pack(expand=YES, fill=BOTH)
        fbot = frame(bsw, TOP)
        Label(fbot, text="Guidance: Select GL Account to run Activity Analysis.", relief=FLAT, bg="yellow").pack(side=TOP, padx=10, pady=10)
        def activity_analysis(master):
            selection = bsTree.selection()
            if selection == (): #no selection
                master.status.set("Select a GL Account for activity analysis")
                return
            elif len(selection) > 1:
                master.status.set("Select only 1 GL Account for activity analysis")
                return
            for item in selection:
                if not item[:11] == 'Particulars':
                    master.status.set("Select a GL Account for activity analysis")
                    return
                else:
                    master.activity_analysis_window(bsw, item[12:])
        Button(fbot, text="Activity Analysis", bg="white", fg="RoyalBlue4", command=lambda: activity_analysis(master)).pack(side=TOP, padx=10, pady=10)
        def table_view(master):
            tvw = Toplevel(bsw)
            tvw.wm_title("Analyze Balance Sheet: Table View")
            f0 = frame(tvw, TOP)
            for col in tuple(df):
                if col in ('Closing Balance', 'Opening Balance', 'Variance'):
                    df[col] = df[col].map(master.format)
                elif col in ('Variance%'):
                    df[col] = df[col].map(master.format_percent)
            pivott = Table(f0, dataframe=df, width=1200, height=400, showtoolbar=False, showstatusbar=False)
            pivott.show()
            f1 = frame(tvw, TOP)
            Label(f1, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
            def export_to_excel():
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                df.to_excel(writer)
                writer.save()            
            Button(f1, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=5)        
            Button(f1, text="Done", command=tvw.destroy).pack(side=LEFT, padx=5, pady=5)
            Label(f1, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
        Button(fbot, text="Table View", command=lambda: table_view(master)).pack(side=TOP, padx=10, pady=10)
        Button(fbot, text="Done", command=bsw.destroy).pack(side=TOP, padx=10, pady=10)

    def activity_analysis_window(master, w, p):
        aaw = Toplevel(w)
        #aaw.geometry('800x350')
        aaw.wm_title("Activity Analysis - "+p)
        #get Data
        tbData = master.project.getTBData()
        glData = master.project.getGLData()
        tData = tbData.loc[(tbData["Particulars"] == p)]
        gData = glData.loc[(glData["Particulars"] == p)]
        #Display info
        ftop = frame(aaw, TOP)
        fg1 = frame(ftop, LEFT)
        Label(fg1, text="Opening Balance", relief=FLAT, bg="white", anchor="e").pack(side=TOP, fill=BOTH, expand=YES)
        Label(fg1, text="GL Movement     ", relief=FLAT, bg="white", anchor="e").pack(side=TOP, fill=BOTH, expand=YES)
        Label(fg1, text="Closing Balance  ", relief=FLAT, bg="white", anchor="e").pack(side=TOP, fill=BOTH, expand=YES)
        fg2 = frame(ftop, LEFT)
        Label(fg2, text='{:,.0f}'.format(tData['Opening Balance'].sum()), relief=FLAT, bg="white", anchor="w").pack(side=TOP, fill=BOTH, expand=YES)
        Label(fg2, text='{:,.0f}'.format(gData['Amount'].sum()), relief=FLAT, bg="white", anchor="w").pack(side=TOP, fill=BOTH, expand=YES)
        Label(fg2, text='{:,.0f}'.format(tData['Closing Balance'].sum()), relief=FLAT, bg="white", anchor="w").pack(side=TOP, fill=BOTH, expand=YES)
        fmid = frame(aaw, TOP)
        Data0 = gData[['Date', 'Amount']]
        if Data0.empty: #if no GL movement
            Label(fmid, text="No GL Movement during the period.", relief=FLAT, bg="white").pack(side=TOP, fill=BOTH, expand=YES)
        else:
            Data0 = Data0.groupby(Data0.Date.dt.to_period('M')).sum()
            graphF = frame(fmid, TOP)
            figure = plt.Figure(figsize=(5,3), dpi=100)
            line = FigureCanvasTkAgg(figure, graphF)
            line.get_tk_widget().pack(side=TOP, fill=BOTH)
            ax1 = figure.add_subplot(111)
            Data0.plot.line(legend=True, ax=ax1)
            os.chdir('images')
            figure.savefig('myplot.png')
            os.chdir('..')
            tableF = frame(fmid, TOP)
            Data0 = Data0.reset_index()
            i=0
            for col in tuple(Data0):
                if not i == 0:
                    Data0[col] = Data0[col].map(master.format)
                i = i+1
            pt = Table(tableF, dataframe=Data0, width=600, height=100, showtoolbar=False, showstatusbar=False)
            pt.show()
            tableF1 = frame(fmid, TOP)
            Data1 = gData.groupby(gData.Source).sum()
            Data1 = Data1.reset_index()
            Data1 = Data1[['Source', 'Amount']]
            Data1 = Data1.rename(columns = {'Amount':'Net Activity'})
            Data2 = gData.loc[(gData['Amount'] >= 0)]
            Data2 = Data2.groupby(Data2.Source).sum()
            Data2 = Data2.reset_index()
            Data2 = Data2[['Source', 'Amount']]
            Data2 = Data2.rename(columns = {'Amount':'Dr Activity'})
            Data3 = gData.loc[(gData['Amount'] < 0)]
            Data3 = Data3.groupby(Data3.Source).sum()
            Data3 = Data3.reset_index()
            Data3 = Data3[['Source', 'Amount']]
            Data3 = Data3.rename(columns = {'Amount':'Cr Activity'})
            Data1 = pd.merge(Data1, Data2, how="outer")
            Data1 = pd.merge(Data1, Data3, how="outer")
            i=0
            for col in tuple(Data1):
                if not i == 0:
                    Data1[col] = Data1[col].map(master.format)
                i = i+1
            pt1 = Table(tableF1, dataframe=Data1, width=600, height=100, showtoolbar=False, showstatusbar=False)
            pt1.show()
        fbot = frame(aaw, TOP)
        Button(fbot, text="Done", command=aaw.destroy).pack(side=TOP, padx=10, pady=10)

    def income_statement_window(master):
        bsw = Toplevel(master)
        bsw.geometry('800x450')
        bsw.wm_title("Analyze Income Statement")
        #create pivot
        tbData = master.project.getTBData()
        caData = master.project.getCAData()
        jData = tbData.merge(caData, on=['Particulars'])
        jData = jData.loc[(jData["Financial Statement Category"] == "P&L")]
        pivotData = pd.pivot_table(jData, values='Closing Balance', index=['Account Type', 'Account Category', 'Account Class', 'Account Subclass', 'Particulars'], aggfunc=np.sum).reset_index()
        #fmid: Middle pane
        fmid = frame(bsw, TOP)
        bsTree = ttk.Treeview(fmid)
        bsT_scroll = Scrollbar(fmid, command= bsTree.yview)
        bsTree.configure(yscrollcommand=bsT_scroll.set)
        bsTree["columns"]=("A")
        bsTree.column("A", width=200)
        bsTree.heading("A", text="Amount")
        accType = pd.pivot_table(pivotData, values='Closing Balance', index=['Account Type'], aggfunc=np.sum).reset_index()
        accCategory = pd.pivot_table(pivotData, values='Closing Balance', index=['Account Type', 'Account Category'], aggfunc=np.sum).reset_index()
        accClass = pd.pivot_table(pivotData, values='Closing Balance', index=['Account Category', 'Account Class'], aggfunc=np.sum).reset_index()
        accSubclass = pd.pivot_table(pivotData, values='Closing Balance', index=['Account Class', 'Account Subclass'], aggfunc=np.sum).reset_index()
        particulars = pd.pivot_table(pivotData, values='Closing Balance', index=['Account Subclass', 'Particulars'], aggfunc=np.sum).reset_index()
        exp = 0
        inc = 0
        df = pd.DataFrame(columns=['Account Type', 'Account Category', 'Account Class', 'Account Subclass', 'Amount'])
        i = 0
        for index, row in accType.iterrows():
            df.loc[i] = [row['Account Type'], None, None, None, row['Closing Balance']]
            i = i+1
            bsTree.insert('', 'end', 'AccType-'+row['Account Type'], text=row['Account Type'], values=('{:,.0f}'.format(row['Closing Balance'])), open=True)
            for indexCat, rowCat in accCategory.loc[(accCategory['Account Type'] == row['Account Type'])].iterrows():
                df.loc[i] = [None, rowCat['Account Category'], None, None, rowCat['Closing Balance']]
                i = i+1
                bsTree.insert('AccType-'+row['Account Type'], 'end', 'AccCategory-'+rowCat['Account Category'], text=rowCat['Account Category'], values=('{:,.0f}'.format(rowCat['Closing Balance'])), open=True)
                for indexClass, rowClass in accClass.loc[(accClass['Account Category'] == rowCat['Account Category'])].iterrows():
                    df.loc[i] = [None, None, rowClass['Account Class'], None, rowClass['Closing Balance']]
                    i = i+1
                    bsTree.insert('AccCategory-'+rowCat['Account Category'], 'end', 'AccClass-'+rowClass['Account Class'], text=rowClass['Account Class'], values=('{:,.0f}'.format(rowClass['Closing Balance'])), open=True)
                    for indexSubClass, rowSubClass in accSubclass.loc[(accSubclass['Account Class'] == rowClass['Account Class'])].iterrows():
                        df.loc[i] = [None, None, None, rowSubClass['Account Subclass'], rowSubClass['Closing Balance']]
                        i = i+1
                        bsTree.insert('AccClass-'+rowClass['Account Class'], 'end', 'AccSubclass-'+rowSubClass['Account Subclass'], text=rowSubClass['Account Subclass'], values=('{:,.0f}'.format(rowSubClass['Closing Balance'])))
                        for indexPart, rowPart in particulars.loc[(particulars['Account Subclass'] == rowSubClass['Account Subclass'])].iterrows():
                            bsTree.insert('AccSubclass-'+rowSubClass['Account Subclass'], 'end', 'Particulars-'+rowPart['Particulars'], text=rowPart['Particulars'], values=('{:,.0f}'.format(rowPart['Closing Balance'])))
            if row['Account Type'] == 'Expenses':
                exp = row['Closing Balance']
                df.loc[i] = ["Total of Expenses", None, None, None, row['Closing Balance']]
                i = i+1
                bsTree.insert('', 'end', 'AccType- Total Expenses', text='Total of Expenses', values=('{:,.0f}'.format(row['Closing Balance'])))
            elif row['Account Type'] == 'Revenue':
                inc = row['Closing Balance']
                df.loc[i] = ["Total of Income", None, None, None, row['Closing Balance']]
                i = i+1
                bsTree.insert('', 'end', 'AccType- Total Income', text='Total of Income', values=('{:,.0f}'.format(row['Closing Balance'])))
        df.loc[i] = ["(Profit) or Loss", None, None, None, (inc + exp)]
        bsTree.insert('', 'end', 'AccType- Profit', text='(Profit) or Loss', values=('{:,.0f}'.format(inc + exp)))
        bsTree.pack(side=LEFT, fill=BOTH, expand=YES, pady=10)
        bsT_scroll.pack(side=RIGHT, fill=Y)
        fmid.pack(expand=YES, fill=BOTH)
        fbot = frame(bsw, TOP)
        Label(fbot, text="Guidance: Select GL Account to run Activity Analysis.", relief=FLAT, bg="yellow").pack(side=TOP, padx=10, pady=10)
        def activity_analysis(master):
            selection = bsTree.selection()
            if selection == (): #no selection
                master.status.set("Select a GL Account for activity analysis")
                return
            elif len(selection) > 1:
                master.status.set("Select only 1 GL Account for activity analysis")
                return
            for item in selection:
                if not item[:11] == 'Particulars':
                    master.status.set("Select a GL Account for activity analysis")
                    return
                else:
                    master.activity_analysis_window(bsw, item[12:])
        Button(fbot, text="Activity Analysis", bg="white", fg="RoyalBlue4", command=lambda: activity_analysis(master)).pack(side=TOP, padx=10, pady=10)
        def table_view():
            tvw = Toplevel(bsw)
            tvw.wm_title("Analyze Income Statement: Table View")
            f0 = frame(tvw, TOP)
            for col in tuple(df):
                if col in ('Amount'):
                    df[col] = df[col].map(master.format)
            pivott = Table(f0, dataframe=df, width=1050, height=400, showtoolbar=False, showstatusbar=False)
            pivott.show()
            f1 = frame(tvw, TOP)
            Label(f1, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
            def export_to_excel():
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                df.to_excel(writer)
                writer.save()            
            Button(f1, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=5)        
            Button(f1, text="Done", command=tvw.destroy).pack(side=LEFT, padx=5, pady=5)
            Label(f1, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
        Button(fbot, text="Table View", command=table_view).pack(side=TOP, padx=10, pady=10)
        Button(fbot, text="Done", command=bsw.destroy).pack(side=TOP, padx=10, pady=10)
        fbot.pack(expand=YES, fill=BOTH)

    def format(self, x):
        if str(x) in ('NaN','nan', ''):
            return '0'
        elif type(x) == str:
            return x
        elif x is None:
            return '0'
        else:
            return '{:,.0f}'.format(x)

    def format_percent(self, x):
        if str(x) in ('NaN','nan', ''):
            return '0%'
        elif type(x) == str:
            return x
        elif x is None:
            return '0%'
        else:
            return '{:,.1f}%'.format(x)

    def preparer_map_window(master):
        prmw = Toplevel(master)
        prmw.wm_title("Preparer Map Analysis")
        glData = master.project.getGLData()
        caData = master.project.getCAData()
        prepData = master.project.getPreparerInput()
        jData = glData.merge(caData, on=['Particulars'])
        jData = jData.merge(prepData, on=['Preparer'])
        jData["User"] = jData["Preparer"] + " - " + jData["Title"] + " - " + jData["Department"]
        pivotData = pd.pivot_table(jData, values='Amount', index=['Account Category','Particulars'], columns='User', aggfunc=np.sum).reset_index()
        i = 0
        for col in tuple(pivotData):
            if i > 1:
                pivotData[col] = pivotData[col].map(master.format)
            i=i+1
        #f1: Top pane
        f1 = frame(prmw, TOP)
        pivott = Table(f1, dataframe=pivotData, width=1000, height=21, showtoolbar=True, showstatusbar=True)
        pivott.show()
        f1.pack(expand=YES, fill=BOTH)
        fmid = frame(prmw, TOP)
        def showDetails(master):
            col = pivott.getSelectedColumn()
            row = pivott.getSelectedRow()
            tempD = pivott.model.df
            if tempD.columns[col] in ('Account Category', 'Particulars') :
                return
            if str(tempD.iloc[row, col]) in ('NaN', 'nan', ''):
                return
            username = (tempD.columns[col])[:(tempD.columns[col]).find(' - ')]
            if tempD.columns[0] != 'Account Category' or tempD.columns[1] != 'Particulars':
                ind = list(tempD.index)
                if tempD.columns[0] == 'Account Category':
                    acc_cat = tempD.iloc[row, 0]
                    particulars = ind[row]
                else:
                    acc_cat = ind[row]
                    particulars = tempD.iloc[row, 0]
            else:
                acc_cat = tempD.iloc[row, 0]
                particulars = tempD.iloc[row, 1]
            sdw = Toplevel(prmw)
            sdw.wm_title("Preparer Map Analysis: Details")
            detailsData = glData.loc[(glData['Particulars'] == particulars) & (glData['Preparer'] == username)]
            detailsData['Amount'] = detailsData['Amount'].map(master.format)
            #fd1: Top pane
            fd1 = frame(sdw, TOP)
            detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=True, showstatusbar=True)
            detailst.show()
            fd1.pack(expand=YES, fill=BOTH)
            fd2 = frame(sdw, TOP)
            def showJVDetails(master):
                coli = detailst.getSelectedColumn()
                rowi = detailst.getSelectedRow()
                tD = detailst.model.df
                if not tD.columns[coli] == 'JV Number':
                    return
                if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                    return
                sjdw = Toplevel(sdw)
                sjdw.wm_title("Preparer Map Analysis: JV Number Details")
                jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                fj1 = frame(sjdw, TOP)
                pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                pt.show()
                fj1.pack(expand=YES, fill=BOTH)            
                fj2 = frame(sjdw, TOP)
                def tag_jv(master, jvno):
                    tjw = Toplevel(sjdw)
                    ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                    def ok(master, jvno):
                        if ipt_tag.get() == '':
                            master.status.set("Input Tag comment is mandatory!")
                            return
                        master.project.addTag(jvno, "Process Map: "+ipt_tag.get())
                        tjw.destroy()
                    Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                    Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                    ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                    return
                Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                fj2.pack(expand=YES, fill=BOTH)            
            Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
            fd2.pack(expand=YES, fill=BOTH)
            fd3 = frame(sdw, TOP)
            Button(fd3, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=10)
            fd3.pack(expand=YES, fill=BOTH)
        Button(fmid, text="Details", command=lambda: showDetails(master)).pack(side=TOP)
        fmid.pack(expand=YES, fill=BOTH)
        fbot = frame(prmw, TOP)
        Button(fbot, text="Done", command=prmw.destroy).pack(side=TOP, padx=10, pady=10)
        fbot.pack(expand=YES, fill=BOTH)

    def process_map_window(master):
        pmw = Toplevel(master)
        pmw.wm_title("Process Map Analysis")
        #create pivot
        glData = master.project.getGLData()
        caData = master.project.getCAData()
        jData = glData.merge(caData, on=['Particulars'])
        pivotData = pd.pivot_table(jData, values='Amount', index=['Account Category','Particulars'], columns='Source', aggfunc=np.sum).reset_index()
        i = 0
        for col in tuple(pivotData):
            if i > 1:
                pivotData[col] = pivotData[col].map(master.format)
            i=i+1
        #f1: Top pane
        f1 = frame(pmw, TOP)
        pivott = Table(f1, dataframe=pivotData, width=1000, height=21, showtoolbar=True, showstatusbar=True)
        pivott.show()
        f1.pack(expand=YES, fill=BOTH)
        fmid = frame(pmw, TOP)
        def showDetails(master):
            col = pivott.getSelectedColumn()
            row = pivott.getSelectedRow()
            tempD = pivott.model.df
            if tempD.columns[col] in ('Account Category', 'Particulars') :
                return
            if str(tempD.iloc[row, col]) in ('NaN', 'nan', ''):
                return
            source = tempD.columns[col]
            if tempD.columns[0] != 'Account Category' or tempD.columns[1] != 'Particulars':
                ind = list(tempD.index)
                if tempD.columns[0] == 'Account Category':
                    acc_cat = tempD.iloc[row, 0]
                    particulars = ind[row]
                else:
                    acc_cat = ind[row]
                    particulars = tempD.iloc[row, 0]
            else:
                acc_cat = tempD.iloc[row, 0]
                particulars = tempD.iloc[row, 1]
            sdw = Toplevel(pmw)
            sdw.wm_title("Process Map Analysis: Details")
            detailsData = glData.loc[(glData['Particulars'] == particulars) & (glData['Source'] == source)]
            detailsData['Amount'] = detailsData['Amount'].map(master.format)
            #fd1: Top pane
            fd1 = frame(sdw, TOP)
            detailst = Table(fd1, dataframe=detailsData, width=800, showtoolbar=True, showstatusbar=True)
            detailst.show()
            fd1.pack(expand=YES, fill=BOTH)
            fd2 = frame(sdw, TOP)
            def showJVDetails(master):
                coli = detailst.getSelectedColumn()
                rowi = detailst.getSelectedRow()
                tD = detailst.model.df
                if not tD.columns[coli] == 'JV Number':
                    return
                if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                    return
                sjdw = Toplevel(sdw)
                sjdw.wm_title("Process Map Analysis: JV Number Details")
                jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                fj1 = frame(sjdw, TOP)
                pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                pt.show()
                fj1.pack(expand=YES, fill=BOTH)            
                fj2 = frame(sjdw, TOP)
                def tag_jv(master, jvno):
                    tjw = Toplevel(sjdw)
                    ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                    def ok(master, jvno):
                        if ipt_tag.get() == '':
                            master.status.set("Input Tag comment is mandatory!")
                            return
                        master.project.addTag(jvno, "Process Map: "+ipt_tag.get())
                        tjw.destroy()
                    Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                    Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                    ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                    return
                Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
                fj2.pack(expand=YES, fill=BOTH)            
            Button(fd2, text="Details", command=lambda: showJVDetails(master)).pack(side=TOP, padx=10, pady=10)
            fd2.pack(expand=YES, fill=BOTH)
            fd3 = frame(sdw, TOP)
            Button(fd3, text="Done", command=sdw.destroy).pack(side=TOP, padx=10, pady=10)
            fd3.pack(expand=YES, fill=BOTH)
        Button(fmid, text="Details", command=lambda: showDetails(master)).pack(side=TOP, padx=10, pady=10)
        fmid.pack(expand=YES, fill=BOTH)
        fbot = frame(pmw, TOP)
        Button(fbot, text="Done", command=pmw.destroy).pack(side=TOP, padx=10, pady=10)
        fbot.pack(expand=YES, fill=BOTH)

    def __init__(self):
        pd.set_option('display.max_columns', None)
        pd.set_option('display.max_rows', None)
        pd.set_option('display.expand_frame_repr', False)
        pd.options.display.float_format = '{:,.0f}'.format
        log_filename = datetime.now().strftime('%d-%m-%Y-%H:%M:%S')+'.log'
        logging.basicConfig(filename=log_filename, level=logging.INFO)
        logging.info('Application started')
        Frame.__init__(self)
        self.project=None
        self.status = StringVar()
        self.status.set("Started")
        self.pack(expand=YES, fill=BOTH)
        self.master.title('Audit-Eye')
        os.chdir("images")
        #self.master.iconbitmap(os.path.abspath("logo.svg"))
        icon = ImageTk.PhotoImage(Image.open('logo.png'))
        self.master.tk.call('wm', 'iconphoto', self.master._w, icon)
        img = ImageTk.PhotoImage(Image.open("base.jpg"))
        os.chdir("..")
        self.master.resizable(0,0)
        mBar = Frame(self, relief=RAISED, borderwidth=2)
        mBar.pack(fill=X)
        fileBtn = self.makeFileMenu(mBar)
        toolsBtn = self.makeToolsMenu(mBar)
        helpBtn = self.makeHelpMenu(mBar)
        mBar.tk_menuBar(fileBtn, toolsBtn, helpBtn)
        self.w = Frame(self, relief=SUNKEN, borderwidth=1)
        self.w.pack(side=TOP, expand=YES, fill=BOTH)
        self.l1 = Label(self.w, image=img, relief=SUNKEN)
        self.l1.pack(side=TOP, fill=BOTH, expand=YES, padx=5)
        self.l1.image = img
        self.f0 = frame(self.w, TOP)
        lbl_status = Entry(self.w, textvariable=self.status, justify=LEFT, relief=RAISED)
        lbl_status.pack(side=BOTTOM, fill=BOTH, expand=YES, padx=5)
        self.login_enabled = False
        self.userid = ""

    def login_window(self):
        login = Toplevel(self)
        login.wm_title("Login Window")
        f1 = frame(login, TOP)
        Label(f1, text="Username: ", relief=FLAT, anchor="e").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
        ipt_username = Entry(f1, relief=SUNKEN, width=20)
        ipt_username.pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
        f2 = frame(login, TOP)
        Label(f2, text="Password: ", relief=FLAT, anchor="e").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
        ipt_password = Entry(f2, relief=SUNKEN, width=20, show='*')
        ipt_password.pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=5)
        def onLogin(master):
            url = "http://www.jporthotics.com/login.php?userid="+ipt_username.get()+"&password="+ipt_password.get()
            url_file = urlopen(url)
            master.login_enabled = False
            master.status.set("Login Failed")
            for line in url_file.readlines():
                if line == "Active":
                    master.login_enabled = True
                    master.userid = ipt_username.get()
                    master.status.set("Login successful!")
                    login.destroy()
                break
        Button(login, text="Login", command=lambda: onLogin(self)).pack(side=TOP, padx=5, pady=5)
        login.grab_set()

    def date_validation_window(self):
        ipw = Toplevel(self)
        ipw.wm_title("'Journal Entry' Dates Validation")
        #read gl and get max and min effective and entry dates
        glData = self.project.getGLData()
        min_entry_dt = glData['Date'].min()
        max_entry_dt = glData['Date'].max()
        min_eff_dt = glData['Effective Date'].min()
        max_eff_dt = glData['Effective Date'].max()
        columns = tuple(glData)
        f0 = frame(ipw, TOP)
        #f1: left pane
        f1 = frame(f0, LEFT)
        Label(f1, text=" ", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_dt = Label(f1, text="Start Date:", relief=FLAT, anchor="e")
        lbl_start_dt.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_dt = Label(f1, text="End Date:", relief=FLAT, anchor="e")
        lbl_end_dt.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Center pane
        f2 = frame(f0, LEFT)
        Label(f2, text="Entry Date", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_entry = Label(f2, text=min_entry_dt, relief=SUNKEN)
        lbl_start_entry.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_entry = Label(f2, text=max_entry_dt, relief=SUNKEN)
        lbl_end_entry.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #f3: Right pane
        f3 = frame(f0, LEFT)
        Label(f3, text="Effective Date", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_start_effective = Label(f3, text=min_eff_dt, relief=SUNKEN)
        lbl_start_effective.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        lbl_end_effective = Label(f3, text=max_eff_dt, relief=SUNKEN)
        lbl_end_effective.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f3.pack(expand=YES, fill=BOTH)
        f0.pack(expand=YES, fill=BOTH)
        #f4: Bottom pane
        f4 = frame(ipw, BOTTOM)
        def onOK(master):
            #save entry and effective dates in project object
            master.project.setEntryEffDates(min_entry_dt, max_entry_dt, min_eff_dt, max_eff_dt)
            master.save_project_file()
            ipw.destroy()
        Button(f4, text="Ok", command=lambda: onOK(self)).pack(side=RIGHT, padx=10, pady=10)        
        f4.pack(expand=YES, fill=BOTH)
        Button(f4, text="Cancel", command=ipw.destroy).pack(side=RIGHT, padx=10, pady=10)        
        #f5: Bottom pane
        f5 = frame(ipw, BOTTOM)
        Label(f5, text="Guidance: Audit team to review the start and end date of the data extracted (i.e. the date of \n   first and last transaction) in line with the audit period under consideration.", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f5.pack(expand=YES, fill=BOTH)

    def JEvalidate_window(master):
        #open window
        ipjw = Toplevel(master)
        ipjw.wm_title("Journal Entries Validation")
        glData = master.project.getGLData()
        #A. highlight high JE line item counts
        glData_subset = glData[["JV Number", 'Amount']]
        countli_byJE = glData_subset.pivot_table(index=["JV Number"], aggfunc='count').reset_index()
        countli_byJE = countli_byJE.rename(columns = {'Amount':'Line Item Count'})
        countli_byJE = countli_byJE.sort_values(by=['Line Item Count'], ascending=False)
        f1 = frame(ipjw, TOP)
        Label(f1, text="Count of line items in each JE in descending order").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        if not countli_byJE.empty:
            f1_1 = frame(f1, TOP)
            t = Table(f1_1, dataframe=countli_byJE, showtoolbar=False, showstatusbar=False)
            t.show()
            t.setWrap()
            f1_2 = frame(f1, TOP)
            Label(f1_2, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
            def showJVDetails(master):
                coli = t.getSelectedColumn()
                rowi = t.getSelectedRow()
                tD = t.model.df
                if not tD.columns[coli] == 'JV Number':
                    return
                if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                    return
                sjdw = Toplevel(ipjw)
                sjdw.wm_title("JE Validation: JV Number Details")
                jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                fj1 = frame(sjdw, TOP)
                pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                pt.show()
                fj2 = frame(sjdw, TOP)
                def tag_jv(master, jvno):
                    tjw = Toplevel(sjdw)
                    ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                    def ok(master, jvno):
                        if ipt_tag.get() == '':
                            master.status.set("Input Tag comment is mandatory!")
                            return
                        master.project.addTag(jvno, "JE Validation: "+ipt_tag.get())
                        tjw.destroy()
                    Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                    Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                    ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
            Button(f1_2, text="Details", command=lambda: showJVDetails(master)).pack(side=LEFT, padx=5)
            def export_to_excel():
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                countli_byJE.to_excel(writer)
                writer.save()            
            Button(f1_2, text="Export to Excel", command=export_to_excel).pack(side=LEFT, padx=5)        
            Label(f1_2, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        #B. Unbalanced JEs
        f2 = frame(ipjw, TOP)
        amount_by_JE = glData_subset.groupby(['JV Number']).sum()
        amount_by_JE = amount_by_JE.replace(0, np.nan)
        unbalancedJE = amount_by_JE.dropna(how='any', axis=1) 
        Label(f2, text="Guidance: Audit team may want to analyze few JE's with high line item count \nto check for batch processing of entries.", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        Label(f2, text="Unbalanced JE's: Displays the sum of amount per 'Unbalanced JE'").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        if unbalancedJE.empty:
            Label(f2, text="No Unbalanced JE's to display").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        else:
            f2_1 = frame(f2, TOP)
            t1 = Table(f2_1, dataframe=unbalancedJE, showtoolbar=False, showstatusbar=False)
            t1.show()
            t1.setWrap()
            f2_2 = frame(f2, TOP)
            Label(f2_2, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
            def showDetails(master):
                coli = t1.getSelectedColumn()
                rowi = t1.getSelectedRow()
                tD = t1.model.df
                if not tD.columns[coli] == 'JV Number':
                    return
                if str(tD.iloc[rowi, coli]) in ('NaN', 'nan', ''):
                    return
                sjdw = Toplevel(ipjw)
                sjdw.wm_title("JE Validation: JV Number Details")
                jvdetailsData = glData.loc[(glData['JV Number'] == tD.iloc[rowi, coli])]
                jvdetailsData['Amount'] = jvdetailsData['Amount'].map(master.format)
                fj1 = frame(sjdw, TOP)
                pt = Table(fj1, dataframe=jvdetailsData, width=700, showtoolbar=True, showstatusbar=True)
                pt.show()
                fj2 = frame(sjdw, TOP)
                def tag_jv(master, jvno):
                    tjw = Toplevel(sjdw)
                    ipt_tag = Entry(tjw, relief=SUNKEN, width=40)
                    def ok(master, jvno):
                        if ipt_tag.get() == '':
                            master.status.set("Input Tag comment is mandatory!")
                            return
                        master.project.addTag(jvno, "JE Validation: "+ipt_tag.get())
                        tjw.destroy()
                    Button(tjw, text="Done", command=lambda:ok(master, jvno)).pack(side=BOTTOM, padx=10, pady=10)
                    Label(tjw, text="Document rationale for JVno.("+str(jvno)+"):", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                    ipt_tag.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
                Button(fj2, text="Tag JV", command=lambda: tag_jv(master, tD.iloc[rowi, coli])).pack(side=TOP, padx=10, pady=10)
                Button(fj2, text="Done", command=sjdw.destroy).pack(side=TOP, padx=10, pady=10)
            Button(f2_2, text="Details", command=lambda: showDetails(master)).pack(side=LEFT, padx=5)
            def export_toexcel():
                savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
                if savefile == '':
                    return
                writer = pd.ExcelWriter(savefile, engine='xlsxwriter')
                unbalancedJE.to_excel(writer)
                writer.save()            
            Button(f2_2, text="Export to Excel", command=export_toexcel).pack(side=LEFT, padx=5)        
            Label(f2_2, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5)
        fmid = frame(ipjw, TOP)
        Label(fmid, text="Guidance: If the amount is not zero for any JE, the audit team needs to \nre-validate the data from the client.", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        fmid.pack(expand=YES, fill=BOTH)
        f3 = frame(ipjw, TOP)
        Label(f3, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=10)
        def onCancel(master, ipjw):
            ipjw.destroy()
        Button(f3, text="Cancel", command=lambda: onCancel(master, ipjw)).pack(side=LEFT, padx=5, pady=10)        
        def onApprove(master, ipjw):
            master.project.saveJEvalidated()
            master.save_project_file()
            ipjw.destroy()
        Button(f3, text="Ok", command=lambda: onApprove(master, ipjw)).pack(side=LEFT, padx=5, pady=10)
        Label(f3, text=" ").pack(side=LEFT, fill=BOTH, expand=YES, padx=5, pady=10)

    def ipt_select_sysvalues_window(self):
        if self.project is None:
            self.status.set("First Create Project or Load existing Project!")
            return
        elif self.project.getGLInputFile() == '' or self.project.getTBInputFile() == '' or self.project.getCAInputFile() == '':
            self.status.set("First Upload Data Files. Select Tools -> Manage Data")
            return        
        ssw = Toplevel(self)
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
            #master.cleanup()
            ssw.destroy()
        Button(f1, text="Cancel", command=lambda: onCancel(self, ssw)).pack(side=BOTTOM, padx=10, pady=10)        
        f1.pack(expand=YES, fill=BOTH)
        #f2: Center pane
        f2 = frame(ssw, LEFT)
        ipt_entries = ttk.Combobox(f2, values=("Only Manual Entries","Only System Entries","Both System and Manual Entries"))
        ipt_entries.set("Only Manual Entries")
        ipt_entries.pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        ipt_sysField = ttk.Combobox(f2, values=columns)
        ipt_sysValues = Listbox(f2,selectmode='multiple', exportselection=False)
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
        Button(f2, text="Ok and Next", command=lambda: onOK(self, ssw)).pack(side=BOTTOM, padx=10, pady=10)        
        f2.pack(expand=YES, fill=BOTH)

    def ipt_acc_def_window(self, master, ssw):
        ssw.destroy()
        iadw = Toplevel(master)
        iadw.wm_title("Validate Input Parameters: Account Definition")
        #read CoA
        caData = master.project.getCAData()
        #f1: Top pane
        f1 = frame(iadw, TOP)
        Label(f1, text="Displays the Chart of Accounts imported:", relief=FLAT).pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        ftop = frame(iadw, TOP)
        accTree = ttk.Treeview(ftop)
        accT_scroll = Scrollbar(ftop, command= accTree.yview)
        accTree.configure(yscrollcommand=accT_scroll.set)
        caData_subset = caData[['Account Category', 'Account Class', 'Account Subclass', 'Particulars']]
        df2 = pd.DataFrame({'Account Category': caData_subset['Account Category'].unique()})
        df2['Account Class'] = [list(set(caData_subset['Account Class'].loc[caData_subset['Account Category'] == x['Account Category']])) for _, x in df2.iterrows()]
        df3 = pd.DataFrame({'Account Class': caData_subset['Account Class'].unique()})
        df3['Account Subclass'] = [list(set(caData_subset['Account Subclass'].loc[caData_subset['Account Class'] == x['Account Class']])) for _, x in df3.iterrows()]
        df4 = pd.DataFrame({'Account Subclass': caData_subset['Account Subclass'].unique()})
        df4['Particulars'] = [list(set(caData_subset['Particulars'].loc[caData_subset['Account Subclass'] == x['Account Subclass']])) for _, x in df4.iterrows()]
        gi = 0
        for item in caData_subset['Account Category'].unique().tolist():
            if str(item) == 'nan':
                continue
            accTree.insert('', 'end', 'AccCategory-'+item, text=item)
            for ite in df2['Account Class'].loc[df2['Account Category'] == item]:
                for x in ite:
                    accTree.insert('AccCategory-'+item, 'end', 'AccClass-'+x, text=x)
                    for it in df3['Account Subclass'].loc[df3['Account Class'] == x]:
                        for y in it:
                            accTree.insert('AccClass-'+x, 'end', 'AccSubclass-'+y, text=y)
                            for i in df4['Particulars'].loc[df4['Account Subclass'] == y]:
                                for z in i:
                                    gl_no = str(z)[:-3]
                                    accTree.insert('AccSubclass-'+y, 'end', 'AccGL-'+str(gi), text=gl_no)
                                    gi += 1
        accTree.pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        accT_scroll.pack(side=RIGHT, fill=Y)
        ftop.pack(expand=YES, fill=BOTH)
        fmid = frame(iadw, TOP)
        Label(fmid, text="Guidance: Audit team to review the mapping here.", relief=FLAT, bg="yellow").pack(side=TOP, fill=BOTH, expand=YES, padx=10, pady=10)
        fmid.pack(expand=YES, fill=BOTH)
        #f2: Middle pane
        f2 = frame(iadw, TOP)
        def uploadNewCOA(master):
            iadw.destroy()
            master.input_data_window()
        Button(f2, text="Manage Data to upload new COA", command=lambda: uploadNewCOA(master)).pack(side=RIGHT, padx=10, pady=10)        
        def export_COA(master):
            caData = master.project.getCAData()
            savefile = asksaveasfilename(filetypes=(("Xlsx files","*.xlsx"),("All files","*")))
            if savefile == '':
                return
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
        if master.project.getSourceInputF() != '':
            text_source.insert(END, master.project.getSourceInput()) #display dataframe in text
        def browseSourceF(master):
            master.sourceFileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"))))
            if master.sourceFileName.get() == (): #in case of cancel or no selection
                master.sourceFileName.set('')
                return
            master.changeInputF += 1
            master.project.setSourceInputF(master.sourceFileName.get())
            text_source.delete(1.0, END)
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
        iupw.wm_title("Validate Input Parameters: Preparer")
        #f1: Top pane
        f1 = frame(iupw, TOP)
        Label(f1, text="Verify that preparer file has following fields: Preparer, Full Name, Title, Department and Role", relief=FLAT).pack(side=LEFT, fill=BOTH, expand=YES, padx=10, pady=10)
        master.preparerFileName = StringVar()
        master.preparerFileName.set('')
        master.changeInputF = 0
        Button(f1, text="Preparer File...", command=lambda: browsePreparerF(master)).pack(side=LEFT, padx=10, pady=10)
        f1.pack(expand=YES, fill=BOTH)
        #f2: Middle pane
        f2 = frame(iupw, TOP)
        text_preparer = Text(f2, height=20, width=100)
        if master.project.getPreparerInputF() != '':
            text_preparer.insert(END, master.project.getPreparerInput()) #display dataframe in text
        def browsePreparerF(master):
            master.preparerFileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"))))
            if master.preparerFileName.get() == (): #in case of cancel or no selection
                master.preparerFileName.set('')
                return
            master.changeInputF += 1
            master.project.setPreparerInputF(master.preparerFileName.get())
            text_preparer.delete(1.0, END)
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
        if master.project.getBUInputF() != '':
            text_BU.insert(END, master.project.getBUInput()) #display dataframe in text
        def browseBUFile(master):
            master.BUFileName.set(askopenfilename(filetypes=(("xlsx", "*.xlsx"),("xls", "*.xls"))))
            if master.BUFileName.get() == (): #in case of cancel or no selection
                master.BUFileName.set('')
                return
            master.changeInputF += 1
            master.project.setBUInputF(master.BUFileName.get())
            text_BU.delete(1.0, END)
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
        if self.project is None:
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
        if not self.login_enabled:
            self.status.set("Login first!")
            return
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
            url_file = urlopen("http://www.jporthotics.com/load_project.php?userid="+master.userid+"&project="+project_name)
            self.winfo_toplevel().title("Audit-Eye: "+project_name)
            master.status.set("Project Created. Now select Tools -> Manage Data")
            parent.destroy()
        Button(f2, text="Submit", command=lambda: onSubmit(cpw, self, ipt_project_name.get(), ipt_fy_end.get_date().strftime('%d/%m/%Y'), ipt_timing.get_date().strftime('%d/%m/%Y'), ipt_creator.get(), ipt_sector.get())).pack(side=TOP, padx=10, pady=10)
        f2.pack(expand=YES, fill=BOTH)
        #grab_set to refrain any activity on main window
        cpw.grab_set()

    def load_project(self):
        if not self.login_enabled:
            self.status.set("Login first!")
            return
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
                self.winfo_toplevel().title("Audit-Eye: "+projectName)
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
        url_file = urlopen("http://www.jporthotics.com/load_project.php?userid="+self.userid+"&project="+projectName)
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
                self.status.set("Loading Project...Done. Now select Tools -> Input Parameters")
        else:
            self.project = self.Project(projectName, fy_end, timing, creator, sector, fname)
            try:
                try:
                    self.project.setTags()
                except IOError:
                    self.status.set("Tags missing")
                finally:
                    cwd = os.getcwd()
                    if cwd[-4:] == 'Data':
                        os.chdir('..')
                self.project.setGLInputFile(glInputFile)
                self.project.setTBInputFile(tbInputFile)
                self.project.setCAInputFile(caInputFile)
                self.status.set("Loading Project...Done")
            except IOError:
                self.status.set("Missing Data Files... Select Tools -> Manage Data; and upload data files again.")
                return
            else:
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
        CmdBtn.menu.add_command(label="Login", command=self.login_window)
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
        CmdBtn.menu.add_command(label='Input Parameters', underline=0, command=self.ipt_select_sysvalues_window)
        CmdBtn['menu'] = CmdBtn.menu
        return CmdBtn

    def makeHelpMenu(self, mBar):
        Help_button = Menubutton(mBar, text='Help', underline=0)
        Help_button.pack(side=LEFT, padx='2m')
        Help_button["state"] = DISABLED
        return Help_button

if __name__ == '__main__':
    Application().mainloop()
