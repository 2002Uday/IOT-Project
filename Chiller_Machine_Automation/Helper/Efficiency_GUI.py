import sys
sys.path.append("../")

import tkinter.scrolledtext as tkscrolled
from tkinter import messagebox
import tkinter as tk
from tkinter import ttk
import random
import os

from Configuration import config
from Helper import Excel_Helper
from Helper import Text_Helper

InputHeadings = {"Chilled Water Flow":"m^3/hr", 
                 "Chiller Power":"Kw",
                 "Ambient Dry Buld Temp. (DBT)":"Deg C",
                 "Ambient Wet Buld Temp. (WBT)":"Deg C",
                 "Relative Humidity (RH)":"%"
                 }

ChilledWaterHeadings = {"Chilled Water Temp. Inlet (Tc in)":"Deg C",
                        "Chilled Water Temp. Outlet (Tc out)":"Deg C",
                        "Refrigerent Temp. in Chiller": "Deg C"
                        }

CoolingWaterHeadings = {"Cooling Water Temp. Inlet (Tcw in)":"Deg C",
                        "Cooling Water Temp. Outlet (Tcw out)":"Deg C",
                        "Refrigerent Temp. in Condensor": "Deg C"}

OutputHeadings = {"Chiller Range":"Deg C",
                 "Condensor Range":"Deg C",
                 "Chiller Operating TR":"TR",
                 "Chiller Power":"Kw"}

TimeCountHeadings = ["Timer (Minutes)",
                     "Counter (Numbers)"]
      

MeasuredParameters = ["Chiller Efficiency (KW/TR)",
                      "COP",
                      "Chiller Approch (Deg C)",
                      "Condensor Approch (Deg C)",
                      ]

class Efficiency:
    def __init__(self):
        self.TimingValue = {}
        self.InputValues = {}
        self.ChilledWaterValues = {}
        self.CoolingWaterValues = {}
        self.OutputValues = {}
        self.MeasuredValues = {}
        self.AllPort = {}
        
        self.excelObj = Excel_Helper.ExcelOperation(config.EXCEL_FILE_TEMPLATE) 
        self.excelObj.openExcel()
        self.excelObj.getSheet("Sheet")
        self.StartRow = 3
         
        self.measurement = 1
        self.count = 0
        self.timeOut = 0
        self.length = False
        
        self.root = tk.Tk()
        self.root.state('zoomed')
        self.root.title ("Chiller Machine Efficiency")
        
        self.LogWindow = tkscrolled.ScrolledText(self.root,height=200,width=80)
        self.LogWindow.pack(side=tk.RIGHT, fill=tk.NONE)
        self.LogWindow.insert(tk.END,"*** This is your Log Window ***\n")
        self.LogWindow.insert(tk.END,"="*80)
        self.LogWindow.insert(tk.END,"\nUsage: Please select all the port number and fill up timing and counting parameter then press START button\n")
        self.LogWindow.insert(tk.END,"="*80)
        self.LogWindow.yview(tk.END)
        
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)
        
        self.ents = self.MakeForm(self.root)

        self.StartButton = tk.Button(self.root,text="Start",bg="green", fg="black", font= ("bold"),
                                     command=(lambda e=self.ents: self.FetchTimingCount(e)))
        self.StartButton.pack(side=tk.LEFT,padx=15)
        
        self.StopButton = tk.Button(self.root,text="Stop",bg="red",fg="black",font=("bold"), command=self.FindBestEffiency)
        self.StopButton.pack(side=tk.LEFT)

        self.root.bind('<Return>', (lambda event, e=self.ents: self.FetchTimingCount(e)))

        self.root.mainloop()

    def RemoveExistingFile(self):
        self.excelObj.RemoveExistingFileIfAvailable()

    def DefaulFrame(self,LableText=None,Width=None,FG=None,BG=None,Font=None,PadY=None):
        frame = tk.Frame(self.root)
        frame.pack(side=tk.TOP,fill=tk.X)
        Label = tk.Label(frame,text=LableText,width=Width,bg=BG,fg=FG,font=Font)
        Label.pack(side=tk.LEFT,pady=PadY)

    def HeadingLayout(self,Heads=None,Width=23):
        frame = tk.Frame(self.root)
        frame.pack(side=tk.TOP,fill=tk.X)
        for head in Heads:
            if head == "Port No.":
                label = tk.Label(frame,width=14,text=head)
            elif head == "Parameters":
                label = tk.Label(frame,width=Width,text=head)
            else:
                label = tk.Label(frame,width=10,text=head)
            label.pack(side=tk.LEFT)

    def MakeForm(self,root):
        self.RemoveExistingFile()

        MeasuredUnit = []
        for parameter in MeasuredParameters:
            frame1 = tk.Frame(root)
            frame1.pack(side=tk.TOP,fill=tk.X)
            ParameterLabel = tk.Label(frame1,text=parameter,width=25,bg="lime",fg="black",font=("bold"))
            ParameterLabel.pack(side=tk.LEFT,padx=10)
            Entry = tk.Entry(frame1,bg="lime",width=15,font=("bold"))
            Entry.pack(side=tk.LEFT,pady=3)
            Entry.insert(0,1)
            Entry["state"] = "disabled"
            
            MeasuredUnit.append((parameter,Entry))
        
        #***********************************************
        # Input Heading Layout
        #***********************************************
        self.DefaulFrame(LableText="INPUTS",Width=50,FG="Black",BG="Gray",Font=("bold"),PadY=3)
        self.HeadingLayout (Heads=["Parameters", "Unit", "Port No.", "Value"])

        #***********************************************
        # Input Parameters Layout
        #***********************************************
        vcmd = (self.root.register(self.validate),
                     '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        InputParameters = []
        for field in InputHeadings:
            frame1 = tk.Frame(root)
            frame1.pack(side=tk.TOP,fill=tk.X)
            
            label1 = tk.Label(frame1,width=23,text=field)
            label1.pack(side=tk.LEFT)
            
            label2 = tk.Label(frame1,width=10,text=InputHeadings[field])
            label2.pack(side=tk.LEFT)
            
            temp = tk.StringVar()
            DropDown = ttk.Combobox(frame1,width=10,textvariable=temp,state="readonly")
            DropDown["value"] = ("COM1","COM2","COM3")
            DropDown.pack(side=tk.LEFT, fill=tk.X, padx=10)
            
            entry1 = tk.Entry(frame1,width=13,validate = 'key',validatecommand = (vcmd))
            entry1.pack(side=tk.LEFT,fill=tk.X)
            entry1["state"] = "disabled"
        
            InputParameters.append((field,DropDown,entry1))
            
        #***********************************************
        # Chilled Water Heading Layout
        #***********************************************
        self.DefaulFrame(LableText="CHILLED WATER PARAMETERS",Width=50,FG="Black",BG="Gray",Font=("bold"),PadY=3)
        self.HeadingLayout (Heads=["Parameters", "Unit", "Port No.", "Value"],Width=27)
        
        #***********************************************
        # Chilled Water Layout
        #***********************************************
        vcmd = (self.root.register(self.validate),
                     '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        
        ChilledWaterParameters = []
        for field in ChilledWaterHeadings:
            frame1 = tk.Frame(root)
            frame1.pack(side=tk.TOP,fill=tk.X)
            
            label1 = tk.Label(frame1,width=27,text=field)
            label1.pack(side=tk.LEFT)
            
            label2 = tk.Label(frame1,width=10,text=ChilledWaterHeadings[field])
            label2.pack(side=tk.LEFT)
            
            temp = tk.StringVar()
            DropDown = ttk.Combobox(frame1,width=10,textvariable=temp,state="readonly")
            DropDown["value"] = ("COM1","COM2","COM3")
            DropDown.pack(side=tk.LEFT, fill=tk.X, padx=10)
            
            entry1 = tk.Entry(frame1,width=13,validate = 'key',validatecommand = (vcmd))
            entry1.pack(side=tk.LEFT,fill=tk.X)
            entry1["state"] = "disabled"

            ChilledWaterParameters.append((field,DropDown,entry1))

        #***********************************************
        # Cooling Water Heading Layout
        #***********************************************
        self.DefaulFrame(LableText="COOLING WATER PARAMETERS",Width=50,FG="Black",BG="Gray",Font=("bold"),PadY=3)
        self.HeadingLayout (Heads=["Parameters", "Unit", "Port No.", "Value"],Width=30)
        
        #***********************************************
        # Cooling Water Layout
        #***********************************************
        vcmd = (self.root.register(self.validate),
                     '%d', '%i', '%P', '%s', '%S', '%v', '%V', '%W')
        
        CoolingWaterParameters = []
        for field in CoolingWaterHeadings:
            frame1 = tk.Frame(root)
            frame1.pack(side=tk.TOP,fill=tk.X)
            
            label1 = tk.Label(frame1,width=30,text=field)
            label1.pack(side=tk.LEFT)
            
            label2 = tk.Label(frame1,width=10,text=CoolingWaterHeadings[field])
            label2.pack(side=tk.LEFT)
            
            temp = tk.StringVar()
            DropDown = ttk.Combobox(frame1,width=10,textvariable=temp,state="readonly")
            DropDown["value"] = ("COM1","COM2","COM3")
            DropDown.pack(side=tk.LEFT, fill=tk.X, padx=10)
            
            entry1 = tk.Entry(frame1,width=13,validate = 'key',validatecommand = (vcmd))
            entry1.pack(side=tk.LEFT,fill=tk.X)
            entry1["state"] = "disabled"
        
            CoolingWaterParameters.append((field,DropDown,entry1))
            
        #***********************************************
        # Output Heading Layout
        #***********************************************
        self.DefaulFrame(LableText="OUTPUTS",Width=50,FG="Black",BG="Gray",Font=("bold"),PadY=3)
        
        Heads = ["Parameters", "Unit", "Value"]
        frame = tk.Frame(root)
        frame.pack(side=tk.TOP,fill=tk.X)
        for head in Heads:
            if head == "Parameters":
                label = tk.Label(frame,width=15,text=head)
            else:
                label = tk.Label(frame,width=10,text=head)
            label.pack(side=tk.LEFT)
            
        #***********************************************
        # Output Parameters Layout
        #***********************************************
        OutputParameters = []
        for field in OutputHeadings:
            frame1 = tk.Frame(root)
            frame1.pack(side=tk.TOP,fill=tk.X)
            
            label1 = tk.Label(frame1,width=15,text=field)
            label1.pack(side=tk.LEFT)
            
            label2 = tk.Label(frame1,width=10,text=OutputHeadings[field])
            label2.pack(side=tk.LEFT)
            
            entry1 = tk.Entry(frame1,width=13)
            entry1.pack(side=tk.LEFT,fill=tk.X)
            entry1.insert(0,1)
            entry1["state"] = "disabled"
            
            OutputParameters.append((field,entry1))

        #***********************************************
        # Timing Heading Layout
        #***********************************************
        self.DefaulFrame(LableText="TIMING AND COUNTING",Width=50,FG="Black",BG="Gray",Font=("bold"),PadY=5)
   
        #***********************************************
        # Timer and Counter Layout
        #***********************************************
        TimingParameters = []
        for field in TimeCountHeadings:
            frame2 = tk.Frame(root)
            frame2.pack(side=tk.TOP,fill=tk.X)
            
            label1 = tk.Label(frame2,width=17,text=field)
            label1.pack(side=tk.LEFT)

            entry1 = tk.Entry(frame2,width=12)
            entry1.pack(side=tk.LEFT,fill=tk.X)
            
            TimingParameters.append((field,entry1))
            
        return [MeasuredUnit,InputParameters,ChilledWaterParameters,CoolingWaterParameters,OutputParameters,TimingParameters]
                        
    #*************************************************************************************************
    #Function Name  : FindBestEffiency()
    #Description    : This function is used to find out the best efficiency and related configurations
    #Parameters     : [ Entry List ]
    #*************************************************************************************************
    def FindBestEffiency(self):
        self.excelObj.saveExcel()
        BestEfficiencyData = self.excelObj.fetchEffiency()
        
        textFileObject = Text_Helper.TextFileOperations(config.TEXT_FILE)
        textFileObject.openFile()
        textFileObject.writeTextFile(BestEfficiencyData)
        textFileObject.closeTextFile()
        print ("Please refer the below path to get your output data")
        path = config.OUTPUT_PATH
        os.chdir(path)
        currentPath = os.getcwd()
        print (currentPath)
        self.root.destroy()
    
    #*************************************************************************************************
    #Function Name  : on_closing()
    #Description    : This function is used to close the window
    #Parameters     : [ Entry List ]
    #*************************************************************************************************
    def on_closing(self):
        if messagebox.askokcancel("Quit", "Do you want to quit?"):
            self.FindBestEffiency()
            sys.exit()
            
    #*************************************************************************************************
    #Function Name  : FetchTimingCount()
    #Description    : This function is used to fetch the timing and counting values
    #Parameters     : [ Entry List ]
    #*************************************************************************************************
    def FetchTimingCount(self,entries):
        self.TimingValue.clear()
        TimeingCounting = entries[-1]

        for entry in TimeingCounting:
            parameter = entry[0]
            value = entry[1].get()
            
            if value != "":
                value = float(value)
                self.TimingValue[parameter] = value
            else:
                pass
        
        if (len(self.TimingValue) == 2):
            self.count = int(self.TimingValue["Counter (Numbers)"])
            self.timeOut = float(self.TimingValue["Timer (Minutes)"])
            self.StartButton["state"] = "disabled"
            self.length = True
            self.CheckCounter()
             
        else:
            self.StartButton["state"] = "normal"
            self.LogWindow.insert(tk.END,"\nWarning: Please fill up all Timing and Counting parameters and then press START button again\n")
            self.LogWindow.yview(tk.END)
            self.length = False

    #*************************************************************************************************
    #Function Name  : CheckCounter()
    #Description    : This function is used to verify the counter value
    #Parameters     : NA
    #*************************************************************************************************
    def CheckCounter(self):
        if self.count != 0:
            self.FetchAllPortDetails()
            
            if self.length:
                self.LogWindow.insert(tk.END, "\n**********************\n")
                self.LogWindow.insert(tk.END, "Measurement - "+str(self.measurement))
                self.LogWindow.insert(tk.END, "\n**********************\n")
                self.LogWindow.yview(tk.END)
                self.measurement = self.measurement + 1
                self.count = self.count-1
                tempTimeOut = int(self.timeOut * 60 * 1000)
                self.root.after(tempTimeOut, self.CheckCounter)
            else:
                self.LogWindow.insert(tk.END,"\nWarning: Please select all the ports and then press START button again\n")
                self.LogWindow.yview(tk.END)
                #self.LogWindow.insert(tk.END,"\nWarning: Please fill up all the fields\n")
                
        else:
            self.measurement = 1
            self.LogWindow.insert(tk.END,"="*80)
            self.LogWindow.insert(tk.END,"\nYour execution is completed. Now press")
            self.LogWindow.insert(tk.END,"\nSTART: To start execution again")
            self.LogWindow.insert(tk.END,"\nSTOP: To close the GUI and fetch best efficiency data\n")
            self.LogWindow.insert(tk.END,"="*80)
            self.LogWindow.yview(tk.END)
            self.StartButton["state"] = "normal"
            
    #*************************************************************************************************
    #Function Name  : FetchAllPortDetails()
    #Description    : This function is used to fetch the input parameters values
    #Parameters     : [ Entry List, Submit Button, Log Window ]
    #*************************************************************************************************
    def FetchAllPortDetails(self):
        self.AllPort.clear()
        
        for ent in self.ents[1:4]:
            for entry in ent:
                parameter = entry[0]
                port = entry[1].get()
                
                if port != "":
                    self.AllPort[parameter] = port
                else:
                    pass
                     
        if (len(self.AllPort) == 11):
            self.StartButton["state"] = "disabled"
            self.length = True
            self.FetchInputValue()
        else:
            self.StartButton["state"] = "normal"
            self.length = False            
                
    #*************************************************************************************************
    #Function Name  : FetchInputValue()
    #Description    : This function is used to fetch the input parameters values
    #Parameters     : [ Entry List, Submit Button, Log Window ]
    #*************************************************************************************************
    def FetchInputValue(self):
        self.InputValues.clear()
        InputEntries = self.ents[1]

        startRange = 28000
        endRange = 30000
        for entry in InputEntries:
            parameter = entry[0]
            entry[2]["state"] = "normal"
            entry[2].delete(0,tk.END)
            value = random.randrange(startRange,endRange)
            startRange = startRange-20
            endRange = endRange-20
            entry[2].insert(0,value)
            port = entry[1].get()
            value = entry[2].get()
            
            if port != "" and value != "":
                value = float(value)
                tempValue = (port,value)
                self.InputValues[parameter] = tempValue
            else:
                pass
            entry[2]["state"] = "disabled"
            
        self.FetchChilledWater()
  
    #*************************************************************************************************
    #Function Name  : FetchChilledWater()
    #Description    : This function is used to fetch the chilled water parameters value
    #Parameters     : [ Entry List, Submit Button, Log Window ]
    #*************************************************************************************************
    def FetchChilledWater(self):
        self.ChilledWaterValues.clear()
        ChilledWaterEntries = self.ents[2]
        
        startRange = 18000
        endRange = 20000
        for entry in ChilledWaterEntries:
            parameter = entry[0]
            entry[2]["state"] = "normal"
            entry[2].delete(0,tk.END)
            value = random.randrange(startRange,endRange)
            startRange = startRange-20
            endRange = endRange-20
            entry[2].insert(0,value)
            port = entry[1].get()
            value = entry[2].get()
            
            if port != "" and value != "":
                value = float(value)
                tempValue = (port,value)
                self.ChilledWaterValues[parameter] = tempValue
            else:
                pass
            entry[2]["state"] = "disabled"
            
        self.FetchCoolingWater()

    #*************************************************************************************************
    #Function Name  : FetchCoolingWater()
    #Description    : This function is used to fetch the cooling water parameters value
    #Parameters     : [ Entry List, Submit Button, Log Window ]
    #*************************************************************************************************
    def FetchCoolingWater(self):
        self.CoolingWaterValues.clear()
        CoolingWaterEntries = self.ents[3]
        
        startRange = 8000
        endRange = 10000
        for entry in CoolingWaterEntries:
            parameter = entry[0]
            entry[2]["state"] = "normal"
            entry[2].delete(0,tk.END)
            value = random.randrange(startRange,endRange)
            startRange = startRange-20
            endRange = endRange-20
            entry[2].insert(0,value)
            port = entry[1].get()
            value = entry[2].get()
            
            if port != "" and value != "":
                value = float(value)
                tempValue = (port,value)
                self.CoolingWaterValues[parameter] = tempValue
            else:
                pass
            entry[2]["state"] = "disabled"
            
        self.FetchOutputs()

    #*************************************************************************************************
    #Function Name  : FetchOutputs()
    #Description    : This function is used to fetch the output parameters value
    #Parameters     : [ Entry List, Submit Button, Log Window ]
    #*************************************************************************************************
    def FetchOutputs(self):
        self.OutputValues.clear()
        OutputEntries = self.ents[4]
    
        for entry in OutputEntries:
            parameter = entry[0]
            ent = entry[1]
            ent["state"] = "normal"
            ent.delete(0,tk.END)
            
            SetValue = None
            if parameter == "Chiller Range":
                SetValue = ((self.ChilledWaterValues["Chilled Water Temp. Inlet (Tc in)"][1]) -
                             (self.ChilledWaterValues["Chilled Water Temp. Outlet (Tc out)"][1])) 
            
            elif parameter == "Condensor Range":
                SetValue = ((self.CoolingWaterValues["Cooling Water Temp. Outlet (Tcw out)"][1]) -
                             (self.CoolingWaterValues["Cooling Water Temp. Inlet (Tcw in)"][1])) 
                
            elif parameter == "Chiller Operating TR":
                tempValue1 = (self.InputValues["Chilled Water Flow"][1])*4.18*1000
                tempValue2 = ((self.ChilledWaterValues["Chilled Water Temp. Inlet (Tc in)"][1]) -
                              (self.ChilledWaterValues["Chilled Water Temp. Outlet (Tc out)"][1]))
                tempValue3 = 3.51*3600
                SetValue = (tempValue1*tempValue2)/tempValue3
                
            elif parameter == "Chiller Power":
                SetValue = (self.InputValues["Chiller Power"][1])
                
                
            ent.insert(0,SetValue)
   
            getValue = (entry[1].get())
            self.OutputValues[parameter] = getValue
            ent["state"] = "disabled"

        self.FetchMeasuredParameters()
        
    #*************************************************************************************************
    #Function Name  : FetchMeasuredParameters()
    #Description    : This function is used to fetch the measured parameters value
    #Parameters     : [ Entry List, Submit Button, Log Window ]
    #*************************************************************************************************
    def FetchMeasuredParameters(self):
        self.MeasuredValues.clear()
        MeasuredEntries = self.ents[0]

        for entry in MeasuredEntries:
            parameter = entry[0]
            ent = entry[1]
            ent["state"] = "normal"
            ent.delete(0,tk.END)
            
            SetValue = None
            if parameter == "Chiller Efficiency (KW/TR)":
                try:
                    SetValue = (float(self.OutputValues["Chiller Power"])/float(self.OutputValues["Chiller Operating TR"])) 
                except ZeroDivisionError:
                    print ("Error: Received Zero division error and closed the window")
                    sys.exit()
                    
            elif parameter == "COP":
                SetValue = 3.517/float(self.MeasuredValues["Chiller Efficiency (KW/TR)"]) 
            
            elif parameter == "Chiller Approch (Deg C)":
                SetValue = ((self.ChilledWaterValues["Chilled Water Temp. Outlet (Tc out)"][1]) -
                             (self.ChilledWaterValues["Refrigerent Temp. in Chiller"][1]))
            
            elif parameter == "Condensor Approch (Deg C)":
                SetValue = ((self.CoolingWaterValues["Refrigerent Temp. in Condensor"][1]) -
                             (self.CoolingWaterValues["Cooling Water Temp. Outlet (Tcw out)"][1]))
            
            
            ent.insert(0,SetValue)
   
            getValue = (ent.get())
            self.MeasuredValues[parameter] = getValue
            ent["state"] = "disabled"
        
        self.writeIntoExcel()
        
    #*************************************************************************************************
    #Function Name  : writeIntoExcel()
    #Description    : This function is used to validate the value entered in entry field
    #Parameters     : [action, index, value_if_allowed,prior_value, text, validation_type, trigger_type, widget_name]
    #*************************************************************************************************
    def writeIntoExcel(self):
        StartClm = 1
        self.excelObj.writeWorkSheet(self.StartRow,StartClm,self.StartRow-2)
        StartClm += 1
        for ele in self.InputValues:
            val = float(self.InputValues[ele][1])
            self.excelObj.writeWorkSheet(self.StartRow,StartClm,val)
            StartClm += 1
        for ele in self.ChilledWaterValues:
            val = float(self.ChilledWaterValues[ele][1])
            self.excelObj.writeWorkSheet(self.StartRow,StartClm,val)
            StartClm += 1
        for ele in self.CoolingWaterValues:
            val = float(self.CoolingWaterValues[ele][1])
            self.excelObj.writeWorkSheet(self.StartRow,StartClm,val)
            StartClm += 1
        for ele in self.OutputValues:
            val = float(self.OutputValues[ele])
            self.excelObj.writeWorkSheet(self.StartRow,StartClm,val)
            StartClm += 1
        for ele in self.MeasuredValues:
            val = float(self.MeasuredValues[ele])
            self.excelObj.writeWorkSheet(self.StartRow,StartClm,val)
            StartClm += 1
        
        self.StartRow += 1
        self.excelObj.saveExcel()
        
    #*************************************************************************************************
    #Function Name  : validate()
    #Description    : This function is used to validate the value entered in entry field
    #Parameters     : [action, index, value_if_allowed,prior_value, text, validation_type, trigger_type, widget_name]
    #*************************************************************************************************
    def validate(self, action, index, value_if_allowed,
                       prior_value, text, validation_type, trigger_type, widget_name):
        
        if not value_if_allowed:
            return True
        
        if value_if_allowed:
            try:
                float(value_if_allowed)
                return True
            except ValueError:
                return False
        else:
            return False 

if __name__ == "__main__":
    Efficiency()
    