#**********************************
# Standard and Project libraries
#**********************************
import sys
sys.path.append("../")

from Configuration import config


#*************************************************************************************************
#Class Name     : TextFileOperations()
#Description    : This class is used to perform the text file operations
#Parameters     : NA
#*************************************************************************************************
class TextFileOperations:
    def __init__(self,TextFile):
        self.textFile = TextFile
        self.fileObj = None
        
    #*************************************************************************************************
    #Function Name  : openFile()
    #Description    : This function is used to open the text file in write mode
    #Parameters     : NA
    #*************************************************************************************************
    def openFile(self):
        self.fileObj = open(self.textFile,"w")
        
    #*************************************************************************************************
    #Function Name  : writeTextFile()
    #Description    : This function is used to write the data in text file
    #Parameters     : [ data ]
    #*************************************************************************************************
    def writeTextFile(self,data):
        totalLenght = 50
        self.fileObj.write("="*60)
        self.fileObj.write("\nBest Efficiency = ")
        self.fileObj.write(str(data["Chiller Efficiency - (KW/TR)"])+" (KW/TR)\n")
        self.fileObj.write("="*60)
        self.fileObj.write("\n")
        
        for ele in data:
            if ele != "Chiller Efficiency - (KW/TR)":
                eleLength = len(ele)
                padding = totalLenght-eleLength
                tempEle = ele+" "*padding
                self.fileObj.write(tempEle+":")
                self.fileObj.write(str(data[ele]))
                self.fileObj.write("\n")
            else:
                pass
            
    #*************************************************************************************************
    #Function Name  : closeTextFile()
    #Description    : This function is used to save and close the text file
    #Parameters     : []
    #*************************************************************************************************    
    def closeTextFile(self):
        self.fileObj.close()
        
        
        
if __name__ == "__main__":
    tempdata = {'Sr. No.': 4, 'Chilled Water Flow - (m^3/hr)': 29.0, 'Chiller Power - (Kw)': 85.0, 'Ambient Dry Build Temp.(DBT) - (Deg C)': 36.0, 'Ambient Wet Build Temp.(WBT) -  (Deg C)': 38.0, 'Relative Humidity (RH) - (%)': 33.0, 'Chilled Water Temp. Inlet (Tc in) - (Deg C)': 197.0, 'Chilled Water Temp. Outlet (Tc out) - (Deg C)': 153.0, 'Refrigerent Temp. in Chiller (Deg C)': 239.0, 'Cooling Water Temp. Inlet (Tcw in) - (Deg C)': 230.0, 'Cooling Water Temp. Outlet (Tc out) - (Deg C)': 253.0, 'Chiller Range - (Deg C)': 44.0, 'Condensor Range - (Deg C)': 23.0, 'Chiller Operating TR - (TR)': 422.1019309908199, 'Chiller Efficiency - (KW/TR)': 0.2013731607445516, 'COP': 17.465088132878982, 'Chiller Approch - (Deg C)': -22.0, 'Condensor Approch - (Deg C)': -14.0}
    obj1 = TextFileOperations(config.TEXT_FILE)
    obj1.openFile()
    obj1.writeTextFile(tempdata)
    obj1.closeTextFile()
    
    
    