# CoFC
  - Release scripts for to demonstrate parsing of xml and excel files and transformation of tabular data into new format, for this case to xml.
  - Test excel file (not included in repo) is a mask house released file with tabled extractable data. 
  - Text xml files (not included in repo) are mask house generated file where we selectively extract information and create a stripped down version of a new xml based on a configuration file provided by user.
    
Contents:
  **CoFC-Alpha-Demo/tree/main/CoFCLib/**
    COFCExcelCopyProtect.py - Copy protected worksheed to new excel file.
    COFCExcelCriticalDimensionWorksheetReader.py - Reading of specific worksheet and data transfomation to xml.
    COFCExcelPhaseShiftTransmissionWorksheetReader.py - Reading of specific worksheet and data transfomation to xml.
    COFCExcelRegistrationWorksheetReader.py - Reading of specific worksheet and data transfomation to xml.
    COFCExcelStatisticsWorksheetReader.py - Reading of specific worksheet and data transfomation to xml.
    COFCXMLParser.py - Reading of XML  and data transfomation to xml.
    plmget.py - Download file from webservice sample.
    
  **CoFC-Alpha-Demo/tree/main/config**
    * All contents - sample json config for all xpaths available and instructions.

  **CoFC-Alpha-Demo/blob/main/cofc_gui.py**
    * Sample basic gui for functional gui based demo.

    
