import os
import io
import time
import configparser
import re

import win32com.client as win32

def getCompanyList(SchooListData):
    companyList = list(map(lambda x: x["Short Name"], SchooListData))
    companyList.sort()
    return companyList

def getCompanyInfo(Company, CompanyData):
    for row in CompanyData:
        if row["Short Name"] == Company:
            return row

def getLineItemPresets(CompanyListData):
    lineItems = [x["Line Item"].strip() for x in CompanyListData["services"]]
    return lineItems

def getServiceDescriptionPresets(CompanyListData):
    descriptions = [x["Description"].strip() for x in CompanyListData["services"]]
    return descriptions

def getLastQuoteNumber(directory):
    if not anyQuotesAtThisPath(directory): #if no quotes found in folder
        return 0

    pdfs = list(filter(lambda x: x.endswith(".pdf") and (x.startswith("S") or x.startswith("s")),
                       os.listdir(directory)))  # ensure that all quotes in folder are pdf and begin with "s"

    numbers = list(map(lambda x: int(x[1:x.index(" ")]), pdfs))
    numbers.sort(reverse=True)
    return numbers[0] + 1

# write with Win32com
def writeToQuoteTemplate(company_name, guiData):
    cwd = os.getcwd()
    quotesDirectory = guiData['quotes dir']
    quoteTemplatePath = os.path.join(cwd,"src\\Quote Template.xlsx")
    quoteNumber = guiData['quote number']
    pdfFileName = r"{}\S{} {} Quote - {}.pdf".format(quotesDirectory, quoteNumber, company_name,
                                                         guiData['company data']['Full Name'])
    xlFileName = r"{}\S{} {} Quote - {}.xlsx".format(quotesDirectory, quoteNumber, company_name,
                                                         guiData['company data']['Full Name'])

    def openWorkbook(xlapp, xlfile):
        try:
            xlwb = xlapp.Workbooks(xlfile)
        except Exception as e:
            try:
                xlwb = xlapp.Workbooks.Open(xlfile)
            except Exception as e:
                print(e)
                xlwb = None
        return (xlwb)

    try:
        excel = win32.Dispatch('Excel.Application')

        wbsOpen = excel.Workbooks.Count # number of currently open excel windows
        protectedWBsOpen = excel.ProtectedViewWindows.Count # number of currently open protectedWB excel windows

        if wbsOpen or protectedWBsOpen: # if an unrelated excel file is already open, prevent it from closing when running this script
            excel.Visible = True

        wb = openWorkbook(excel, quoteTemplatePath) # create workbook object
        ws = wb.Worksheets("Template") # create worksheet object

        #### input header info
        ws.Range("F3").Value = guiData['date']
        ws.Range("F5").Value = f"Quote #S{quoteNumber}"
        ws.Range("D11").Value = guiData['company data']['Contact name']
        ws.Range("D12").Value = guiData['company data']['Full Name']
        ws.Range("D13").Value = guiData['company data']['Address']
        ws.Range("D14").Value = guiData['company data']['City, State, Zip']
        ws.Range("D15").Value = guiData['company data']['Phone']
        ws.Range("D16").Value = guiData['company data']['Contact email']

        #### input line item info
        count = 0
        for i in range(20, 24):
            ws.Cells(i, 2).Value = guiData['line items'][count]["description"]
            ws.Cells(i, 4).Value = guiData['line items'][count]["quantity"]
            ws.Cells(i, 5).Value = guiData['line items'][count]["unit price"]
            count += 1

        #### input discount and tax info
        ws.Range("F33").Value = float(guiData['discount'])
        ws.Range("F35").Value = float(guiData['tax']) / 100

        #### save and close
        wb.SaveAs(xlFileName)  # output xlsx file
        wb.SaveAs(pdfFileName, FileFormat=57)  ### output pdf

        wb.Close(True)

        ### check for existing
        createdBothFiles = wereFilesCreated(pdfFileName,xlFileName)

        if createdBothFiles:
            return {"pdfFileName":pdfFileName,"xlFileName":xlFileName}
        else: 
            return None

        # excel.Quit

    except Exception as e:
        print(e)

    finally:
        # RELEASES RESOURCES
        ws = None
        wb = None
        excel = None


def hasQuoteNumberBeenUsed(directory,quoteNum):
    for fileName in os.listdir(directory):
        if quoteNum in fileName:
            fileNumber = re.search(r"S\d+", fileName).group()[1:]
            if quoteNum == fileNumber:
                return True
    return False

def doesPathExist(directory)->"boolean":
    return os.path.exists(directory)

def anyQuotesAtThisPath(directory):
        try:
            for file in os.listdir(directory):
                if "quote" in file.lower():
                    numString = file[1:file.index(" ")] #extract the quote number if it exist
                    if numString.isnumeric() and (file.endswith(".pdf") or file.endswith(".xlsx")): #if it is an actual number and is a pdf or excel file
                        return True #found a file that matches conditions
        except:
            return False
        return False #no file found that meets conditions


def pathCheck(directory):
    if not doesPathExist(directory):
        return "THIS PATH DOES NOT EXIST"
    elif not anyQuotesAtThisPath(directory):
        return "THERE ARE NO QUOTES SAVED AT THIS PATH"
    else:
        return ""

def wereFilesCreated(file1,file2):
    totalTime = 0
    timeIncrement = .25
    while totalTime < 5: # seconds
        if os.path.exists(file1) and os.path.exists(file2):
            return True
        time.sleep(timeIncrement)
        totalTime +=  timeIncrement
    return False

def wasFileCreated(file):
    totalTime = 0
    timeIncrement = .25
    while totalTime < 5: # seconds
        if os.path.exists(file):
            return True
        time.sleep(timeIncrement)
        totalTime +=  timeIncrement
    return False



def deleteFile(filePath):
    os.remove(filePath)