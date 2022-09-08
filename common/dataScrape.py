import os
import io
import pprint
import time
import configparser
import re

import win32com.client as win32

def testImport():
    print("datasCrape inported")

def getCompanyList(SchooListData):
    companyList = list(map(lambda x: x["Short Name"], SchooListData))
    companyList.sort()
    return companyList

def getCompanyInfo(Company, CompanyData):
    for row in CompanyData:
        if row["Short Name"] == Company:
            return row
# pprint.pprint(getCompanyInfo(Company, CompanyData))

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

#write with openpyxl(no pdf output)
# def writeToQuoteTemplate(guiData):
#     cwd = os.getcwd()
#     quotesDirectory = guiData['quotes dir']
#     quoteTemplatePath = os.path.join(cwd,"src\\quote-template.xlsx")
#     quoteNumber = guiData['quote number']

#     xlFileName = r"{}\S{} {} Quote - {}.xlsx".format(quotesDirectory, quoteNumber, company_name,
#                                                         guiData['company data']['Full Name'])
#     try:
#         wb = load_workbook(quoteTemplatePath)
#         ws = wb.active
#
#         #### input header info
#         ws["F3"] = guiData['date']
#         ws["F5"] = f"Quote #S{quoteNumber}"
#         ws["D11"] = guiData['company data']['Contact name']
#         ws["D12"] = guiData['company data']['Full Name']
#         ws["D13"] = guiData['company data']['Address']
#         ws["D14"] = guiData['company data']['City, State, Zip']
#         ws["D15"] = guiData['company data']['Phone']
#         ws["D16"] = guiData['company data']['Contact email']
#
#         for count,i in enumerate(range(20, 26)):
#             description = guiData['line items'][count]["description"]
#             quantity = guiData['line items'][count]["quantity"]
#             unitPrice = guiData['line items'][count]["unit price"]
#             ws.cell(i, 2).value = description
#             ws.cell(i, 4).value = float(quantity) if bool(quantity) else None
#             ws.cell(i, 5).value = float(unitPrice) if bool(unitPrice) else None
#
#         #### input discount and tax info
#         ws["F33"] = float(guiData['discount'])
#         ws["F35"] = float(guiData['tax']) / 100
#
#         wb.save(xlFileName)
#         wb.close()
#
#         # createdBothFiles = wasFileCreated(xlFileName)
#         return xlFileName
#
#     except ValueError as e:
#         print(f"Error raised with creating xlsx file: {e}")
#         return None

# def xlsxFileToPDF(xlsxFile):
#     pdfFileName = xlsxFile.replace(".xlsx", "(TEMP).pdf")
#
#     try:
#         # jvmStarted = jpype.isJVMStarted()
#         # print(f"is JVM running: {bool(jpype.isJVMStarted)}")
#         if not jpype.isJVMStarted():
#             jpype.startJVM()
#
#         from asposecells.api import Workbook, FileFormatType, PdfSaveOptions
#
#         workbook = Workbook(xlsxFile)
#         workbook.calculateFormula()
#         saveOptions = PdfSaveOptions()
#         saveOptions.setOnePagePerSheet(True)
#         workbook.save(pdfFileName, saveOptions)
#
#         # jpype.shutdownJVM()
#         # print(f"is JVM running: {bool(jpype.isJVMStarted)}")
#         return pdfFileName
#     except Exception as e:
#         print(f"Error raised with pdf processing: {e}")
#
#
# def hideAsposePDFWatermark(existingPDF):
#     outputFileName = existingPDF.replace("(TEMP)","")
#
#     #start Stream
#     packet = io.BytesIO()
#
#     #create overlay
#     pdfOverlay = canvas.Canvas(packet,bottomup=0) #pagesize=letter,
#     pdfOverlay.setLineWidth(11)
#
#     #Match Color of header bar in quote template
#     pdfOverlay.setStrokeColorRGB(179/255, 4/255, 4/255)
#     y = 100
#     #draw a line stroke at same coords where watermark is placed
#     pdfOverlay.line(1,y,578,y)
#     pdfOverlay.save()
#
#     #move to the beginning of the StringIO buffer
#     packet.seek(0)
#     new_pdf = PdfFileReader(packet)
#
#     # read your existing PDF
#     existing_pdf_Obj = open(existingPDF, "rb")
#     existing_pdf = PdfFileReader(existing_pdf_Obj)
#     output = PdfFileWriter()
#
#     # overlay new pdf on the existing pdf page
#     page = existing_pdf.getPage(0)
#     page.mergePage(new_pdf.getPage(0))
#     output.addPage(page)
#
#     # finally, write "output" to a final new file
#     outputStream = open(outputFileName, "wb")
#     output.write(outputStream)
#
#     #close file objects
#     existing_pdf_Obj.close()
#     outputStream.close()
#
#     return outputFileName



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