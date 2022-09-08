import os,pprint
import configparser

import common.dataScrape as dataScrape


def getDirectory():
	pass


def writeConfig(userInputs):
	parser = configparser.ConfigParser()
	parser.add_section('Header')
	parser.set('Header', 'quoteDirectory', userInputs['quotes dir']) #quote dir text
	parser.set('Header', 'company', userInputs['company']) #quote dir text

	parser.add_section('Line Item Header')   
	parser.set('Line Item Header', 'discount', userInputs['discount'])
	parser.set('Line Item Header', 'tax', userInputs['tax'])

	for index,item in enumerate(userInputs["line items"]):
		parser.add_section(f'Line{index + 1}')
		parser.set(f'Line{index + 1}', 'description', item["description"])
		parser.set(f'Line{index + 1}', 'qty', item["quantity"])
		parser.set(f'Line{index + 1}', 'unit price', item["unit price"])

	fp=open('config.ini','w')
	parser.write(fp)
	fp.close()
# writeConfig()

def readConfig(configFile):
	configData = {}
	parser = configparser.ConfigParser()
	parser.read(configFile)
	configData["quotes dir"] = parser["Header"]["quotedirectory"]
	configData["company"] = parser["Header"]["company"]
	configData["discount"] = parser["Line Item Header"]["discount"]
	configData["tax"] = parser["Line Item Header"]["tax"]
	lineItemText = []
	for row in range(1,7):
		LineNum = f"Line{row}"
		rowDict = {}
		rowDict["description"] = parser[LineNum]["description"]
		rowDict["qty"] = parser[LineNum]["qty"]
		rowDict["unit price"] = parser[LineNum]["unit price"]
		lineItemText.append(rowDict)

	configData["line items"] = lineItemText
	return configData

# readConfig("config.ini")

cwd = os.getcwd()
configFile = os.path.join(cwd,"config.ini")
#ON OPEN
#check for config file
def onOpen(configFile):
	configData = {}
	defaultQuotesDir = "D:// quotes dir"
	if dataScrape.doesPathExist(configFile): # if config exists
		print("File EXISTS")
		#read values
	else:

		print("no file")

# onOpen(configFile)

def onClose(userInputs):
	writeConfig(userInputs)
	#write to file

