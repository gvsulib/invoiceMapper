import sys
import csv
from datetime import datetime
from collections import Counter

spendCategoryMapping = {"CC0141-SC0049-FD100-P44100":	"SC0049",
"CC0141-SC0050-FD100-P44100":	"SC0050",
"CC0141-SC0146-FD100-P44100":	"SC0146",
"CC0141-SC0147-FD100-P44100":	"SC0147",
"CC0141-SC0148-FD100-P44100":	"SC0148",
"CC0141-SC0177-FD100-P44100":	"SC0177",
"CC0141-SC0229-FD100-P44100":	"SC0229",
"CC0141-SC0230-FD100-P44100":	"SC0230",
"CC0141-SC0231-FD100-P44100":	"SC0231",
"CC0141-SC0232-FD100-P44100":	"SC0232",
"CC0141-DS0125-FD100-P44100":	"SC0049"}


mapping = {"Invoice Key":"folio invoice number",
           "Lib Document Number": "",
           "Supplier":"accounting code",
           "Invoice Date":"invoice date",
           "Invoice Received Date": "approved date",
           "Supplier Invoice Number": "vendor invoice number",
           "External PO Number": "",
           "Memo": "",
           "Original Supplier Invoice Number":"vendor invoice number",
           "Line Order":"invoice line number",
           "Item Description":"description (title)",
           "Line Memo": "",
           "Spend Category":"external account number",
           "Extended Amount":"total",
           "Designation":"",
           "Fund":"",
           "Cost Center":"",
           "Program":"",
           "Gift":"",
           "Grant":"",
           "Project":"",
           "Activity":""
           }

def formatDate(date):
  year = ""
  month = ""
  day = ""
  splitDate = date.split("/")
  if len(splitDate[0]) < 2:
    month = "0" + splitDate[0]
  else:
    month = splitDate[0]
  if len(splitDate[1]) < 2:
    day = "0" + splitDate[1]
  else:
    day = splitDate[1]
  if len(splitDate[2]) < 4:
    year = "20" + splitDate[2]
  else:
    year = splitDate[2]
  return year + "-" + month + "-" + day

def checkSyntax(csvFile, fieldNames, requiredFields):
  print("Checking file syntax...")
  errors = []
  # check the file for required columns

  #normalize all field names to lower case and strip problematic characters
  for i in range (len(fieldNames)):
    fieldNames[i] = fieldNames[i].lower().trim().replace("\ufeff","").replace("\"", "")
  
  #check if required columns are present
  for name in requiredFields:
    if name not in fieldNames:
      errors.append(" error: Required column not present:\n" + name + " not found in column headers. Check the names or generate a new file with the missing column present.")
  # check for duplicate column names in required columns
  countList = Counter(fieldNames)
  for fieldname in countList:
    if fieldname not in requiredFields:
      continue
    if countList[fieldname] > 1:
      errors.append("Error:  There are duplicate column names in this file.  Column names must be unique.\nCheck for duplicates of column name: " + fieldname + " and remove or rename.")
  #check for values in required columns
  emptyData = []

  for row in csvFile:
    for key, value in row.items():
      cleanKey = key.lower().replace("\ufeff","")
      if cleanKey not in requiredFields:
        continue
      if row[key] == "" and key not in emptyData:
        emptyData.append(key)

  if len(emptyData) > 0:
    for value in emptyData:
      errors.append("Required column " + value + " has empty values. Values are required in this field.  Check and fill in missing values and retry.")


  if len(errors) > 0:
    print("Errors found in csv file, aborting...\n")
    for msg in errors:
      print(msg + "\n")
    sys.exit()

try:
  filename = sys.argv[1]
except IndexError:
   print("No invoice filename provided, shutting down.")
   sys.exit()

print("Opening file " + filename + " For processing")
# initializing the titles and rows list

requiredFields = list(mapping.values())

requiredFields = list(set(requiredFields))

requiredFields.remove("")
invoiceData = []
fieldNames = []
try:
  with open(filename, 'r', newline="\n") as csvfile:
    csvReader = csv.DictReader(csvfile)
    fieldNames = csvReader.fieldnames
    for row in csvReader:
      invoiceData.append(row)
    csvfile.close
except FileNotFoundError:
   print("Cannot find file with filename: " + filename + " in current directory, check your input, shutting down.")
   sys.exit
except SyntaxError:
  print("Problem with structure of csv file:  File cannot be parsed, indicating the .CSV file is corrupt.\n")
  print("Try exporting a new csv file.")
  sys.exit

checkSyntax(invoiceData, fieldNames, requiredFields)

print("File retreived and parsed. Starting transformation.")
print("screening out pcard orders and johnson center orders")

workdayFile = []
for row in invoiceData:
  if row["Vendor code"] == "AMAZO" or row["External account number"].strip("\"") == "CC0159-FD620-P10000-EN655700" or row["Payment method"].lower() == "credit card" or row["Acquisitions units"].lower() == "library designated fund":
    continue

  mapped_dict ={}
  for key, value in row.items():
    key = key.lower().strip("\ufeff")

    if key == "invoice date" or key == "approved date":
      value = formatDate(value)
    elif key == "external account number":
      value = spendCategoryMapping[value.strip("\"")]
    elif key == "description (title)":
      value = value.strip("\"")
      value = '"' + value + '"'
    mapped_dict[key] = value
    WDrow = {}
  
  for key, element in mapping.items():
    if element == "":
      if key == "Fund":
        WDrow[key] = "FD100"
      elif key == "Cost Center":
        WDrow[key] = "CC0141"
      elif key == "Program":
        WDrow[key] = "P44100"
      elif key == "Lib Document Number":
        WDrow[key] = "LIB-" + mapped_dict["folio invoice number"]
      else:
        WDrow[key] = ""
    else:
      WDrow[key] = mapped_dict[element]
  workdayFile.append(WDrow)

with open("workdayfile.csv", "w", newline="\n") as outFile:
    fieldnames = list(mapping.keys())
    w = csv.DictWriter(outFile, fieldnames=fieldnames)
    w.writeheader()
    w.writerows(workdayFile)
    outFile.close()
print("Transformation finished, closing down.")
     