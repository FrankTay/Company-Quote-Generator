# if __name__ == "__main__":
#     # import os
#     # import xlwings as xw

#     # # Initialize new excel workbook
#     # book = xw.Book()
#     # sheet = book.sheets[0]
#     # sheet.range("A1").value = "dolphins"

#     # # Construct path for pdf file
#     # current_work_dir = os.getcwd()
#     # pdf_path = os.path.join(current_work_dir, "workbook_printout.pdf")

#     # # Save excel workbook to pdf file
#     # print(f"Saving workbook as '{pdf_path}' ...")
#     # book.api.ExportAsFixedFormat(0, pdf_path)

#     # # Open the created pdf file
#     # print(f"Opening pdf file with default application ...")
#     # os.startfile(pdf_path)

# import json
# json_file = "dummy_client_info.json"
# with open(json_file) as f:
#    data = json.load(f)

# print(data["company data"])


import os, sys

print(os.path.dirname(sys.executable))