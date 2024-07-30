from googleapiclient.discovery import build
import tkinter as tk
import os
from datetime import datetime as dt
import pandas as pd


# Google Sheets API
#Credentials
from google.oauth2 import service_account
# Enter the path to where the auth key is
os.chdir(r"D:\workingrepos\expense sheet")
# enter the file name of the keys
SERVICE_ACCOUNT_FILE = 'keys_sheet-expense-auto-input.json'
SCOPES = ["https://www.googleapis.com/auth/spreadsheets"]

creds = None
creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# The ID and range of the spreadsheet.
SPREADSHEET_ID = "1voUVTsBWJKPs7tqswDfrVDgxeDjH2YkP-04Y8ora-9A" # Finan-Sial[Expense Table]

service = build("sheets", "v4", credentials=creds)
# Call the Sheets API
sheet = service.spreadsheets()


#reading the sheets with API
expense_table = (
  sheet.values()
  .get(spreadsheetId=SPREADSHEET_ID, range="Expense Table!A:G")
  .execute()
)

# Function Definitions
def insert_row(rownum:int,rowvals:list):
  (
    sheet.values()
    .update(
        spreadsheetId=SPREADSHEET_ID,
        range=f"Expense Table!A{rownum}:G{rownum}",
        valueInputOption='USER_ENTERED',
        body={'values':[rowvals]})
        .execute()
  )
  return

def value_queries(df:pd.DataFrame):
  values = []
  for col in df.columns:
    match col:
      case "Date":
        entry = input_Date()
        
      case "Category":
        entry = input_Category()
      
      case "Method":
        entry = input_Method()
      
      case "Name":
        entry = input(f"Enter the item name: ")
        print("")
      
      case "Amount":
        entry = input(f"Enter the Amount in € (comma->'.'): ")
        entry = float(entry) if entry != "" else 0.0
        print("")
      
      case "Marke":
        entry = input(f"Enter the brand / store: ")
        print("")
      
      case "Notes":
        entry = input(f"Enter any additional notes: ")
        print("")
    
    values.append(entry)
  return values

def input_Date():
  print(f"Enter a Date with Format 'dd.mm.yy', 'today' or enter to copy last entry")
  entry = input(f"here: ")
  if entry == "today":
    return f"{dt.today().day}.{dt.today().month}.{dt.today().year}"
  elif entry == "":
    if first_run:
      print("Date from last entry will be assumed\n")
      return df_table_vals.iloc[-1]["Date"][4:]
    else:
      return CURR_VALS[0]
  else:
    return entry
  
def input_Category():
  cat_list = ["Social","Groceries","Household","Clothing", "Misc", "Snacking",
            "School / Work", "Health/Cosmetics", "Emergency",
            "Hobby","Utilities ","Transportation", "Financial"]
  cat_dict = dict(zip(  
                      list(range(len(cat_list))),
                      cat_list))
  
  print(f"Select a category:")
  for key, value in cat_dict.items():
    print(f"{key}: {value}")
  entry = input(f"Enter number here: ")
  print("")
  return cat_dict[int(entry)] if entry != "" else ""

def input_Method():
  method_list = ["Card DB", "Payback PAY", "Card Commerz",
               "Card N26", "Cash", "Paypal", "Transfer"]
  method_dict = dict(zip(
                        list(range(len(method_list))),
                        method_list))
  
  print(f"Select a payment method:")
  for key, value in method_dict.items():
    print(f"{key}: {value}")
  entry = input(f"Enter number here: ")
  print("")
  if entry != "":
    return method_dict[int(entry)]
  else:
    if first_run:
      return df_table_vals.iloc[-1]["Method"]
    else:
      return CURR_VALS[2]

first_run = True

#Saving the table from the sheet as a DataFrame for further needs
expense_table_vals = expense_table.get("values", [])
df_table_vals = pd.DataFrame(expense_table_vals[1:],columns=expense_table_vals[0])
newrownum = len(expense_table_vals)+1   #tracking how far the sheet are already filled


# running the functions
while True:
    CURR_VALS = value_queries(df_table_vals)
    insert_row(newrownum,CURR_VALS)
    newrownum += 1

    cont = input("Do you want to add another row? (yes/no): ")
    first_run = False
    if cont.lower() != 'yes':
        break