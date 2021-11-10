from re import L
from gspread.models import Spreadsheet, Worksheet
from oauth2client.service_account import ServiceAccountCredentials
import gspread
import pandas as pd
import mysql.connector

###DESCRIPTION##
print("Welcome To Gspread-MySql-PowerBi Automation Script")
print("**************************************************")

###AUTHENTICATION###
gc = gspread.oauth()

###LIST SPREADSHEETS IN DRIVE###
Sheet_list = gc.list_spreadsheet_files()
Sheets_new = dict()
print(Sheet_list)
for i,sheet in enumerate(Sheet_list):
    print(str(i)+". "+sheet.get('name'))
    Sheets_new[str(i)] = sheet.get('name')
###GETTING USER SPECIFIED SPREADSHEET FROM LIST###

Acquire_Worksheet = input("Input your spreadsheet ID shown above:")

worksheet = gc.open(Sheets_new.get(str(Acquire_Worksheet)))
Worksheet_list = worksheet.worksheets()
select_worksheet = dict()

for Worksheet in Worksheet_list:
  select_worksheet[str(Worksheet.id)] = Worksheet.title
  print(str(Worksheet.id)+': '+Worksheet.title)

selection = input("Choose Your Worksheet ID: ")
Selected_Sheet = worksheet.worksheet(select_worksheet.get(str(selection)))
###STORING SHEET DATA IN A LOCAL DICTIONARY###
dataset = Selected_Sheet.get_all_records()

##SETTING UP SQL CONNCETOR##
mydb = mysql.connector.connect(
  host="192.168.186.128",
  user="root",
  password="my-secret-pw",
  database="test_db"
)

cursor = mydb.cursor()

choice = input("Create Table? y or n ?")
if choice == 'y':
    table_name = input("Please input table name:")
    query = "CREATE TABLE {} (name VARCHAR(30),population INT(32));".format(table_name)
    cursor.execute(query)
else:
    table_name = 'afghanistan'
##SETTING UP QUERY STRING##
query_string = ''
for record in dataset:
    name = record.get('name')
    population = record.get('population')
    query_string = query_string + "('{}','{}'),".format(name,population)

query_string = query_string[:-1]

###CONCATINATING QUERY###
query = "INSERT INTO {} VALUES {};".format(table_name,query_string)
print(query)
menu_option = input("Execute Query? Y/N")
if choice == 'y' or choice =='Y':
  cursor.execute(query)
  mydb.commit()
else:
  print("Query Aborted")
cursor.close()

###VIEW QUERY##
