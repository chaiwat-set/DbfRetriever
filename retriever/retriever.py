#!/usr/bin/env python3

from datetime import datetime
from tkinter import *
from tkinter.filedialog import askopenfilename, asksaveasfilename
import dbf
import numpy
import pandas
import csv
import os

INDEX_TABLE = {"STAT": 0, "TRUCK": 1, "CARTYPE": 2, "COMPANY": 3, "PRODUCT": 4, "SUBCON": 5, "REMARK1": 6,
              "REMARK2": 7, "REMARK3": 8, "FACTOR": 9, "VATCASE": 10, "PRICE": 11, "RATE": 12, "VAT_R": 13,
              "TICKET1": 14, "DAYIN": 15, "TMIN": 16, "W1": 17, "TICKET2": 18, "DAYOUT": 19, "TMOUT": 20,
              "W2": 21, "ADJ_W1": 22, "ADJ_W2": 23, "ADJ_M": 24, "STAFF": 25, "PROCESS": 26, "PRINT1": 27,
              "PRINT2": 28, "SCALEIN": 29, "SCALEOUT": 30, "LINK": 31}
PRODUCT_TABLE = {"001": "รำสด", "002": "แกลบ", "003": "น้ำมันรำดิบ", "999": "แกลบดำ"}
SCRIPT_PATH = os.path.dirname(__file__)
COMPANY_FILE_PATH = SCRIPT_PATH + r"/company.xlsx"
PRODUCT_FILE_PATH = SCRIPT_PATH + r"/product.xlsx"
DEBUG_ENABLED = False
entries = list()

def get_possible_dayin_options(dbf_file_path):
  dbf_file = dbf.Table(dbf_file_path)
  dbf_file.open(dbf.READ_ONLY)
  possible_day_in_datetime = set()
  for record in dbf_file:
    possible_day_in_datetime.add(record[INDEX_TABLE["DAYIN"]].strftime("%Y/%m/%d"))
  dbf_file.close()
  return sorted(possible_day_in_datetime, reverse=True)


def convert_index_to_info(lstbox):
  lst = list()
  for index in lstbox.curselection():
    lst.append(lstbox.get(index))
  return lst


def search_data(dayins, dbf_file_path, result_box):
  selected_day_ins = convert_index_to_info(dayins)
  if DEBUG_ENABLED:
    print(sorted(selected_day_ins))
  search(sorted(selected_day_ins), dbf_file_path, result_box)


def search(dayins, dbf_file_path, result_box):
  entries.clear()
  dbf_file = dbf.Table(dbf_file_path)
  dbf_file.open(dbf.READ_ONLY)
  index = dbf_file.create_index(lambda rec: (rec.DAYIN))

  if DEBUG_ENABLED:
    print(f"Opening {COMPANY_FILE_PATH}...")
    print(f"Opening {PRODUCT_FILE_PATH}...")

  company_file = pandas.read_excel(COMPANY_FILE_PATH, sheet_name='Sheet1')
  product_file = pandas.read_excel(PRODUCT_FILE_PATH, sheet_name='Sheet1')
  for dayin in dayins:
    date = datetime.strptime(dayin, "%Y/%m/%d").date()
    match = index.search(match=(date,), partial=True)
    for row in match:
      entry = create_entry(row, company_file, product_file)
      entries.append(entry)
  dbf_file.close()

  if DEBUG_ENABLED:
    for entry in entries:
      print(f"\t{entry}")
    print(f"Total: {len(entries)}")

  result_box.config(state="normal")
  result_box.delete("1.0", END)
  count = 1
  for entry in entries:
    result_box.insert(END, f"Entry {count}:\n\t{entry[0]}, {entry[1]}, {entry[2]}, {entry[3]}, {entry[4]}, {entry[5]}, {entry[6]}, {entry[7]}, {entry[8]}\n\n")
    count += 1
  result_box.config(state="disabled")


def create_entry(row, company_file, product_file):
  entry = list()
  company_code = row[INDEX_TABLE["COMPANY"]].strip()
  company_name = lookup_code(company_code, company_file)
  product_code = row[INDEX_TABLE["PRODUCT"]].strip()
  product_name = lookup_code(product_code, product_file)
  day_in = lookup_day(row, "DAYIN")
  day_out = lookup_day(row, "DAYOUT")
  try:
    net_weight = abs(row[INDEX_TABLE["W2"]] - row[INDEX_TABLE["W1"]])
  except TypeError:
    net_weight = -1
  entry.append(product_name)
  entry.append(day_in)
  entry.append(day_out)
  entry.append(row[INDEX_TABLE["TRUCK"]].strip())
  entry.append(company_code)
  entry.append(company_name)
  entry.append(net_weight)
  entry.append(row[INDEX_TABLE["REMARK1"]].strip())
  entry.append(row[INDEX_TABLE["REMARK2"]].strip())
  entry.append(row[INDEX_TABLE["REMARK3"]].strip())
  return entry

def lookup_day(row, type):
  try:
    return row[INDEX_TABLE[type]].strftime("%Y/%m/%d")
  except AttributeError:
    return ""

def lookup_code(code, df):
  dataframe = df[df["CODE,C,10"] == int(code)]
  try:
    return dataframe["NAME,C,60"].loc[dataframe.index[0]]
  except (KeyError, IndexError):
    return f"UNDEFINED CODE '{code}'"


def create_listbox(root, name, width, data):
  frame = LabelFrame(root, text=name)
  y_scrollbar = Scrollbar(frame, orient=VERTICAL)

  lstbox = Listbox(frame, width=width, height=15, yscrollcommand=y_scrollbar.set, selectmode=EXTENDED, exportselection=False)
  y_scrollbar.config(command=lstbox.yview)
  y_scrollbar.pack(side=RIGHT, fill=Y)
  lstbox.pack()
  frame.pack(side=LEFT)

  for element in data:
    lstbox.insert(END, element)

  return lstbox


def save_entries_to_csv():
  file_type = [('All tyes(*.*)', '*.*'),("csv file(*.csv)","*.csv")]
  save_file_name = asksaveasfilename(initialfile = 'output.csv', defaultextension=file_type, filetypes=file_type)
  fields = ["Product", "Day-In", "Day-Out", "Truck", "Company Code", "Company Name", "Net Weight", "Remark 1", "Remark 2", "Remark 3"]
  with open(save_file_name, "w", newline='', encoding="utf-8") as f:
    writer = csv.writer(f)
    writer.writerow(fields)
    writer.writerows(entries)


def main():
  root = Tk()
  root.title("DBF Retriever")
  root.geometry("1000x280")
  root.resizable(0,0)

  dbf_file_path = askopenfilename(title="Open your wdata.dbf file")
  possible_dayin_options = get_possible_dayin_options(dbf_file_path)
  day_in_lstbox = create_listbox(root=root, name="Day-In Date (yyyy/mm/dd)", width=20, data=possible_dayin_options)
  result_box = Text(root, height=15, width=70)
  search_button = Button(root, text="Search", width=15, height=3, command=lambda:search_data(day_in_lstbox, dbf_file_path, result_box))
  search_button.pack(padx=5, side=LEFT)
  result_box.pack(padx=5, side=LEFT)
  save_button = Button(root, text="Save", width=15, height=3, command=lambda:save_entries_to_csv())
  save_button.pack(padx=5, side=LEFT)

  root.mainloop()

if __name__ == '__main__':
  main()