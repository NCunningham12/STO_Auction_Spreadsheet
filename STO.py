from openpyxl import Workbook, load_workbook
import numpy as np
import pandas as pd

# wb = load_workbook('PythonTest.xlsx')
wb = Workbook()
ws = wb.active


ws.title = "SoCal Trucks Only LLC Auction Calculations"

print("Workbook Updated")

truckData = [["Auction Price", "Province", "Canadian Price (CAD)", "US Price (USD)", "Book (USD)", "Total US (USD)", "BOB"]]

for row in truckData:
  ws.append(row)

def BOB():
  original_price = ws['E2'].value
  flip_price = original_price
  bb_value = ws['G2'].value

  if ws['C2'].value >= 2020:
   flip_price += 2000
  
  if ws['D2'].value <= 50000:
    flip_price += 5000

  if ws['F2'].value == 3:
    flip_price += 20000
  elif ws['F2'].value == 2:
    flip_price += 10000

  prof_marg = flip_price - bb_value
  ws['I2'] = prof_marg

  prof_perc = (prof_marg / original_price) * 100

  ws['J2'] = prof_perc

  ws['H2'] = flip_price

wb.save('STO_BOB.xlsx')