from openpyxl import Workbook, load_workbook
import numpy as np
import pandas as pd

wb = load_workbook('STO_BOB.xlsx')
ws = wb.active


ws.title = "STO LLC Auction Calculations"

truckCategories = [["Auction Price", "Province", "Canadian Price (CAD)", "US Price (USD)", "Book (USD)", "Total US (USD)", "BOB"]]

# for row in truckCategories:
#   ws.append(row)

# INPUTS: Auction Price, Province, US Price, Book
# OUTPUTS: Canadian Price, Total US, BOB

def BOB():
  auction_price = ws['A2'].value
  province = ws['B2'].value
  # us_price = ws['E2'].value
  book = ws['C2'].value
  bob = ws['G2'].value
  exchange = ws['I2'].value

  print("Province: " + province)

  if province.lower() == 'alberta' or province.lower() == 'bc':
    print("WEST:")
    auction_price += 1000
    auction_price *= 1.05
    canadian_price = auction_price
    ws['D2'] = canadian_price
    print("Canadian Price: " + str(canadian_price))

    us_price = round(canadian_price * exchange)
    ws['E2'] = us_price
    print("US Price: " + str(us_price))

    us_price += 2300
    us_price *= 1.01
    total_us_price = round(us_price, 2)
    ws['F2'] = total_us_price
    print("Total US Price: " + str(total_us_price))
  elif province.lower() == 'ontario':
    print('EAST:')
    auction_price += 1000
    auction_price *= 1.13
    canadian_price = auction_price
    ws['D2'] = canadian_price
    print("Canadian Price: " + str(canadian_price))

    us_price = round(canadian_price * exchange)
    ws['E2'] = us_price
    print("US Price: " + str(us_price))

    us_price += 3400
    us_price *= 1.01
    total_us_price = round(us_price, 2)
    ws['F2'] = total_us_price
    print("Total US Price: " + str(total_us_price))

  bob = round(total_us_price - book, 2)
  ws['G2'] = bob
  print("BOB: " + str(bob))

BOB()

print("✅ WORKBOOK UPDATED")

# how to iterate by col-name
ColNames = {}
Current  = 0
for COL in ws.iter_cols(1, ws.max_column):
    ColNames[COL[0].value] = Current
    Current += 1

 ## Now you can access by column name
for row_cells in ws.iter_rows(min_row=1, max_row=4):
    print(row_cells[ColNames['Province']].value) 

wb.save('STO_BOB.xlsx')