from openpyxl import Workbook, load_workbook
from openpyxl.styles import Alignment

wb = load_workbook('STO_BOB.xlsx')
ws = wb.active

currentCell = ws.cell('A1') #or currentCell = ws['A1']
currentCell.alignment = Alignment(horizontal='center')

ws.title = "STO LLC Auction Calculations"

truckCategories = [["ID", "Auction Price", "Province", "Canadian Price (CAD)", "US Price (USD)", "Book (USD)", "Total US (USD)", "BOB"]]

# for row in truckCategories:
#   ws.append(row)

# INPUTS: Auction Price, Province, US Price, Book
# OUTPUTS: ID, Canadian Price, Total US, BOB

def BOB(): 
  maxRow = ws.max_row + 1
  id = 0

  for row in range(2, maxRow):
    id += 1
    ws['A' + str(row)].value = id                                     # Creates a unique ID number for each vehicle for reference only

    auction_price = round(ws['B' + str(row)].value, 2)
    province = ws['C' + str(row)].value
    us_price = ws['F' + str(row)].value
    book = ws['D' + str(row)].value
    exchange = ws['K2'].value

    if province.lower() == 'alb' or province.lower() == 'bc':         # Checks if province is WEST
      auction_price += 1000
      auction_price *= 1.05                                           # Adds 1000 and then adds 5% of the new total
      canadian_price = round(auction_price, 2)
      ws['E' + str(row)] = round(canadian_price, 2)

      us_price = round(canadian_price * exchange, 2)                  # Converts the CAD to USD based on the input "Exchange rate"
      ws['F' + str(row)] = round(us_price, 2)

      us_price += 2300                                                
      us_price *= 1.01                                                # Adds 2300 and then adds 1% of the new total
      total_us_price = round(us_price, 2)  
      ws['G' + str(row)] = total_us_price                       
    elif province.lower() == 'ont':                                   # Checks if province is EAST
      auction_price += 1000
      auction_price *= 1.13                                           # Adds 1000 and then adds 13% of the new total
      canadian_price = round(auction_price, 2)                        
      ws['E' + str(row)] = round(canadian_price, 2)

      us_price = round(canadian_price * exchange)                     # Converts the CAD to USD based on the input "Exchange Rate"
      ws['F' + str(row)] = round(us_price, 2)

      us_price += 3400                                                
      us_price *= 1.01                                                # Adds 3400 and then adds 1% of the new total
      total_us_price = round(us_price, 2)
      ws['G' + str(row)] = total_us_price

    bob = round(total_us_price - book, 2)                             # Subtracts "Total US - Book" to find out the BOB
    ws['H' + str(row)] = bob


    # print("Auction Price: " + str(round(auction_price, 2)))
    # print("Canadian Price: " + str(canadian_price))
    # print("US Price: " + str(round(us_price, 2)))
    # print("Total US Price: " + str(total_us_price))
    # print("BOB: " + str(bob))

BOB()

print("✅ WORKBOOK UPDATED")
wb.save('STO_BOB.xlsx')