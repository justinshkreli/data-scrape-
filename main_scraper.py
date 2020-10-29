import requests
from bs4 import BeautifulSoup as bsoup
import json
import re
import pandas as pd
import xlsxwriter
import xlrd
from datetime import datetime
from collections import OrderedDict

#A program that webscrapes stock prices from companies provided by the user.
#An excel sheet collects a the ticker symbol for a stock, the stock's price, 
#the timestamp at which the stock's price is scraped, the last timestamp at which the stock's price is scraped,
# and the growth percentage of the price since the last time it was collected. 

#following headers enable program to scrape bloomberg
#source:
#https://www.reddit.com/r/learnpython/comments/f76gg6/cant_scrape_bloomberg_price/
headers_bloomberg = {
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_11_6) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/56.0.2924.87 Safari/537.36',
    'referrer': 'https://google.com',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'en-US,en;q=0.9',
    'Pragma': 'no-cache'}

#find ticker symbols of each company inputted by user
def find_all_stock_symbols():
    stock_symbols = []
    print("Add a company to a list of queries. Type 'stop' to finish your list.")
    company_query = input("Which company are you interested in: ")
    
    while company_query != "stop":
        new_string = ""
        for character in company_query:
            if character == " ":
                new_string += "+" #prepare for url
            else:
                new_string += character
        symbol = ""
        url = "https://www.marketwatch.com/tools/quotes/lookup.asp?siteID=mktw&Lookup="
        url += new_string
        url += "&Country=us&Type=All"
        r = requests.get(url) 
        soup = bsoup(r.content, 'html5lib') #get html of query
        try: #query leads to page containing a bunch of results
            table_of_results = soup.find("div", {"class": "results"}).find("tbody").find("tr")
            first_entry = table_of_results.find("td", {"class": "bottomborder"})
            parse_this = first_entry.decode_contents()
            
            first_index_symbol = get_start_index_symbol(parse_this)

            sub_parse_this = parse_this[first_index_symbol:]
            symbol = get_symbol(sub_parse_this)
        except: #query leads to page containing one result
            meta_tag = soup.find('meta', attrs={'name': 'tickerSymbol'})
            symbol = meta_tag["content"]

        stock_symbols.append(symbol)
        company_query = input("Which company are you interested in: ")
    return stock_symbols

#get the index at which ticker symbol starts in string (from html file)
def get_start_index_symbol(parse_this):
    first_index_symbol = 0
    for index1 in range(len(parse_this)):
        if parse_this[index1] == "\"" and index1 <= len(parse_this)-1:
            if parse_this[index1+1] == ">":
                first_index_symbol = index1+2
                break #found the index
    return first_index_symbol

#extract ticker symbol from string (that comes from html file)
def get_symbol(sub_parse_this):
    symbol = ""
    for character in sub_parse_this:
            if character != "<":
                symbol += character
            else:
                break
    return symbol

#get prices for each ticker symbol found
def get_stock_prices(stock_symbols):
    stock_prices = OrderedDict()
    stock_prices["stocks"] = stock_symbols
    stock_prices["prices"] = [] 
    stock_prices["now_time"] = []
    for stock_symbol in stock_symbols:
        url = "https://www.bloomberg.com/quote/"
        url += stock_symbol
        url += ":US"
        r = requests.get(url, headers=headers_bloomberg)
        soup = bsoup(r.content, 'html5lib')

        stock_price = 0
        for part in soup.select('span[class*="priceText"]'): #dig in html for the part with stock's price
            stock_price = part.get_text()
        stock_prices["prices"].append(str(stock_price.replace(",",""))) #only get the number characters

        timestamp = datetime.now()
        current_time = timestamp.strftime("%H:%M:%S") #time right after we retrieve price
        stock_prices["now_time"].append(current_time)
    return stock_prices 

#update the whole excel file 
def apply_changes(stock_prices):
    stock_prices["prev_time"] = []
    stock_prices["growth_percents"] = []
    sheet = get_data_from_excel()
    for index,stock1 in enumerate(stock_prices["stocks"]):
        already_in_sheet = False
        for stock2 in range(1,sheet.nrows): 

            if stock1 == sheet.cell_value(stock2,1): #company data has been recorded
                already_in_sheet = True

                previous_price = sheet.cell_value(stock2,2)
                new_price = stock_prices["prices"][index]
                growth = ((float(new_price)-float(previous_price))/float(previous_price))*100.0
                stock_prices["growth_percents"].append(str(growth)) #make sure to append in string format

                stock_prices["prev_time"].append(sheet.cell_value(stock2,3))

        if not already_in_sheet: #never looked up company
            stock_prices["growth_percents"].append("N/A")
            stock_prices["prev_time"].append("N/A")
    return stock_prices

#open datasheet from excel
def get_data_from_excel():
    location = ("data_collector.xlsx")
    excel_file = xlrd.open_workbook(location) 
    sheet = excel_file.sheet_by_index(0) 
    return sheet

def display_ticker(data):
    print(data)

#rewrite excel sheet
def export_to_excel(stock_prices):
    data = pd.DataFrame.from_dict(stock_prices)
    display_ticker(data)
    datatoexcel = pd.ExcelWriter("data_collector.xlsx", engine="xlsxwriter")
    data.to_excel(datatoexcel, sheet_name="Sheet1")
    datatoexcel.save()

#function executes all helper functions in order
def execute_program():
    stock_symbols = find_all_stock_symbols()
    stock_prices = get_stock_prices(stock_symbols)
    stock_data = apply_changes(stock_prices)
    export_to_excel(stock_data)

def main():
    execute_program()

if __name__ == "__main__":
    main()
