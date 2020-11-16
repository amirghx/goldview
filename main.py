import json
from urllib.request import urlopen
from xml.etree.ElementTree import parse
import requests
import xlwings as xw
import datetime


def shifter(List):
    output = []
    for item in List:
        temp = []
        temp.append(item)
        output.append(temp)
    return output


# Function for extracting trades value out of tadbir API
def get_val(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("TotalTradeValue")


# Function for extracting trades volume out of tadbir API
def get_vol(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("TotalNumberOfTrades")


# Function for extracting  Last trade price out of tadbir API
def get_Last_traded_price(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("LastTradedPrice")


# Function for extracting Last trade price out of tadbir API
def get_yesterday(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return int((x.get("HighAllowedPrice") + x.get("LowAllowedPrice")) / 2)


# Function for extracting Last trade price out of tadbir API
def get_NAV(url):
    response = requests.get('http://mdapi.tadbirrlc.com/api/NavDetails/' + url)
    data = response.text
    parsed = json.loads(data)
    x = parsed['NavDetails']

    return x.get("CancelNAV")


def get_min_traded(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("LowPrice")


def get_max_traded(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("HighPrice")


def get_max_day(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("HighAllowedPrice")


def get_min_day(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("LowAllowedPrice")


def get_bid_vol(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("BidAskFirstRow").get("BestSellQuantity")


def get_bid_price(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("BidAskFirstRow").get("BestSellPrice")


def get_ask_price(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("BidAskFirstRow").get("BestBuyPrice")


def get_ask_vol(url):
    response = requests.get('http://mdapi.tadbirrlc.com/API/symbol?$filter=SymbolISIN+eq+%27' + url + '%27')
    data = response.text
    parsed = json.loads(data)
    x = parsed['List'][0]

    return x.get("BidAskFirstRow").get("BestBuyQuantity")


lb_header = ['تغييرات (درصد)', 'قيمت (ريال)', 'نام', 'شماره']
var_url = urlopen('http://parsijoo.ir/api?serviceType=price-API&query=Gold')
xmldoc = parse(var_url)
lb_list = []
name_tag = []
price_tag = []

for item in xmldoc.iterfind('sadana-services/price-service/item'):
    name = item.findtext('name')
    price = item.findtext('price')
    change = item.findtext('change')
    percent = item.findtext('percent')

    lb_list += [(percent + change, price, name)]
    name_tag.append(name)
    price_tag.append(price)

# Define writer for python
wb = xw.Book('Goldview.xlsx')
worksheet = wb.sheets('Sheet1')

# Insert value for each ISIN in excel
worksheet.range('C2').value = get_val('IRK1K00500C1')
worksheet.range('C4').value = get_val('IRK1K00701C1')
worksheet.range('C6').value = get_val('IRK1K00499C1')
worksheet.range('C8').value = get_val('IRK1K00600B1')
worksheet.range('C10').value = get_val('IRTKLOTF0001')
worksheet.range('C12').value = get_val('IRTKMOFD0001')
worksheet.range('C14').value = get_val('IRTKKIAN0001')
worksheet.range('C16').value = get_val('IRTKZARF0001')
print('phase1 done')
# Insert volume for each ISIN in excel
worksheet.range('D2').value = get_vol('IRK1K00500C1')
worksheet.range('D4').value = get_vol('IRK1K00701C1')
worksheet.range('D6').value = get_vol('IRK1K00499C1')
worksheet.range('D8').value = get_vol('IRK1K00600B1')
worksheet.range('D10').value = get_vol('IRTKLOTF0001')
worksheet.range('D12').value = get_vol('IRTKMOFD0001')
worksheet.range('D14').value = get_vol('IRTKKIAN0001')
worksheet.range('D16').value = get_vol('IRTKZARF0001')
print('phase2 done')
# Insert Last trade price for each ISIN in excel
worksheet.range('E2').value = get_Last_traded_price('IRK1K00500C1')
worksheet.range('E4').value = get_Last_traded_price('IRK1K00701C1')
worksheet.range('E6').value = get_Last_traded_price('IRK1K00499C1')
worksheet.range('E8').value = get_Last_traded_price('IRK1K00600B1')
worksheet.range('E10').value = get_Last_traded_price('IRTKLOTF0001')
worksheet.range('E12').value = get_Last_traded_price('IRTKMOFD0001')
worksheet.range('E14').value = get_Last_traded_price('IRTKKIAN0001')
worksheet.range('E16').value = get_Last_traded_price('IRTKZARF0001')
print('phase3 done')
# Insert yesterday price for each ISIN in excel
worksheet.range('G2').value = get_yesterday('IRK1K00500C1')
worksheet.range('G4').value = get_yesterday('IRK1K00701C1')
worksheet.range('G6').value = get_yesterday('IRK1K00499C1')
worksheet.range('G8').value = get_yesterday('IRK1K00600B1')
worksheet.range('G10').value = get_yesterday('IRTKLOTF0001')
worksheet.range('G12').value = get_yesterday('IRTKMOFD0001')
worksheet.range('G14').value = get_yesterday('IRTKKIAN0001')
worksheet.range('G16').value = get_yesterday('IRTKZARF0001')
print('phase4 done')
# Insert NAV price for each ISIN in excel
worksheet.range('E11').value = get_NAV('IRTKLOTF0001')
worksheet.range('E13').value = get_NAV('IRTKMOFD0001')
worksheet.range('E15').value = get_NAV('IRTKKIAN0001')
worksheet.range('E17').value = get_NAV('IRTKZARF0001')
print('phase5 done')
# Insert Time in excel
worksheet.range('D11').value = datetime.datetime.now()
worksheet.range('D13').value = datetime.datetime.now()
worksheet.range('D15').value = datetime.datetime.now()
worksheet.range('D17').value = datetime.datetime.now()
print('phase6 done')
# insert date in excel
worksheet.range('C11').value = datetime.datetime.today()
worksheet.range('C13').value = datetime.datetime.today()
worksheet.range('C15').value = datetime.datetime.today()
worksheet.range('C17').value = datetime.datetime.today()
print('phase7 done')
# Insert minimum traded price for each ISIN in excel
worksheet.range('H2').value = get_min_traded('IRK1K00500C1')
worksheet.range('H4').value = get_min_traded('IRK1K00701C1')
worksheet.range('H6').value = get_min_traded('IRK1K00499C1')
worksheet.range('H8').value = get_min_traded('IRK1K00600B1')
worksheet.range('H10').value = get_min_traded('IRTKLOTF0001')
worksheet.range('H12').value = get_min_traded('IRTKMOFD0001')
worksheet.range('H14').value = get_min_traded('IRTKKIAN0001')
worksheet.range('H16').value = get_min_traded('IRTKZARF0001')
print('phase8 done')
# Insert maximum traded price for each ISIN in excel
worksheet.range('I2').value = get_max_traded('IRK1K00500C1')
worksheet.range('I4').value = get_max_traded('IRK1K00701C1')
worksheet.range('I6').value = get_max_traded('IRK1K00499C1')
worksheet.range('I8').value = get_max_traded('IRK1K00600B1')
worksheet.range('I10').value = get_max_traded('IRTKLOTF0001')
worksheet.range('I12').value = get_max_traded('IRTKMOFD0001')
worksheet.range('I14').value = get_max_traded('IRTKKIAN0001')
worksheet.range('I16').value = get_max_traded('IRTKZARF0001')
print('phase9 done')
# Insert minimum price for each ISIN in excel
worksheet.range('J2').value = get_min_day('IRK1K00500C1')
worksheet.range('J4').value = get_min_day('IRK1K00701C1')
worksheet.range('J6').value = get_min_day('IRK1K00499C1')
worksheet.range('J8').value = get_min_day('IRK1K00600B1')
worksheet.range('J10').value = get_min_day('IRTKLOTF0001')
worksheet.range('J12').value = get_min_day('IRTKMOFD0001')
worksheet.range('J14').value = get_min_day('IRTKKIAN0001')
worksheet.range('J16').value = get_min_day('IRTKZARF0001')
print('phase10 done')
# Insert max price for each ISIN in excel
worksheet.range('K2').value = get_max_day('IRK1K00500C1')
worksheet.range('K4').value = get_max_day('IRK1K00701C1')
worksheet.range('K6').value = get_max_day('IRK1K00499C1')
worksheet.range('K8').value = get_max_day('IRK1K00600B1')
worksheet.range('K10').value = get_max_day('IRTKLOTF0001')
worksheet.range('K12').value = get_max_day('IRTKMOFD0001')
worksheet.range('K14').value = get_max_day('IRTKKIAN0001')
worksheet.range('K16').value = get_max_day('IRTKZARF0001')
print('phase11 done')
# Insert BID VOL for each ISIN in excel
worksheet.range('M2').value = get_bid_vol('IRK1K00500C1')
worksheet.range('M4').value = get_bid_vol('IRK1K00701C1')
worksheet.range('M6').value = get_bid_vol('IRK1K00499C1')
worksheet.range('M8').value = get_bid_vol('IRK1K00600B1')
worksheet.range('M10').value = get_bid_vol('IRTKLOTF0001')
worksheet.range('M12').value = get_bid_vol('IRTKMOFD0001')
worksheet.range('M14').value = get_bid_vol('IRTKKIAN0001')
worksheet.range('M16').value = get_bid_vol('IRTKZARF0001')
print('phase12 done')
# Insert BID price for each ISIN in excel
worksheet.range('N2').value = get_bid_price('IRK1K00500C1')
worksheet.range('N4').value = get_bid_price('IRK1K00701C1')
worksheet.range('N6').value = get_bid_price('IRK1K00499C1')
worksheet.range('N8').value = get_bid_price('IRK1K00600B1')
worksheet.range('N10').value = get_bid_price('IRTKLOTF0001')
worksheet.range('N12').value = get_bid_price('IRTKMOFD0001')
worksheet.range('N14').value = get_bid_price('IRTKKIAN0001')
worksheet.range('N16').value = get_bid_price('IRTKZARF0001')
print('phase13 done')
# Insert Ask price for each ISIN in excel
worksheet.range('P2').value = get_ask_price('IRK1K00500C1')
worksheet.range('P4').value = get_ask_price('IRK1K00701C1')
worksheet.range('P6').value = get_ask_price('IRK1K00499C1')
worksheet.range('P8').value = get_ask_price('IRK1K00600B1')
worksheet.range('P10').value = get_ask_price('IRTKLOTF0001')
worksheet.range('P12').value = get_ask_price('IRTKMOFD0001')
worksheet.range('P14').value = get_ask_price('IRTKKIAN0001')
worksheet.range('P16').value = get_ask_price('IRTKZARF0001')
print('phase14 done')
# Insert Ask vol for each ISIN in excel
worksheet.range('Q2').value = get_ask_vol('IRK1K00500C1')
worksheet.range('Q4').value = get_ask_vol('IRK1K00701C1')
worksheet.range('Q6').value = get_ask_vol('IRK1K00499C1')
worksheet.range('Q8').value = get_ask_vol('IRK1K00600B1')
worksheet.range('Q10').value = get_ask_vol('IRTKLOTF0001')
worksheet.range('Q12').value = get_ask_vol('IRTKMOFD0001')
worksheet.range('Q14').value = get_ask_vol('IRTKKIAN0001')
worksheet.range('Q16').value = get_ask_vol('IRTKZARF0001')
print('phase15 done')
worksheet.range('D22').value = shifter(name_tag)
worksheet.range('E22').value = shifter(price_tag)
print('phase16 done')
