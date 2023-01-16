import urllib.request
import xlsxwriter
import json
import os
from bs4 import BeautifulSoup

data = []

# get url of item page and fint the data.
def item_crawler(url):
    source_code = urllib.request.urlopen(url)
    plain_text = source_code.read()
    soup = BeautifulSoup(plain_text , "html.parser")

    # name
    link = soup.find('h1', {'data-test': 'product-overview-name'})
    name = link.text
    print("name: " + name)

    # pieces
    link = soup.find('div', {'data-test': 'pieces-value'})
    pieces_count = link.string
    print("pieces: " + pieces_count)

    # price
    link = soup.find('span', {'data-test': 'product-price'})
    price = link.text
    price = price.split('$')[1]
    print("price: " + price)

    # minifigures
    link = soup.find('div', {'data-test': 'minifigures-value'})
    if (link):
        minifigures = link.text
    else:
        minifigures = "0"
    print("minifigures: " + minifigures)

    # dimensions
    link = soup.find('div', {'data-test': 'dimensions-value'})
        # if there is dimensions info
    if (link):
        str = link.text
        mk1 = str.find('(') + 1
        mk2 = str.find(')')
        height = str[mk1 : mk2]
        str = str[mk2+1:]
        mk1 = str.find('(') + 1
        mk2 = str.find(')')
        width = str[mk1: mk2]
        str = str[mk2 + 1:]
        mk1 = str.find('(') + 1
        mk2 = str.find(')')
        depth = str[mk1: mk2]
        print("dimensions-value: H: " + height + " W: " + width + " D: " + depth)

    else:
        height = "0cm"
        width = "0cm"
        depth = "0cm"
        print("dimensions-value: no info")


    ratio = float(pieces_count) / float(price)
    ratio = round(ratio, 2)
    print("preces/price = " ,ratio)

    # add to data
    data.append({'name': name, 'pieces count': pieces_count,
                 'price': price, 'height': height, 'width': width,
                 'depth': depth, 'minifigures': minifigures,
                 'pieces per dollar': ratio, 'url': url})

# getting max pages to run on.
# currently run on: Sets by theme -> 'creator-expert'
def trade_crawler(max_pages):
    page = 1
    while page <= max_pages:
        base_url = "https://www.lego.com"
        url = "https://www.lego.com/en-us/themes/creator-expert?page=" + str(page)
        source_code = urllib.request.urlopen(url)
        plain_text = source_code.read()
        soup = BeautifulSoup(plain_text, "html.parser")
        for link in soup.findAll('a', {'data-test': 'product-leaf-title-link'}):
            href = link.get('href')
            href = base_url + href
            print("\n" + href)
            item_crawler(href)

        page += 1

# write to json, create new file or overtie the old one
def json_write(data):
    file = open("./data.json", "w+")
    file.seek(0)
    file.write(json.dumps(data))

# write to excel
def excel_write(data):
    wb = xlsxwriter.Workbook("data.xlsx")
    sheet = wb.add_worksheet()
    row = 1
    bold = wb.add_format({'bold': True})

    #title row
    sheet.write(0, 0, 'Name', bold)
    sheet.write(0, 1, 'Price', bold)
    sheet.write(0, 2, 'Pieces', bold)
    sheet.write(0, 3, 'Minifigures', bold)
    sheet.write(0, 4, 'Height', bold)
    sheet.write(0, 5, 'Width', bold)
    sheet.write(0, 6, 'Depth', bold)
    sheet.write(0, 7, 'Pieces per $', bold)
    sheet.write(0, 8, 'url', bold)

    # write data
    for item in data:
        sheet.write(row, 0, item['name'])
        sheet.write(row, 1, item['price'])
        sheet.write(row, 2, item['pieces count'])
        sheet.write(row, 3, item['minifigures'])
        sheet.write(row, 4, item['height'])
        sheet.write(row, 5, item['width'])
        sheet.write(row, 6, item['depth'])
        sheet.write(row, 7, item['pieces per dollar'])
        sheet.write(row, 8, item['url'])
        row += 1

    wb.close()

# url for test item_crawler:
# url_flower_item = "https://www.lego.com//en-us/product/flower-bouquet-10280"
# url_bookshop_item = "https://www.lego.com/en-us/product/bookshop-10270"

# run the trade_crawler - max 2 page
trade_crawler(2)

# save the data
json_write(data)
excel_write(data)
