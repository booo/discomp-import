from lxml import html
from lxml.cssselect import CSSSelector
import urllib.request
import xlsxwriter
import sys

sel_name = CSSSelector('#sti_detail_head > h1:nth-child(1)')
# doesn't work when there is only one price and no recycling fee
# sel_price = CSSSelector('tr.prc:nth-child(2) > td:nth-child(2)')
sel_price = CSSSelector('.price_without_vat')
sel_code = CSSSelector('tr.code > td:nth-child(2)')
sel_partno = CSSSelector('tr.partno > td:nth-child(2)')
sel_datasheet = CSSSelector('a.downloadlink:nth-child(2)')

# Create an new Excel file and add a worksheet.
workbook = xlsxwriter.Workbook('products.xlsx')
worksheet = workbook.add_worksheet()

# Widen the first column to make the text clearer.
worksheet.set_column('A:A', 20)

# Add a bold format to use to highlight cells.
bold = workbook.add_format({'bold': True})

# Internal Reference	Name	Public Price	Customer Taxes	Vendor Taxes

worksheet.write('A1', 'Internal Reference')
worksheet.write('B1', 'Name')
worksheet.write('C1', 'Public Price')
worksheet.write('D1', 'External ID')
worksheet.write('E1', 'Vendor Taxes')
worksheet.write('F1', 'Description')
worksheet.write('G1', 'Cost')
#worksheet.write('D1', 'Customer Taxes')

for i, url in enumerate(sys.stdin):
    print(f'Scraping {url}')
    row_index = i + 1
    with urllib.request.urlopen(url) as f:
        text = f.read().decode('utf-8')
        tree = html.fromstring(text)

        name = sel_name(tree)[0].text
        price = sel_price(tree)[0].text.replace("EUR", "")
        code = sel_code(tree)[0].text
        partno = sel_partno(tree)[0].text
        datasheet = sel_datasheet(tree)[0].attrib['href'] if sel_datasheet(tree) else ""

        worksheet.write(row_index, 0, partno)
        worksheet.write(row_index, 1, name)
        worksheet.write(row_index, 2, price)
        worksheet.write(row_index, 3, code)
        # TODO change taxes for other vendors
        worksheet.write(row_index, 4, 'Innergem. Erwerb 19%USt/19%VSt')
        worksheet.write(row_index, 5, datasheet)
        worksheet.write(row_index, 6, price)


        #print('External ID; Name; Product Type; Internal Reference;	Sales Price; Cost')
        #print(f'{partno}; {name}; Consumable; {code}; {price}; {price}')

workbook.close()
