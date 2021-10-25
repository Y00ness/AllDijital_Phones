from bs4 import BeautifulSoup
import requests, openpyxl, os
from openpyxl.styles import Font, Color, Alignment

#   make an excel file
wb = openpyxl.Workbook()
ws = wb.active
ws.title = 'Phones and Prices'
ws.sheet_view.rightToLeft = True

# adjust header
ws.append(['__phone__', '__price__'])
for cell in list(ws[1]):
    cell.font = Font(bold=True, color='5F05AA')
ws.row_dimensions[1].height = 20

# urls for scrap
Apple_url = 'https://www.alldigitall.ir/category/206?page=1&order=sort_order,desc%7Cdate_added,desc&perpage=50'
Samsung_url = 'https://www.alldigitall.ir/category/430?page=1&order=sort_order,desc%7Cdate_added,desc&perpage=50'
NOKIA_url = 'https://www.alldigitall.ir/category/344?page=1&order=sort_order,desc%7Cdate_added,desc&perpage=50'
Huaewi_url = 'https://www.alldigitall.ir/category/247?page=1&order=sort_order,desc%7Cdate_added,desc&perpage=50'
Xiaomi_url = 'https://www.alldigitall.ir/category/1240?page=1&order=sort_order,desc%7Cdate_added,desc&perpage=50'
brands_url = [Apple_url, Samsung_url, NOKIA_url, Huaewi_url, Xiaomi_url]

def Brand():
    while True:
        brand_num = input('insert Num: 0-Apple  1-Samsung  2-NOKIA  3-Huawei  4-Xiaomi')
        if brand_num not in str([0, 1, 2, 3, 4]):
            print('Wrong Num, Try Again... :)')
        else:
            return int(brand_num)
# response the selected url
res = requests.get(brands_url[Brand()])
soup = BeautifulSoup(res.text, 'lxml')

# find phones block
phones_main_block = soup.find('div', class_='q-card__section q-card__section--vert q-py-none q-px-none row items-end rtl')
all_phones = phones_main_block.find_all('div', class_= 'VProduct text-center col q-pa-sm Square q-card q-card--bordered q-card--flat no-shadow')

# loop for names and prices
for phone in all_phones:
    name_parent = phone.find('div', class_= 'q-py-sm q-px-none full-width ProductName q-card__section q-card__section--vert')
    name = name_parent.find('div', class_='text-right text-subtitle2 text-black ellipsis-2-lines PName')
    name = name.text.strip()
    price = phone.a.find('div', class_='q-pb-none text-left price')

    # a code for finding content to find coming soon phones
    try:
        price_content = price['content']
    except:
        price_content = 1
    red_price = phone.a.find('div', class_= 'newprice text-red text-left')

    # show all phones we find and add them to excel file
    if int(price_content) == 0:
        print(name, '\nComing Soon\n--------------------')
        ws.append([name, 'Coming Soon'])
    elif price:
        print(name, '\n', price.text.strip(' تومان'), '\n--------------------')
        ws.append([name, price.text.strip()])
    elif red_price:
        print(name, '\n', red_price.text.strip(' تومان'), '\n--------------------')
        ws.append([name, red_price.text.strip()])
    else:
        print(name, '\nIs Not Available\n--------------------')
        ws.append([name, 'Not Available'])

# adjust the worksheet
ws.column_dimensions['A'].width = 45
ws.column_dimensions['B'].width = 30
cell_alignment = Alignment(horizontal='center', vertical='center')
cell_alignment.readingOrder = 2 #RTL
i = 2
for row in list(ws.rows):
    ws.row_dimensions[i].height = 25
    i += 1
    for cell in row:
        cell.alignment = cell_alignment
        if cell.value == 'Coming Soon':
            cell.font = Font(color='0BCF61')
        elif cell.value == 'Not Available':
            cell.font = Font(color='C91304')

# save and open excel file
wb.save('Alldigital_Phones.xlsx') 
os.system('start excel.exe Alldigital_Phones.xlsx')

i = input('\nHave a good day... :)\n\ndeveloped by ==Unes==')