import requests
from bs4 import BeautifulSoup
import re
import xlsxwriter



def get_num_of_pages():
    source = requests.get('https://www.josephprince.com/sermons?page=3').text
    soup = BeautifulSoup(source, 'lxml')
    button = soup.find_all('li')
    return button[21].text


sermons = []
dates = []
links = []
descriptions = []

last_page = int(get_num_of_pages())

def get_description(link):
    links.append(link)
    source = requests.get(link).text
    soup = BeautifulSoup(source, 'lxml')
    script = soup.find_all('script')
    text = script[8].text
    descrip = re.findall(r'''"description": "[^"]+"''', text)[0][16:-1]
    return descrip

for page in range(1 , last_page):
    source = requests.get(f'https://www.josephprince.com/sermons?page={page}').text
    soup = BeautifulSoup(source, 'lxml')
    button = soup.find_all('a',href=True)
    try:
        for item in button:
            if item.h2 and item.div:
                link = item['href']
                sermons.append(item.h2.text)
                dates.append(item.div.text.strip())
                print(item.h2.text)
                print(item.div.text.strip())
                print(link)
                try:
                    descriptions.append(get_description(link))
                except:
                    print('Cannot get description')
                    descriptions.append('NIL')
    except:
        pass


workbook = xlsxwriter.Workbook('Joseph Prince Sermons v3.xlsx')
worksheet = workbook.add_worksheet()


row = 0
column = 0

worksheet.write(row, column, 'Sermons')
worksheet.write(row, column + 1, 'Date')
worksheet.write(row, column + 2, 'Link')
worksheet.write(row, column + 3, 'Sermon description')

row += 1

# iterating through content list
for i in range(len(links)):
    # write operation perform
    worksheet.write(row, column, sermons[i])
    worksheet.write(row, column + 1, dates[i])
    worksheet.write(row, column + 2, links[i])
    worksheet.write(row, column + 3, descriptions[i])

    row += 1

workbook.close()