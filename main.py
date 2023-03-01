import pandas as pd
import openpyxl

import json
import os
import os.path
import re
from bs4 import BeautifulSoup
import requests

from dotenv import load_dotenv
load_dotenv()

TIKI = os.environ.get("TIKI")
# remove all file in folder excel
parent = os.getcwd()
path = os.path.join(parent, 'excel\\')
for file_name in os.listdir(path):
    file = path + file_name
    if os.path.isfile(file):
        os.remove(file)
#

headers = {
    "user-agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/110.0.0.0 Safari/537.36"}
res = requests.get(TIKI, headers=headers)  # must have headers
soup = BeautifulSoup(res.text, 'html.parser')  # get main source

# listTitle[]
# ---->title, link, more[name, link]
listTitle = []  # contain title
for ahref in soup.find_all('div', class_="styles__StyledCategory-sc-17y817k-1 iBByno"):
    item = {"title": ahref.find('a').get_text()}
    item.update({"link": ahref.find('a').attrs['href']})
    listMore = []
    for bhref in ahref.find('p').find_all('a'):
        listMore.append({"name": bhref.get_text(),
                        "link": bhref.attrs['href']})
    item.update({"more": listMore})
    listTitle.append(item)


id_product = []

# infor one
# https://tiki.vn/api/v2/products/{id}


def _writeData(lk, f, ite):
    id_product = []
    page = "https://tiki.vn/api/personalish/v1/blocks/listings?limit=40&include=advertisement&aggregations=2&trackity_id=fb703a37-2f3d-2956-96ac-bd90536bb792&category={}&page=".format(
        re.sub(r'\D', '', ite['link']))
    print()
    print(page)
    num = 1
    dict = {'Id': [], 'Name': [], 'Price': [], 'Image': []}
    while True:
        print('Process "{}" page {}'.format(ite['name'], num))
        result = requests.get(page+str(num), headers=headers).text
        if (json.loads(result).get('data') == None or json.loads(result).get('data') == []):
            break
        for item in json.loads(result).get('data'):
            if item.get('id') not in id_product:
                dict['Id'].append(item.get('id'))
                dict['Name'].append(item.get('name'))
                dict['Price'].append(item.get('price'))
                dict['Image'].append(item.get('thumbnail_url'))

                id_product.append(item.get('id'))
                f.write(str(item.get('id'))+'\n')
                f.write(item.get('name')+'\n')
                f.write(str(item.get('price'))+'\n')
                f.write(item.get("thumbnail_url")+'\n\n')

        num += 1
    df = pd.DataFrame(dict)
    if os.path.isfile(lk):
        writer = pd.ExcelWriter(lk, engine='openpyxl', mode='a')
    else:
        writer = pd.ExcelWriter(lk, engine='openpyxl')
    try:
        df.to_excel(writer, sheet_name=ite['name'], encoding='utf8')
    except:
        pass
    writer.close()


for title in listTitle:
    print('\nProcess "{}"'.format(title['title']))
    print('-'*30)
    file = 'data/' + title['title'] + '.txt'
    lk = 'excel/' + title['title']+".xlsx"
    f = open(file, 'w', encoding="utf-8")
    for item in title['more']:
        f.write('[NEW] '+item['name']+'\n\n')
        _writeData(lk, f, item)
    f.close()
    print('DONE!!!!!!!!!!!!!!!!')
    print('-'*30)
