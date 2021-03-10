import requests
import math
import pymysql
from urllib import parse

# sku_id_code_dic
conn = pymysql.connect()
cur = conn.cursor()
result = []
page = 1
headers = { 'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
            'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
            'cookie': 'Hm_lvt_11e03fb83d007f6132ac6d12a0f5eb78=1607623441; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6IjkyMzgiLCJOYW1lIjoi5byg5Y2O5rSLIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTY0MzA5NDM1MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.8WIlEJkaS9KDSvSCo7tVU9tfvdf6ZArVX3C5PfuC1e0; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1611558328,1611652912,1611728909,1611806152; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1611828984'}
url = 'https://erp.banmaerp.com/Product/Spu/SkuListData'
data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22IsValid%22%3A%7B%22Value%22%3Atrue%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A{0}%2C%22PageSize%22%3A500%7D%7D'.format(str(page))
r = requests.post(url=url, headers=headers, data=data)
for sku in r.json()['Results']:
    sku_id = sku['SKUID']
    spu_id = sku['SPUID']
    sku_code = sku['SkuCode']
    sku_price = sku['CostPrice']
    sku_weight = sku['Weight']
    result.append([sku_id, spu_id, sku_code, sku_price, sku_weight])
max_page = math.ceil(r.json()['TotalCount'] / 500)
page += 1
while page <= max_page:
    data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22IsValid%22%3A%7B%22Value%22%3Atrue%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A{0}%2C%22PageSize%22%3A500%7D%7D'.format(
        str(page))
    r = requests.post(url=url, headers=headers, data=data)
    for sku in r.json()['Results']:
        sku_id = sku['SKUID']
        spu_id = sku['SPUID']
        sku_code = sku['SkuCode']
        sku_price = sku['CostPrice']
        sku_weight = sku['Weight']
        result.append([sku_id, spu_id, sku_code, sku_price, sku_weight])
    page += 1
print(result)
cur.execute('truncate table sku_id_code_dic')
sql = '''insert into sku_id_code_dic (sku_id, spu_id, sku_code, sku_price, sku_weight) '''
cur.executemany(sql + '''values (%s,%s,%s,%s,%s)''', result)
conn.commit()
print('success')
cur.close()
conn.close()