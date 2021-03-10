import requests
import json
import pandas
import numpy as np
import pandas as pd
import pymysql
from sqlalchemy import create_engine

# 连接 plutus 数据库
engine = create_engine(
    'mysql+pymysql://leiming:pQx2WhYhgJEtU5r@rm-2ze314ym42f9iq2xflo.mysql.rds.aliyuncs.com:3306/plutus')
conn = pymysql.connect(host='rm-2ze314ym42f9iq2xflo.mysql.rds.aliyuncs.com',
                       port=3306, user='leiming',
                       passwd='pQx2WhYhgJEtU5r',
                       db="plutus",
                       charset='utf8',
                       cursorclass=pymysql.cursors.DictCursor)

# 读取porduct表
product = pd.read_sql_table('product', engine)
print(product.columns)
# 读取斑马erp所有匹配的spu的图片数据
# 遍历所有id
spu = product[['zebra_spu_id']]
original_url = 'https://erp.banmaerp.com/Product/Spu/Edit/'

headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}

sku_code = []
sku_image = []
sku_images = []
spu_image = []
spu_images = []
wrong_box = []
for i in range(1, spu.shape[0]):
    url = original_url + str(int(spu.loc[i, 'zebra_spu_id']))
    try:
        r = requests.get(url, headers=headers)
        product = r.text[r.text.find('var product = ') + 14:r.text.find(('var isAdd = '))][:-11]
        images = ''
        spu_img = []
        for src in json.loads(product)['Images']:
            images += src['Src'] + ','
            spu_img.append(src['Src'])
        images = images[:len(images) - 1]
        main_image = set()
        for sku in json.loads(product)['Skus']:
            main_image.add(sku['Image'])
        spu_images.append(images)
        spu_image.append(json.loads(product)["Spu"]['Image'])
        for sku in json.loads(product)['Skus']:
            first_idx = spu_img.index(sku['Image'])
            imgs = ""
            imgs += sku['Image']
            if first_idx + 1 < len(spu_img):
                while spu_img[first_idx + 1] not in main_image:
                    imgs = imgs + ',' + spu_img[first_idx + 1]
                    first_idx += 1
                    if first_idx == len(spu_img) - 1:
                        break
            sku_images.append(imgs)
            sku_code.append(sku['Code'])
            sku_image.append(sku['Image'])
    except Exception as e:
        print(e)
        wrong_box.append(str(int(spu.loc[i, 'zebra_spu_id'])))

print(len(wrong_box))
print(wrong_box)
print(len(sku_image))
print(len(sku_images))
print(len(sku_code))
print(len(spu_image))
print(len(spu_images))

for i in range(len(spu_image)):
    sql = '''UPDATE product SET primary_image="{0}", product_images = "{1}" WHERE zebra_spu_id = {2}'''.format(
        spu_image[i], spu_images[i], int(str(int(spu.loc[i+1, 'zebra_spu_id']))))
    # print('spuid:' + str(int(spu.loc[i+1, 'zebra_spu_id'])) + 'spu_image: ' + spu_image[i],
    #       'spu_images: ' + spu_images[i])
    engine.execute(sql)

print('----------------------------------------------------')
for i in range(len(sku_code)):
    sql = '''UPDATE sku_main SET primary_image="{0}", secondary_images = "{1}" WHERE sku_code = "{2}"'''.format(
        sku_image[i], sku_images[i], sku_code[i])
    # print('sku_code:' + sku_code[i] + "  image: " + sku_image[i] + '   images_set: ' + str(sku_images[i]))
    engine.execute(sql)