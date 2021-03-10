import requests
import openpyxl as op
import pandas as pd
import os
import pymysql

# # 连接数据库
# # conn = pymysql.connect(host='rm-2ze314ym42f9iq2xflo.mysql.rds.aliyuncs.com',
# #                        port=3306, user='leiming',
# #                        passwd='pQx2WhYhgJEtU5r',
# #                        db="plutus",
# #                        charset='utf8')
#
# # 连接数据库(测试)
# conn = pymysql.connect(host='rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com',
#                        port=3306, user='leiming',
#                        passwd='vg4wHTnJlbWK8SY',
#                        db="plutus",
#                        charset='utf8')
# cur = conn.cursor()
#
#
# def read_table(path):
#     wb = op.load_workbook(path)
#     ws = wb.active
#     df = pd.DataFrame(ws.values)
#     df = pd.DataFrame(df.iloc[1:].values, columns=df.iloc[0, :])
#     return df
#
#
# url = 'https://erp.banmaerp.com/Product/Platform/ExportSkuMappingHandler'
# data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%7D'
# headers = {
#     'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
#     'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
#     'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
#
# r = requests.post(url=url, headers=headers, data=data)
# file_name = os.getcwd() + '/SKU配对关系表.xlsx'
# with open(file_name, 'wb') as file:
#     file.write(r.content)
# data_sku_pp = read_table(file_name)
# os.remove(file_name)
# # cur.execute('truncate table shopify_2_zebra')
# db_table = 'shopify_2_zebra'
# title = 'zebra_erp_code,shopify_erp_code'
# sql = 'insert into {} ({}) '.format(db_table, title)
# total_data_list = []
# for i in range(data_sku_pp.shape[0]):
#     total_data_list.append(tuple([data_sku_pp.loc[i, '本地SKU'], data_sku_pp.loc[i, '平台SKU']]))
# cur.executemany(sql + '''values (%s,%s)''', total_data_list)
# conn.commit()
# conn.close()

conn_test = pymysql.connect(host='rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com',
                       port=3306, user='leiming',
                       passwd='vg4wHTnJlbWK8SY',
                       db="cider",
                       charset='utf8')
cur_test = conn_test.cursor()
# 建表
# sql = """
# CREATE TABLE warehouse_location_info (
#   id int(10) unsigned NOT NULL AUTO_INCREMENT,
#   location_code varchar(50) NOT NULL,
#   location_id bigint DEFAULT NULL,
#   add_time datetime NOT NULL DEFAULT CURRENT_TIMESTAMP,
#   PRIMARY KEY (id)
# ) ;"""
# cur_test.execute(sql)
# conn_test.commit()
# 刷库位管理进数据库

headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
stock_url = 'https://erp.banmaerp.com/Stock/SelfWarehouse/LocationListData'
start = 1
while True:
    s = 'filter=%7B%22Pager%22%3A%7B%22PageNumber%22%3A{0}%2C%22PageSize%22%3A100%7D%7D&pageNumber={0}&pageSize=100'.format(
        start)
    r = requests.post(url=stock_url, headers=headers, data=s)
    if len(r.json()['Data']['Results']) != 0:
        for i in range(len(r.json()['Data']['Results'])):
            with conn_test.cursor() as cursor:
                sql = '''SELECT * FROM warehouse_location_info WHERE location_code = "{0}" and location_id = {1}'''.format(
                    r.json()['Data']['Results'][i]['Code'], int(r.json()['Data']['Results'][i]['ID']))
                cursor.execute(sql)
                res = cursor.fetchone()
                if res is None:
                    sql = '''INSERT INTO  warehouse_location_info (location_code, location_id) VALUES ("{0}",{1})'''.format(
                        r.json()['Data']['Results'][i]['Code'], int(r.json()['Data']['Results'][i]['ID']))
                    cursor.execute(sql)
                    conn_test.commit()
    else:
        break
    start += 1


