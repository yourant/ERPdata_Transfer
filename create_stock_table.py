import datetime
import datetime as dt
import os
import numpy as np
import openpyxl as op
import pandas as pd
import pymysql
import requests
from sqlalchemy import create_engine


def read_table(path):
    wb = op.load_workbook(path)
    ws = wb.active
    df = pd.DataFrame(ws.values)
    df = pd.DataFrame(df.iloc[1:].values, columns=df.iloc[0, :])
    return df




"""
required documents:
1. inventory_export.csv
"""
# 设置时间
PATH = '/Users/edz/Documents'
start_days = '2020-08-01'
end_days = str(dt.datetime.now().date())
start_day = datetime.datetime.strptime(start_days, '%Y-%m-%d').date()
end_day = datetime.datetime.strptime(end_days, '%Y-%m-%d').date()
daytime = -1
now = dt.datetime.now()

if start_day <= end_day:
    daytime0 = end_day - start_day
    daytime = int(daytime0.days) + 1
else:
    print('起始日期大于结束日期')
    quit()

# # 设置headers
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
# 请求质检表数据
url = 'https://erp.banmaerp.com/Stock/Quality/QualityExportHandler'
data = 'filter=%7B%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name = PATH + '/待质检数据.xlsx'
with open(file_name, 'wb') as file:
    file.write(r.content)
data_dzj = read_table(file_name)
# data_dzj = pd.DataFrame(data_dzj.iloc[1:].values, columns=data_dzj.iloc[0, :])
data_dzj = data_dzj.rename(columns={'本地SKU': 'sku'})
dzj = data_dzj[['sku', '当前入库单SKU未质检数量']].groupby(['sku'], as_index=False).sum()
dzj = dzj.rename(columns={'当前入库单SKU未质检数量': '待质检数量'})
# 请求SKU配对关系表
url = 'https://erp.banmaerp.com/Product/Platform/ExportSkuMappingHandler'
data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name = PATH + '/SKU配对关系表.xlsx'
with open(file_name, 'wb') as file:
    file.write(r.content)
data_sku_pp = read_table(file_name)
os.remove(file_name)
# 请求订单数据
data_dd = None
url = 'https://erp.banmaerp.com/Order/Order/ExportOrderHandler'

# 导出订单列表的订单
# 15天来取
diff_ds = end_day - start_day
diff_days = int(diff_ds.days) + 1
data_dd_by_day_list = []
temp_date = start_day
if diff_days > 15:
    step = 15
    for single_date in (start_day + datetime.timedelta(n) for n in range(15, diff_days, step)):
        print(temp_date,"toooooo",single_date)
        data = 'filter=%7B%22OriginalOrderTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.0000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.9999%22%2C%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%7D&details%5B0%5D%5BFieldID%5D=37&details%5B0%5D%5BSort%5D=1&details%5B0%5D%5BFieldExportName%5D=&details%5B1%5D%5BFieldID%5D=40&details%5B1%5D%5BSort%5D=2&details%5B1%5D%5BFieldExportName%5D=&details%5B2%5D%5BFieldID%5D=43&details%5B2%5D%5BSort%5D=3&details%5B2%5D%5BFieldExportName%5D=&details%5B3%5D%5BFieldID%5D=70&details%5B3%5D%5BSort%5D=4&details%5B3%5D%5BFieldExportName%5D=&details%5B4%5D%5BFieldID%5D=221&details%5B4%5D%5BSort%5D=5&details%5B4%5D%5BFieldExportName%5D=&details%5B5%5D%5BFieldID%5D=253&details%5B5%5D%5BSort%5D=6&details%5B5%5D%5BFieldExportName%5D=&details%5B6%5D%5BFieldID%5D=66&details%5B6%5D%5BSort%5D=7&details%5B6%5D%5BFieldExportName%5D=&details%5B7%5D%5BFieldID%5D=67&details%5B7%5D%5BSort%5D=8&details%5B7%5D%5BFieldExportName%5D=&details%5B8%5D%5BFieldID%5D=68&details%5B8%5D%5BSort%5D=9&details%5B8%5D%5BFieldExportName%5D=&type=1'.format(
            temp_date, single_date)
        r = requests.post(url=url, headers=headers, data=data)
        file_name = '/Users/edz/Documents/{0}到{1}订单数据.xlsx'.format(temp_date, single_date)
        temp_date = single_date + dt.timedelta(days=1)
        with open(file_name, 'wb') as file:
            file.write(r.content)
        data_dd_by_day_list.append(file_name)
        if data_dd is None:
            try:
                data_dd = read_table(file_name)
            except Exception as e:
                print(e)
                continue
        else:
            try:
                data_dd_cur = read_table(file_name)
                data_dd = pd.concat([data_dd, data_dd_cur], ignore_index=True)
            except Exception as e:
                print(e)
                continue
data = 'filter=%7B%22OriginalOrderTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.0000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.9999%22%2C%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%7D&details%5B0%5D%5BFieldID%5D=37&details%5B0%5D%5BSort%5D=1&details%5B0%5D%5BFieldExportName%5D=&details%5B1%5D%5BFieldID%5D=40&details%5B1%5D%5BSort%5D=2&details%5B1%5D%5BFieldExportName%5D=&details%5B2%5D%5BFieldID%5D=43&details%5B2%5D%5BSort%5D=3&details%5B2%5D%5BFieldExportName%5D=&details%5B3%5D%5BFieldID%5D=70&details%5B3%5D%5BSort%5D=4&details%5B3%5D%5BFieldExportName%5D=&details%5B4%5D%5BFieldID%5D=221&details%5B4%5D%5BSort%5D=5&details%5B4%5D%5BFieldExportName%5D=&details%5B5%5D%5BFieldID%5D=253&details%5B5%5D%5BSort%5D=6&details%5B5%5D%5BFieldExportName%5D=&details%5B6%5D%5BFieldID%5D=66&details%5B6%5D%5BSort%5D=7&details%5B6%5D%5BFieldExportName%5D=&details%5B7%5D%5BFieldID%5D=67&details%5B7%5D%5BSort%5D=8&details%5B7%5D%5BFieldExportName%5D=&details%5B8%5D%5BFieldID%5D=68&details%5B8%5D%5BSort%5D=9&details%5B8%5D%5BFieldExportName%5D=&type=1'.format(
        temp_date, end_day)
print(temp_date, end_day)
r = requests.post(url=url, headers=headers, data=data)
file_name = PATH + '/{0}到{1}订单数据.xlsx'.format(temp_date, end_day)
with open(file_name, 'wb') as file:
        file.write(r.content)
data_dd_by_day_list.append(file_name)
if data_dd is None:
    try:
        data_dd = read_table(file_name)
    except Exception as e:
        print(e)
elif temp_date != end_day:
    try:
        data_dd_cur = read_table(file_name)
        data_dd = pd.concat([data_dd, data_dd_cur], ignore_index=True)
    except Exception as e:
        print(e)

file_name_dd = PATH + '/订单数据.xlsx'
data_dd.to_excel(file_name_dd)
for dir_file in data_dd_by_day_list:
    os.remove(dir_file)


# 请求采购单数据
url = 'https://erp.banmaerp.com/Purchase/Sheet/ExportPurchaseHandler'
data_cgd = None
begin_day = '2020-08-01'
begin_day = datetime.datetime.strptime(begin_day, '%Y-%m-%d').date()
diff_days = end_day - begin_day
diff_days = int(diff_days.days) + 1
months = diff_days / 30

data_cgd_by_day_list = []
temp_date = begin_day
if months > 0:
    step = 30
    for single_date in (begin_day + datetime.timedelta(n) for n in range(30, diff_days, step)):
        data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A%22-1%22%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%2C%22CreateTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.998%22%7D%7D'.format(
            temp_date, single_date)
        r = requests.post(url=url, headers=headers, data=data)
        file_name = '/Users/edz/Documents/{0}到{1}采购单数据.xlsx'.format(temp_date, single_date)
        temp_date = single_date + datetime.timedelta(days=1)
        with open(file_name, 'wb') as file:
            file.write(r.content)
        data_cgd_by_day_list.append(file_name)
        if data_cgd is None:
            data_cgd = read_table(file_name)
        else:
            try:
                data_cgd_cur = read_table(file_name)
                data_cgd = pd.concat([data_cgd, data_cgd_cur], ignore_index=True)

            except Exception as e:
                continue
data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A%22-1%22%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%2C%22CreateTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.998%22%7D%7D'.format(
    temp_date, end_day)
r = requests.post(url=url, headers=headers, data=data)
file_name = PATH + '/{0}到{1}采购单数据.xlsx'.format(temp_date, end_day)
with open(file_name, 'wb') as file:
    file.write(r.content)
data_cgd_by_day_list.append(file_name)
if data_cgd is None:
    data_cgd = read_table(file_name)
elif temp_date != end_day:
    try:
        data_cgd_cur = read_table(file_name)
        data_cgd = pd.concat([data_cgd, data_cgd_cur], ignore_index=True)
    except Exception as e:
        pass



# 删除多余订单数据文件
for dir_file in data_cgd_by_day_list:
    os.remove(dir_file)
data_cgd = data_cgd[data_cgd['仓库'] == '自建仓-坑头']


# 请求库存数据
url = 'https://erp.banmaerp.com/Stock/SelfInventory/ExportHandler'
data = 'filter=%7B%22Quantity%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A10000%2C%22PageNumber%22%3A1%7D%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name_kc = '/Users/edz/Documents/库存数据.xlsx'
with open(file_name_kc, 'wb') as file:
    file.write(r.content)

# 请求在线商品数据
url = 'https://erp.banmaerp.com/Shopify/Product/ExportHandler'
data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A-1%7D%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name_cp = '/Users/edz/Documents/在线商品数据.xlsx'
with open(file_name_cp, 'wb') as file:
    file.write(r.content)

# 获取在线商品数据并删除标题
data_cp = read_table(file_name_cp)
if "Shopify产品" in data_cp.columns.tolist():
    data_cp = pd.DataFrame(data_cp.iloc[1:].values, columns=data_cp.iloc[0, :])

data_kc = read_table(file_name_kc)
if "库存清单数据" in data_kc.columns.tolist():
    data_kc = pd.DataFrame(data_kc.iloc[1:].values, columns=data_kc.iloc[0, :])
# 只取 坑头仓库+虹猫蓝兔仓库 库存数据
data_kc = data_kc[(data_kc['仓库'] == '坑头') | (data_kc['仓库'] == '虹猫蓝兔动漫有限公司')]
# 计算出空闲库存=合格总量-合格锁定
data_kc['空闲库存'] = data_kc['合格总量'] - data_kc['合格锁定量']
# 透视求和得到表kc
kc = data_kc[['空闲库存', '本地sku']].groupby(['本地sku'], as_index=False).sum()
kc = kc.rename(columns={'本地sku': 'sku'})

# 导出采购单列表，只取采购中 + 待审核，求得在途库存，透视求和
data_cgd = data_cgd[['采购单号', '状态', '本地SKU', '物品数量', '到货物品数量']]
data_cgd['在途库存'] = data_cgd['物品数量'].astype(float) - data_cgd['到货物品数量'].astype(float)
data_cgd = data_cgd[(data_cgd['状态'] == '采购中') | (data_cgd['状态'] == '待审核')]
cg = data_cgd[['本地SKU', '在途库存']].groupby(['本地SKU'], as_index=False).sum()
cg = cg.rename(columns={'本地SKU': 'sku'})
# 导出订单列表中“缺货中”数据，得到缺货数量
data_dd = read_table(file_name_dd)
missing_quantity = []
for i in range(data_dd.shape[0]):
    if data_dd.loc[i, '缺货数量'] == '--':
        missing_quantity.append(1)
    else:
        missing_quantity.append(int(data_dd.loc[i, '缺货数量']))
data_dd['缺货数量'] = np.array(missing_quantity)
# data_dd = read_table("/Users/edz/Documents/订单导出-20210224214748.xlsx")
# 导出时，按产品导出，勾选“订单状态“，“订单号”，“付款时间”，“缺货数量”，“匹配SKU”，“平台SKU”
dd = data_dd[["订单状态", "订单号", "缺货数量", "匹配SKU", "平台SKU"]]

# 数据透视，“匹配SKU”为行，“缺货数量”为值求和，得出缺货数量
d = dd[['匹配SKU', '缺货数量']].groupby(["匹配SKU"], as_index=False).sum()
d = d.rename(columns={'匹配SKU': 'sku'})

# 空闲库存表里的SKU，在途库存SKU，待质检SKU，缺货列表SKU都放到同一张表格上，删除重复项，得到所有唯一SKU，"表a"
frames = [kc[['sku']], cg[['sku']], d[['sku']], dzj[['sku']]]
a = pd.concat(frames)
a = a.drop_duplicates(subset='sku')

# 将空闲库存数据，在途库存数据，待质检数据，缺货数据都v_lookup到"表a“上
a = a.merge(kc, on='sku', how='left')
a = a.merge(cg, on='sku', how='left')
a = a.merge(d, on='sku', how='left')
a = a.merge(dzj, on='sku', how='left')
# 求得shopify库存数量= 空闲库存+在途库存+待质检-缺货
a['空闲库存'] = a['空闲库存'].fillna(0)
a['在途库存'] = a['在途库存'].fillna(0)
a['待质检数量'] = a['待质检数量'].fillna(0)
a['缺货数量'] = a['缺货数量'].fillna(0)
a['缺货数量'] = np.where(a['缺货数量'].astype('float') < 0, 0, a['缺货数量'])
a['shopify库存数量'] = a['空闲库存'] + a['在途库存'] + a['待质检数量'] - a['缺货数量'].astype('float')
# 利用SKU配对关系表得出平台sku的shopify库存数量，得到“表b”
b = a.merge(data_sku_pp, left_on='sku', right_on='本地SKU', how='right')
b = b[['平台SKU', 'shopify库存数量']]
b['shopify库存数量'] = b['shopify库存数量'].fillna(0)
print(b.columns)
print(b.head())
os.remove(PATH + '/在线商品数据.xlsx')
os.remove(PATH + '/库存数据.xlsx')
os.remove(PATH + '/待质检数据.xlsx')
os.remove(PATH + '/订单数据.xlsx')


conn_test = pymysql.connect(host='rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com',
                            port=3306, user='leiming',
                            passwd='vg4wHTnJlbWK8SY',
                            db="cider",
                            charset='utf8',
                            cursorclass=pymysql.cursors.DictCursor)
cur_test = conn_test.cursor()
engine = create_engine(
    'mysql+pymysql://leiming:vg4wHTnJlbWK8SY@rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com:3306/cider')

# INSERT
data_stock = b
data_stock = data_stock.dropna(subset=['平台SKU', 'shopify库存数量'])
data_stock = data_stock.reset_index()
print(data_stock.head())
for i in range(data_stock.shape[0]):
    with conn_test.cursor() as cursor:
        sql = '''INSERT INTO shopify_stock (sku_code, stock, add_time) VALUES ("{0}", {1}, NOW())'''.format(
            data_stock.loc[i, '平台SKU'], data_stock.loc[i, 'shopify库存数量'])
        engine.execute(sql)
cursor.close()
