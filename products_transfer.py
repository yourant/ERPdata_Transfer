import numpy as np
import openpyxl as op
import pandas as pd
import pymysql
from sqlalchemy import create_engine
import requests
import datetime as dt
import os
import xlrd
def read_table(path):
    wb = op.load_workbook(path)
    ws = wb.active
    df = pd.DataFrame(ws.values)
    df = pd.DataFrame(df.iloc[1:].values, columns=df.iloc[0, :])
    return df

def is_contain_chinese(check_str):
    """
    判断字符串是否包含中文
    """
    for ch in check_str:
        if ord(ch) > 255:
            return True
    return False


def is_chinese(l):
    """
    删除list里含有中文的字符串
    :param l: 待检测的字符串list
    :return: 删去中文字符串后的list
    """
    res = []
    for i in l:
        try:
            if not is_contain_chinese(i):
                res.append(i)
        except:
            continue
    return res


def trim(s):
    """
    删除字符串首位空格
    """
    if s == '':
        return s
    elif s[0] == ' ':
        return trim(s[1:])
    elif s[-1] == ' ':
        return trim(s[:-1])
    else:
        return s


# 连接数据库
# engine = create_engine(
#     'mysql+pymysql://leiming:pQx2WhYhgJEtU5r@rm-2ze314ym42f9iq2xflo.mysql.rds.aliyuncs.com:3306/plutus')
# conn = pymysql.connect(host='rm-2ze314ym42f9iq2xflo.mysql.rds.aliyuncs.com',
#                        port=3306, user='leiming',
#                        passwd='pQx2WhYhgJEtU5r',
#                        db="plutus",
#                        charset='utf8')

# 连接数据库(测试)
engine = create_engine(
    'mysql+pymysql://leiming:vg4wHTnJlbWK8SY@rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com:3306/plutus')
conn = pymysql.connect(host='rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com',
                            port=3306, user='leiming',
                            passwd='vg4wHTnJlbWK8SY',
                            db="plutus",
                            charset='utf8')
# 读取数据
PATH = '/Users/edz/Documents'
url ='https://erp.banmaerp.com/Product/Spu/ExportHandler'
data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%7D'
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
r = requests.post(url=url, headers=headers, data=data)
file_name = PATH + '/本地产品导出.xlsx'.format(dt.datetime.now().date())
with open(file_name, 'wb') as file:
    file.write(r.content)
data_cp = read_table(file_name)
os.remove(file_name)
# 删除第一列主标题
if "本地产品" in data_cp.columns.tolist():
    data_cp = pd.DataFrame(data_cp.iloc[1:].values, columns=data_cp.iloc[0, :])
print(data_cp.columns)
print(data_cp.head())
# 增加specs_one，specs_two，is_delete，category项
data_cp['specs_one'] = data_cp['规格']
data_cp['specs_two'] = data_cp['规格']
data_cp['is_delete'] = np.where(data_cp['状态'] == '已删除', 1, 0)
data_cp['category'] = data_cp['斑马类目']

# 删除spu 和sku状态为已删除的records
data_cp['delete'] = data_cp['is_delete']
data_cp['delete'] = np.where(data_cp['SPU状态'] == '已删除', 1, data_cp['delete'])
data_cp = data_cp[data_cp['delete'] != 1]
data_cp = data_cp.drop(columns='delete')
data_cp = data_cp.reset_index()

# 修改specs_one(color) specs_two(size) spu图集(用','分割)
for i in range(data_cp.shape[0]):
    # 修改category为品类的根结点
    data_cp.loc[i, 'category'] = str(data_cp.loc[i, 'category']).split('»')[-1]
    data_cp.loc[i, 'SPU图集'] = data_cp.loc[i, 'SPU图集'].replace('\n', ',')
    if len(data_cp.loc[i, 'specs_two'].split(';')) >= 2:
        data_cp.loc[i, 'specs_two'] = data_cp.loc[i, 'specs_two'].split(';')[1]
        data_cp.loc[i, 'specs_one'] = data_cp.loc[i, 'specs_one'].split(';')[0]
    elif len(data_cp.loc[i, 'specs_two']) > 2 and data_cp.loc[i, 'specs_two'] != 'One Size':
        data_cp.loc[i, 'specs_one'] = data_cp.loc[i, 'specs_one']
        data_cp.loc[i, 'specs_two'] = np.nan
    else:
        data_cp.loc[i, 'specs_two'] = data_cp.loc[i, 'specs_two']
        data_cp.loc[i, 'specs_one'] = np.nan
# size同类合并
data_cp['specs_two'] = np.where(
    (data_cp['specs_two'] == 'One-Size') | (data_cp['specs_two'] == 'one-size') | (data_cp['specs_two'] == 'One Size'),
    'One Size', data_cp['specs_two'])

# 得到size 和color的唯一值(用于创建product_attr表)
specs_two = data_cp['specs_two'].unique()
specs_one = data_cp['specs_one'].unique()
# 删除含有中文字符的值
specs_two = is_chinese(specs_two)
specs_one = is_chinese(specs_one)
for i in range(data_cp.shape[0]):
    if data_cp.loc[i, '标题'].startswith('\"'):
        data_cp.loc[i, '标题'] = data_cp.loc[i, '标题'].replace('\"','\'')
# 给数据库中product表插入数据:
"""
product 插入数据
"""
data_cp.to_excel('/Users/edz/Documents/data_cp.xlsx')
# 插入data_cp表中spu数据
for i in range(data_cp.shape[0]):
    with conn.cursor() as cursor:
        # 以spu_code为primary key 进行插入数据
        sql = "select spu_code from product where spu_code='{0}'".format(data_cp.loc[i, 'SPU编码'])
        cursor.execute(sql)
        r = cursor.fetchone()
        if r is None:
            sql = '''INSERT INTO product (product_name,spu_code, primary_image, add_time, product_images, zebra_spu_id) VALUES ("{0}",'{1}','{2}',now(),'{3}',{4})'''.format(
                data_cp.loc[i, '标题'], data_cp.loc[i, 'SPU编码'], data_cp.loc[i, 'SPU图片'],
                data_cp.loc[i, 'SPU图集'], int(data_cp.loc[i, '系统SPUID']))
            engine.execute(sql)
        else:
            sql = '''UPDATE product SET product_name ="{0}",primary_image = "{2}",add_time=now(),product_images="{3}",zebra_spu_id={4} WHERE spu_code = "{1}"'''.format(
                data_cp.loc[i, '标题'], data_cp.loc[i, 'SPU编码'], data_cp.loc[i, 'SPU图片'],
                data_cp.loc[i, 'SPU图集'], int(data_cp.loc[i, '系统SPUID']))
            engine.execute(sql)
print('刷完产品')
"""
更新data_cp表中的product_id
"""
# 取出刚刚写入数据库里的product表及其id，根据spu，插入到data_cp里
data_p_id = pd.read_sql_table('product', engine)
data_p_id = data_p_id[['id', 'spu_code']]
data_cp = data_cp.merge(data_p_id, left_on='SPU编码', right_on='spu_code')

# 给数据库中product attr表插入数据
# 插入color属性
"""
product_attr 插入数据
需要: specs_one, specs_two 两个关于color属性和size属性的table
"""
# for i in range(len(specs_one)):
#     with conn.cursor() as cursor:
#         sql = "select attr_name from product_attr where attr_name='{0}'".format(specs_one[i])
#         cursor.execute(sql)
#         r = cursor.fetchone()
#         if r is None:
#             sql = "INSERT INTO product_attr (attr_name, parent_id, ancilla) VALUES ('{0}', 1, NULL)".format(
#                 specs_one[i])
#             engine.execute(sql)
#
# # 插入size属性
# for i in range(len(specs_two)):
#     with conn.cursor() as cursor:
#         sql = "select attr_name from product_attr where attr_name='{0}'".format(specs_two[i])
#         cursor.execute(sql)
#         r = cursor.fetchone()
#         if r is None:
#             sql = "INSERT INTO product_attr (attr_name, parent_id, ancilla) VALUES ('{0}', 2, NULL)".format(
#                 specs_two[i])
#             engine.execute(sql)
"""
更新data_cp表中的specs_one_id和specs_two_id
删除data_cp中属性含有中文字，并把属性id同步到data_cp表中
"""
# 将插入完成后的product_attr表读出，
data_product_attr = pd.read_sql_table('product_attr', engine)
# 删除data_cp里，color或size属性带中文字符的records
for i in range(data_cp.shape[0]):
    if not data_cp.loc[i, 'specs_one'] in specs_one:
        data_cp.loc[i, 'specs_one'] = -1
    if not data_cp.loc[i, 'specs_two'] in specs_two:
        data_cp.loc[i, 'specs_two'] = -1
data_cp = data_cp[~((data_cp['specs_two'] == -1) | (data_cp['specs_one'] == -1))]
# 并且通过合并product_attr表，来获取每行size和color属性对应的属性id
cur = data_cp.merge(data_product_attr, left_on='specs_one', right_on='attr_name', how='left')
data_cp = cur.merge(data_product_attr, left_on='specs_two', right_on='attr_name', how='left')
data_cp = data_cp.astype(object).where(pd.notnull(data_cp), "NULL")

# 添加sku main进数据库:
"""
sku_main插入数据
需要data_cp(包括更新的product_id 和specs_id)
"""

for i in range(data_cp.shape[0]):
    with conn.cursor() as cursor:
        # 以sku_code为primary key 进行插入数据，查看要插入的数据sku
        sql = "select sku_code from sku_main where sku_code='{0}'".format(data_cp.loc[i, 'SKU编码'])
        cursor.execute(sql)
        r = cursor.fetchone()
        # 如果返回为none，则说明该sku不存在于数据库，进行插入操作
        if r is None:
            sql = 'INSERT INTO sku_main (sku_code,product_id ,specs_one, specs_two, specs_three, ' \
                  'cost_price, cost_currency, sale_price, sale_currency,' \
                  'sku_style, primary_image, is_delete, add_time,' \
                  'secondary_images, weight, length, height, width, name,' \
                  'en_name, is_effective, zebra_sku_id) ' \
                  'VALUES ("{0}",{1},{2},{3},NULL,{4},"RMB",NULL,"USD",NULL,"{5}",{6},now(),"{7}",{8},{9},{10},{11},NULL,NULL, 1,{12})'.format(
                data_cp.loc[i, 'SKU编码'], data_cp.loc[i, 'id_x'], data_cp.loc[i, 'id_y'], data_cp.loc[i, 'id'],
                data_cp.loc[i, '成本价'], data_cp.loc[i, 'SKU图'], data_cp.loc[i, 'is_delete'],
                data_cp.loc[i, 'SPU图集'], data_cp.loc[i, '重量'], data_cp.loc[i, '长'], data_cp.loc[i, '高'],
                data_cp.loc[i, '宽'], int(data_cp.loc[i, 'SKUID']))
            engine.execute(sql)
        else:
            sql = '''UPDATE sku_main SET product_id ={1},specs_one = {2},specs_two={3},cost_price={4},cost_currency="RMB", sale_currency = "USD",primary_image = "{5}",
                    is_delete= {6},add_time = now(),secondary_images = "{7}", weight = {8}, length = {9},height ={10}, width = {11}, is_effective = 1,zebra_sku_id = {12}
                    WHERE sku_code = "{0}"'''.format(
                data_cp.loc[i, 'SKU编码'], data_cp.loc[i, 'id_x'], data_cp.loc[i, 'id_y'], data_cp.loc[i, 'id'],
                data_cp.loc[i, '成本价'], data_cp.loc[i, 'SKU图'], data_cp.loc[i, 'is_delete'],
                data_cp.loc[i, 'SPU图集'], data_cp.loc[i, '重量'], data_cp.loc[i, '长'], data_cp.loc[i, '高'],
                data_cp.loc[i, '宽'], int(data_cp.loc[i, 'SKUID']))
            engine.execute(sql)

print('刷完sku_main')


"""
插入product_tag表所有标签
需要data_cp中所有的标签集合
"""
# 设置tag list来储存所有标签属性(unique)，剔除所有标签为空的records
tag = []
notnull_cp = data_cp[~(data_cp['标签'] == "NULL")]
for i in range(notnull_cp.shape[0]):
    tag += str(notnull_cp.iloc[i, 4]).split(',')
tag = list(set(tag))
# 将得到的标签属性值导入到数据库的product_tag表中，得到tag对应的tag_id
# for i in range(len(tag)):
#     with conn.cursor() as cursor:
#         sql = '''SELECT * FROM product_tag WHERE tag_name = "{0}" '''.format(tag[i])
#         cursor.execute(sql)
#         r = cursor.fetchone()
#         if r is None:
#             sql = '''INSERT INTO  product_tag (tag_name, add_time) VALUES ("{0}",now())'''.format(tag[i])
#             engine.execute(sql)

# 设置id list和tag list 将data_cp中的id和该id对应的多个tag组成二元tuple
tr_id = []
tr_tag = []
notnull_cp = notnull_cp.reset_index()
for i in range(notnull_cp.shape[0]):
    if ',' not in str(notnull_cp.loc[i, '标签']):
        tr_id.append(notnull_cp.loc[i, 'id_x'])
        tr_tag.append(notnull_cp.loc[i, '标签'])
    else:
        for tags in str(notnull_cp.loc[i, '标签']).split(','):
            if len(tags) > 1:
                tr_id.append(notnull_cp.loc[i, 'id_x'])
                tr_tag.append(tags)
tuples = list(zip(tr_id, tr_tag))
# 将这两列转化为dataframe
tr = pd.DataFrame(tuples, columns=['product_id', 'tags_name'])
# 删除重复项
tr = tr.drop_duplicates()
# 读出product_tag得到tag及其对应的id，将tag_id通过tag_name合并到product_id上
product_tag = pd.read_sql_table('product_tag', engine)
tr = tr.merge(product_tag, left_on='tags_name', right_on='tag_name', how='left')
tr = tr.dropna(subset=['id'])
tr = tr.reset_index()
"""
插入product_tag_relation表所有tag_id和product_id对应关系
需要tr表(有tag_id 和 product_id 以及 tag_name)
"""
# 将tag_id，product_id写入到product_tag_relation表
for i in range(tr.shape[0]):
    with conn.cursor() as cursor:
        sql = '''SELECT * FROM product_tag_relation WHERE tag_id = {0} and product_id = {1}'''.format(tr.loc[i, 'id'], tr.loc[i, 'product_id'])
        cursor.execute(sql)
        r = cursor.fetchone()
        if r is None:
            sql = '''INSERT INTO  product_tag_relation (tag_id, product_id) VALUES ({0},{1})'''.format(
                tr.loc[i, 'id'], tr.loc[i, 'product_id'])
            engine.execute(sql)
print('刷完product_tag_relation')
"""
更新product中的supplier_id数据
需要supplier表和data_cp
"""
# 从数据库中读出供应商表，并筛选出supplier_name和对应的id
supplier = pd.read_sql_table('supplier', engine)
supplier = supplier[['id', 'supplier_name']]
supplier.rename(columns={'id': 'supplier_id'}, inplace=True)
# 将供应商id加到data_cp中，通过供应商名字
data_cp = data_cp.merge(supplier, left_on='默认供应商', right_on='supplier_name', how='left')

# 更新product表中的供应商id
for i in range(data_cp.shape[0]):
    with conn.cursor() as cursor:
        try:
            sql = 'UPDATE product SET supplier_id ={0} WHERE spu_code = "{1}"'.format(data_cp.loc[i, 'supplier_id'],
                                                                                      data_cp.loc[i, 'SPU编码'])
            engine.execute(sql)
        except:
            continue
print('刷完product中supplier id')
# 从数据库中读出品类，并筛选出category_name和对应的id
category = pd.read_sql_table('product_category', engine)
category = category[['id', 'category_name']]

# 删除品类中的字符串的首位空格
for i in range(data_cp.shape[0]):
    data_cp.loc[i, 'category'] = trim(data_cp.loc[i, 'category'])
category.rename(columns={'id': 'category_id'}, inplace=True)

# 将品类id对应带data_cp上通过category
data_cp = data_cp.merge(category, left_on='category', right_on='category_name', how='left')
# data_cp.to_excel('/Users/edz/Documents/data_cp.xlsx')
data_cp = data_cp.dropna(subset = ['category_id'])
data_cp = data_cp.reset_index()
"""
更新product表中的category_id
data_cp表中的category和product_category中的id
"""
# 更新product中的品类id
for i in range(data_cp.shape[0]):
    with conn.cursor() as cursor:
        sql = 'UPDATE product SET product_category={0} WHERE spu_code = "{1}"'.format(data_cp.loc[i, 'category_id'],
                                                                                          data_cp.loc[i, 'SPU编码'])
        engine.execute(sql)
print('刷完product中product category id')
# 从数据库product表中读取供应商id和产品id
sup = pd.read_sql_table('product', engine)
sup = sup[['id', 'supplier_id']]
sup = sup[~sup['supplier_id'].isnull()][['supplier_id', 'id']]
# 删除重复项
sup = sup.drop_duplicates()
sup = sup.reset_index()
"""
插入product_supplier表中supplier_id, product_id
需要product表获取product_id和supplier_id
"""
# 将供应商id和产品id导入到product_supplier表中
for i in range(sup.shape[0]):
    with conn.cursor() as cursor:
        sql = '''SELECT * FROM product_supplier WHERE supplier_id = {0} AND product_id = {1}'''.format(
        sup.iloc[i, 0], sup.iloc[i, 1])
        cursor.execute(sql)
        r = cursor.fetchone()
        if r is None:
            sql = '''INSERT INTO  product_supplier (supplier_id, product_id) VALUES ({0}, {1})'''.format(
                sup.iloc[i, 0], sup.iloc[i, 1])
            engine.execute(sql)
print('刷完product_supplier')
# 更新sku_id_code_dic数据库
sku_id_code_dic = data_cp[['SKUID', '系统SPUID', 'SKU编码', '成本价', '重量']]
sku_id_code_dic = sku_id_code_dic.drop_duplicates()
sku_id_code_dic = sku_id_code_dic.reset_index()


# for i in range(sku_id_code_dic.shape[0]):
#     with conn.cursor() as cursor:
#         # 以spu_id sku_id,为primary key 进行插入数据
#         sql = "select sku_code from sku_id_code_dic where spu_id='{0}' and sku_id='{1}'".format(sku_id_code_dic.loc[i, '系统SPUID'],sku_id_code_dic.loc[i, 'SKUID'])
#         cursor.execute(sql)
#         r = cursor.fetchone()
#         if r is None:
#             sql = '''INSERT INTO sku_id_code_dic (sku_id,spu_id, sku_code, sku_price, sku_weight) VALUES ({0},{1},'{2}',{3},{4})'''.format(
#                 int(sku_id_code_dic.loc[i, 'SKUID']), int(sku_id_code_dic.loc[i, '系统SPUID']), sku_id_code_dic.loc[i, 'SKU编码'],
#                 sku_id_code_dic.loc[i, '成本价'], sku_id_code_dic.loc[i, '重量'])
#             engine.execute(sql)


url = 'https://erp.banmaerp.com/Stock/SelfInventory/ExportDetailHandler'
data = 'filter=%7B%22Quantity%22%3A%7B%22Sort%22%3A-1%7D%2C%22WarehouseID%22%3A%7B%22Value%22%3A%5B%22adac18f9-a30e-4a4b-937f-ac6700e80334%22%5D%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A10000%2C%22PageNumber%22%3A1%7D%7D'
r = requests.post(url=url, headers=headers, data=data)
file_name = PATH + '/本地产品导出.xlsx'.format(dt.datetime.now().date())
with open(file_name, 'wb') as file:
    file.write(r.content)
d = read_table(file_name)
print(d.head())
print(d.columns)
data = xlrd.open_workbook(file_name)
os.remove(file_name)
table = data.sheets()[0]
nrows = table.nrows
col_dic = {}
index = 1

cur_test = conn.cursor()
# 获取字段名称
for col_index in table.row(0):
    col_dic[index] = col_index.value
    index += 1

# 开始处理数据
for row in range(1, nrows):
    print(row)
    data_list = []
    i = 1
    col_item_dic = {}

    # 获取一行数据
    for col in table.row(row):
        col_item_dic[col_dic[i]] = col.value
        i += 1
    # 判断货位是否存在
    sql = '''select id from warehouse_location where warehouse_location_code='{0}' and warehouse_id = 1'''.format(col_item_dic['货位'])
    cur_test.execute(sql)
    r = cur_test.fetchone()
    if r is None:
        sql = '''insert into warehouse_location(warehouse_id, warehouse_location_code) values(1, '{0}')'''.format(col_item_dic['货位'])
        print(sql)
        cur_test.execute(sql)
        location_id = conn.insert_id()
        print('插入新货位成功')
        print(location_id)
        conn.commit()
    else:
        location_id = r[0]
print('刷完库位')
    # # 判断是否有SKU
    # get_sku_id_sql = '''select id from sku_main where sku_code = '{0}' '''.format(col_item_dic['本地SKU'])
    # cur_test.execute(get_sku_id_sql)
    # r = cur_test.fetchone()
    # if r is None:
    #     print(col_item_dic['本地SKU'] + '不存在sku_main里面！！')
    #     continue
    # else:
    #     sku_id = r[0]

    # # 更新库存
    # total_num = col_item_dic['库存总量'] if '库存总量' in col_item_dic else 'NULL'
    # free_num = col_item_dic['合格空闲量'] if '合格空闲量' in col_item_dic else 'NULL'
    # lock_num = col_item_dic['合格锁定量'] if '合格锁定量' in col_item_dic else 'NULL'
    # imperfect_num = col_item_dic['残次总量'] if '残次总量' in col_item_dic else 'NULL'
    #
    # total_num = int(total_num) if total_num != '' else 'NULL'
    # free_num = int(free_num) if free_num != '' else 'NULL'
    # lock_num = int(lock_num) if lock_num != '' else 'NULL'
    # imperfect_num = int(imperfect_num) if imperfect_num != '' else 'NULL'
    #
    # get_exist_stock = '''select id from warehouse_stock where sku_id={0} and warehouse_id = 1 and warehouse_location_id = {1}'''.format(sku_id, location_id)
    # cur_test.execute(get_exist_stock)
    # r = cur_test.fetchone()
    # if r is None:
    #     insert_sql = '''insert into warehouse_stock(sku_id,warehouse_id,warehouse_location_id,total_num,free_num,lock_num,imperfect_num)
    #                     values({0},1,{1},{2},{3},{4},{5})'''.format(sku_id, location_id, total_num, free_num, lock_num, imperfect_num)
    #     # print(insert_sql)
    #     cur_test.execute(insert_sql)
    #     conn_test.commit()
    # else:
    #     update_sql = '''update warehouse_stock set total_num = {0}, free_num = {1}, lock_num = {2}, imperfect_num = {3}
    #                     where sku_id = {4} and warehouse_id = {5} and warehouse_location_id = {6}'''.format(total_num, free_num, lock_num, imperfect_num, sku_id, 1, location_id)
    #     # print(update_sql)
    #     cur_test.execute(update_sql)
    #     conn_test.commit()