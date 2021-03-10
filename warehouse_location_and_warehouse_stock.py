import pymysql
from openpyxl import Workbook
import datetime
import xlrd
conn_test = pymysql.connect()
cur_test = conn_test.cursor()
data = xlrd.open_workbook('/Users/zhanghuayang/Downloads/库存明细数据导出-20210203214722.xlsx')
table = data.sheets()[0]
nrows = table.nrows
col_dic = {}
index = 1
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
        location_id = conn_test.insert_id()
        print('插入新货位成功')
        print(location_id)
        conn_test.commit()
    else:
        location_id = r[0]
    # 判断是否有SKU
    get_sku_id_sql = '''select id from sku_main where sku_code = '{0}' '''.format(col_item_dic['本地SKU'])
    cur_test.execute(get_sku_id_sql)
    r = cur_test.fetchone()
    if r is None:
        print(col_item_dic['本地SKU'] + '不存在sku_main里面！！')
        continue
    else:
        sku_id = r[0]
    # 更新库存
    total_num = col_item_dic['库存总量'] if '库存总量' in col_item_dic else 'NULL'
    free_num = col_item_dic['合格空闲量'] if '合格空闲量' in col_item_dic else 'NULL'
    lock_num = col_item_dic['合格锁定量'] if '合格锁定量' in col_item_dic else 'NULL'
    imperfect_num = col_item_dic['残次总量'] if '残次总量' in col_item_dic else 'NULL'
    total_num = int(total_num) if total_num != '' else 'NULL'
    free_num = int(free_num) if free_num != '' else 'NULL'
    lock_num = int(lock_num) if lock_num != '' else 'NULL'
    imperfect_num = int(imperfect_num) if imperfect_num != '' else 'NULL'
    get_exist_stock = '''select id from warehouse_stock where sku_id={0} and warehouse_id = 1 and warehouse_location_id = {1}'''.format(sku_id, location_id)
    cur_test.execute(get_exist_stock)
    r = cur_test.fetchone()
    if r is None:
        insert_sql = '''insert into warehouse_stock(sku_id,warehouse_id,warehouse_location_id,total_num,free_num,lock_num,imperfect_num)
                        values({0},1,{1},{2},{3},{4},{5})'''.format(sku_id, location_id, total_num, free_num, lock_num, imperfect_num)
        # print(insert_sql)
        cur_test.execute(insert_sql)
        conn_test.commit()
    else:
        update_sql = '''update warehouse_stock set total_num = {0}, free_num = {1}, lock_num = {2}, imperfect_num = {3}
                        where sku_id = {4} and warehouse_id = {5} and warehouse_location_id = {6}'''.format(total_num, free_num, lock_num, imperfect_num, sku_id, 1, location_id)
        # print(update_sql)
        cur_test.execute(update_sql)
        conn_test.commit()
cur_test.close()
conn_test.close()