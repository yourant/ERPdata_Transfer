import openpyxl as op
import pandas as pd
import psycopg2
import requests
import os

def read_table(path):
    wb = op.load_workbook(path)
    ws = wb.active
    df = pd.DataFrame(ws.values)
    df = pd.DataFrame(df.iloc[1:].values, columns=df.iloc[0, :])
    return df
conn_pg_test = psycopg2.connect(database="plutus", user="plutus", password="2JQsCVddyjOADRy",
                                host="pgm-2zetb1em3zlbjfi9168190.pg.rds.aliyuncs.com", port="1921")
cur_pg_test = conn_pg_test.cursor()
# conn_pg_ol = psycopg2.connect(database="plutus", user="plutus", password="4c5I6hxUmo8khujZdrhS", host="pgm-2ze7v274je7y18ba167580.pg.rds.aliyuncs.com", port="1921")
# cur_pg_ol = conn_pg_ol.cursor()
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
# 采购sku信息
url_cgd = 'https://erp.banmaerp.com/Purchase/Sheet/ExportPurchaseHandler'
data_cgd2020 = 'filter=%7B%22CreateTime%22%3A%7B%22StartValue%22%3A%222020-1-1+00%3A00%3A00.000%22%2C%22EndValue%22%3A%222020-12-31+23%3A59%3A59.998%22%2C%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%7D'
data_cgd2021 = 'filter=%7B%22CreateTime%22%3A%7B%22StartValue%22%3A%222021-1-1+00%3A00%3A00.000%22%2C%22EndValue%22%3A%222021-12-31+23%3A59%3A59.998%22%2C%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%7D'
r = requests.post(url = url_cgd,headers = headers,data = data_cgd2020)
file_name = os.getcwd() + '/采购单2020.xlsx'
with open(file_name, 'wb') as file:
    file.write(r.content)
cgd2020 = read_table('采购单2020.xlsx')
r = requests.post(url = url_cgd,headers = headers,data = data_cgd2021)
file_name = os.getcwd() + '/采购单2021.xlsx'
with open(file_name, 'wb') as file:
    file.write(r.content)
cgd2021 = read_table('采购单2021.xlsx')

print(cgd2020.columns)
print(cgd2020.shape[0])
print(cgd2021.shape[0])
cgd_data = pd.concat([cgd2020, cgd2021],ignore_index=True)
print(cgd_data.shape[0])

# 用户信息
url_user = 'https://erp.banmaerp.com/Account/User/ListData'
data_user = 'filter=%7B%22ID%22%3A%7B%22Sort%22%3A-1%7D%2C%22Status%22%3A%7B%22Value%22%3A0%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A50%7D%7D&pageNumber=1&pageSize=50'
r = requests.post(url=url_user, headers=headers, data=data_user)
user_map = {}
res = r.json()['Data']['Results']
for i in range(len(res)):
    user_map[res[i]['ID']] = res[i]['RealName']

print(user_map)

print()
# 采购单数据
url = 'https://erp.banmaerp.com/Purchase/Sheet/ListData'
page_number = 1
data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A{0}%2C%22PageSize%22%3A100%7D%7D&pageNumber={0}&pageSize=100'.format(
    page_number)

# data_p_id = pd.read_sql_query('select * from warehouse', con=conn_pg_ol)
# data_s = pd.read_sql_query('select * from supplier', con=conn_pg_ol)
# data_u = pd.read_sql_query('select * from "user" ', con=conn_pg_ol)
# data_u.to_excel('user.xlsx')

response = requests.post(url=url, headers=headers, data=data)
all_results = response.json()['Results']
print(all_results[0])
ans = 0
#
while True:
    data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A{0}%2C%22PageSize%22%3A100%7D%7D&pageNumber={0}&pageSize=100'.format(
        page_number)
    response = requests.post(url=url, headers=headers, data=data)
    if len(response.json()['Results']) == 0:
        break
    all_results = response.json()['Results']
    for i in range(len(all_results)):
        if all_results[i]['Sheet']['WarehouseName'] == '自建仓-坑头':
            Warehouse_id = 1
            create_id_banma = all_results[i]['Sheet']['CreateUserID']
            create_name_banma = user_map[create_id_banma]
            sql = """SELECT id FROM "user" WHERE name ='{0}' """.format(create_name_banma)
            cur_pg_test.execute(sql)
            r = cur_pg_test.fetchone()
            if r is None:
                sql = """INSERT INTO "user" (name, password) VALUES ('{0}','53056de5d44b0f2d0799e154df793eed') """.format(
                    create_name_banma)
                cur_pg_test.execute(sql)
                conn_pg_test.commit()
                sql = """SELECT id FROM "user" WHERE name ='{0}' """.format(create_name_banma)
                cur_pg_test.execute(sql)
                r = cur_pg_test.fetchone()
                create_id = r[0]
            else:
                create_id = r[0]

            supplier_name = all_results[i]['Sheet']['SupplierName']
            sql = """select id from supplier where supplier_name = '{0}' """.format(supplier_name)
            cur_pg_test.execute(sql)
            r = cur_pg_test.fetchone()
            if r is None:
                sql = """INSERT INTO supplier (supplier_name, remark, create_uid) VALUES ('{0}','结算方式',1) """.format(
                    supplier_name.replace("'", "''"))
                cur_pg_test.execute(sql)
                conn_pg_test.commit()
                sql = """select id from supplier where supplier_name = '{0}' """.format(supplier_name)
                cur_pg_test.execute(sql)
                supplier_id = cur_pg_test.fetchone()[0]
            else:
                supplier_id = r[0]
            status = all_results[i]['Sheet']['Status']
            add_time = all_results[i]['Sheet']['CreateTime']
            expectArrivalTime = all_results[i]['Sheet']['ExpectArrivalTime']
            confirmed_time = all_results[i]['Sheet']['ConfirmTime']

            update_id_banma = all_results[i]['Sheet']['UpdateUserID']
            update_name_banma = user_map[update_id_banma]
            sql = """SELECT id FROM "user" WHERE name ='{0}' """.format(update_name_banma)
            cur_pg_test.execute(sql)
            r = cur_pg_test.fetchone()
            if r is None:
                sql = """INSERT INTO "user" (name, password) VALUES ('{0}','53056de5d44b0f2d0799e154df793eed') """.format(
                    update_name_banma)
                cur_pg_test.execute(sql)
                conn_pg_test.commit()
                sql = """SELECT id FROM "user" WHERE name ='{0}' """.format(update_name_banma)
                cur_pg_test.execute(sql)
                r = cur_pg_test.fetchone()
                update_id = r[0]
            else:
                update_id = r[0]

            purchase_name = all_results[i]['Sheet']['PurchaseUserName']
            sql = """SELECT id FROM "user" WHERE name ='{0}' """.format(purchase_name)
            cur_pg_test.execute(sql)
            r = cur_pg_test.fetchone()
            if r is None:
                sql = """INSERT INTO "user" (name, password) VALUES ('{0}','53056de5d44b0f2d0799e154df793eed') """.format(
                    purchase_name)
                cur_pg_test.execute(sql)
                conn_pg_test.commit()
                sql = """SELECT id FROM "user" WHERE name ='{0}' """.format(purchase_name)
                cur_pg_test.execute(sql)
                r = cur_pg_test.fetchone()
                purchase_uid = r[0]
            else:
                purchase_uid = r[0]

            update_time = all_results[i]['Sheet']['UpdateTime']
            is_delete_boolean = all_results[i]['Sheet']['IsDelete']
            if is_delete_boolean:
                is_delete = 1
            else:
                is_delete = 0

            sql = """INSERT INTO purchase_order (warehouse_id, create_uid, supplier_id, status, expected_arrival_time, add_time, confirmed_time, update_uid, purchase_uid, update_time, is_delete)
            VALUES({0}, {1}, {2}, {3}, '{4}','{5}', '{6}',{7},{8},'{9}','{10}') RETURNING id """.format(Warehouse_id, create_id,
                                                                                               supplier_id, status,
                                                                                               expectArrivalTime,
                                                                                               add_time, confirmed_time,
                                                                                               update_id, purchase_uid,
                                                                                      update_time, is_delete)
            if None in (Warehouse_id, create_id, supplier_id, status, expectArrivalTime,add_time, confirmed_time,update_id, purchase_uid, update_time, is_delete):
                sql = sql.replace("'None'", "NULL")
                sql = sql.replace("None", "NULL")
            cur_pg_test.execute(sql)
            r_id = cur_pg_test.fetchone()[0]
            conn_pg_test.commit()
            cgd_id = all_results[i]['Sheet']['ID']
            sql = """INSERT INTO purchase_order_remark (purchase_order_id, remark) VALUES ({0}, {1}) """.format(
                r_id, cgd_id
            )
            cur_pg_test.execute(sql)
            conn_pg_test.commit()
            all_info = cgd_data[cgd_data['采购单号'] == cgd_id][['本地SKU','物品数量','采购单价','到货物品数量','创建时间']]
            all_info = all_info.reset_index(drop=True)
            for i in range(all_info.shape[0]):
                sku_code = all_info.loc[i, '本地SKU']
                sql = """SELECT id FROM sku_main WHERE sku_code = '{0}' """.format(sku_code)
                cur_pg_test.execute(sql)
                r = cur_pg_test.fetchone()
                if r is not None:
                    sku_id = r[0]
                    sql = """INSERT INTO purchase_order_sku (purchase_order_id, sku_id, num, price,delivered_num,is_delete,add_time,update_time) VALUES ({0},{1}, {2},{3}, {4}, {5},'{6}','{7}') """.format(
                        r_id, sku_id, all_info.loc[i, '物品数量'], all_info.loc[i, '采购单价'], all_info.loc[i, '到货物品数量'],
                        is_delete, all_info.loc[i, '创建时间'], update_time)
                    if None in (
                    r_id, sku_id, all_info.loc[i, '物品数量'], all_info.loc[i, '采购单价'], all_info.loc[i, '到货物品数量'],
                    is_delete, all_info.loc[i, '创建时间'], update_time):
                        sql = sql.replace("'None'", "NULL")
                        sql = sql.replace("None", "NULL")
                    cur_pg_test.execute(sql)
                    conn_pg_test.commit()
            print("finish" + str(ans))
            ans += 1
    page_number += 1

