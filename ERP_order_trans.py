import requests
import pandas
from sqlalchemy import create_engine
import datetime as dt
engine = create_engine(
    'mysql+pymysql://leiming:vg4wHTnJlbWK8SY@rm-2zeq92vooj5447mqzso.mysql.rds.aliyuncs.com:3306/cider')

# 判断订单是否为全部发货
url = 'https://erp.banmaerp.com/Order/Order/ListDataHandler'
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}
# ddh = 39457268618
# # ddh = 39457197518
# data = 'filter=%7B%22ID%22%3A%7B%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%2C%22DisplayOrderID%22%3A%7B%22Value%22%3A%5B%22%23{0}%22%5D%2C%22Mode%22%3A0%7D%7D&pageNumber=1&pageSize=20'.format(
#     ddh)
# r = requests.get(url=url, headers=headers, data=data)
# orders = r.json()['Results']
# master_id = orders[0]['MasterID']
# print(master_id)
# packages = r.json()['Results'][0]['Details']


# print(packages)
# print(orders[0]['Master'])
# print(orders[0]['Master']['Status'])


# 判断订单状态是否满足
def check_order_states(order, state_List):
    """
    state_List: 0 = '已下单'， 1 = '待审核'， 2 = '待发货'， 3 = '部分发货' ... ...
    """
    return order[0]['Master']['Status'] in state_List


# 判断订单package id是否全部相同且不为0
def check_all_pack_id(packages):
    ans = set()
    for package in packages:
        if package['PackageID'] == '0':
            return False
        else:
            ans.add(package['PackageID'])
        if len(ans) > 1:
            return False
    return True


# 判断是不是全部有货
def check_all_have_stock(packages):
    for package in packages:
        if package['InventoryData'] is not None:
            for small_pac in package['InventoryData']:
                if small_pac['ShortageQuantity'] > 0:
                    return False
        else:
            return False
    return True


# 检查包裹状态
def check_packages_state(packages, pack_state_list):
    """
    state_List: 0 = '已下单'， 1 = '待审核'， 2 = '待发货'， 3 = '部分发货' ... ...
    """
    for package in packages:
        packages_state_url = 'https://erp.banmaerp.com/Order/Package/ListData'
        packages_state_filter = 'filter=%7B%22ID%22%3A%7B%22Value%22%3A%5B%22{0}%22%5D%2C%22Mode%22%3A0%2C%22Sort%22%3A-1%7D%2C%22Details%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%7D&pageNumber=1&pageSize=20'.format(
            package['PackageID'])
        r = requests.post(url=packages_state_url, headers=headers, data=packages_state_filter)
        for pack in r.json()['Results']:
            if pack['Package']['Status'] in pack_state_list:
                return True
    return False


# 暂停包裹
def hold_all_packages(packages):
    for package in packages:
        time = dt.datetime.strftime(dt.datetime.now(),'%Y-%m-%d %H:%M:%S')
        hold_url = 'https://erp.banmaerp.com/Order/Package/Hold'
        f = 'packageIDs={0}&reason=%E6%9A%82%E5%81%9C%E5%8E%9F%E5%9B%A0%EF%BC%9A%E8%AF%A5%E8%AE%A2%E5%8D%95%E8%BF%81%E7%A7%BB%E5%88%B0%E8%87%AA%E5%BB%BAerp%E5%8F%91%E8%B4%A7%2C+{1}'.format(package['PackageID'], time)
        r = requests.post(url=hold_url, headers=headers, data=f)
        print('package_id = ' + str(package['PackageID']) + "   暂停" + r.json()["Message"])


table = pandas.read_sql_table('warehouse_location_info', engine)
print(table.columns)


def get_stock_location_info(Master_id, packages):
    listpac = {}
    for package in packages:
        listpac[package['SKUCode']] = package['Quantity']
    if listpac == {}:
        return []
    result = []
    get_stock_loc_info_url = 'https://erp.banmaerp.com/Stock/Inventory/LockListData'
    get_stock_loc_info_filter = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22BusinessID%22%3A%7B%22Value%22%3A%5B%22{0}%22%5D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%7D&pageNumber=1&pageSize=20'.format(
        Master_id)
    r = requests.post(url=get_stock_loc_info_url, headers=headers, data=get_stock_loc_info_filter)
    for info in r.json()['Data']['Results']:
        if info['WarehouseID'] != 'adac18f9-a30e-4a4b-937f-ac6700e80334':
            return []
        else:
            code = table[table['location_id'] == int(info['Details'][0]['LocationID'])]['location_code'].values[0]
            result.append({'skuCode': info['SkuCode'], 'locationCode': code, 'quantity': listpac.get(info['SkuCode'])})
    return result

def get_data(orderslist):
    INFO =[]
    result = []
    error = {}
    for order in orderslist:
        ddh = order.split('#')[-1]
        data = 'filter=%7B%22ID%22%3A%7B%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%2C%22DisplayOrderID%22%3A%7B%22Value%22%3A%5B%22%23{0}%22%5D%2C%22Mode%22%3A0%7D%7D&pageNumber=1&pageSize=20'.format(
            ddh)
        try:
            r = requests.get(url=url, headers=headers, data=data)
            orders = r.json()['Results']
            master_id = orders[0]['MasterID']
            packages = r.json()['Results'][0]['Details']
            state_List = [2]
            pack_state_list = [0]
            print(str(ddh) + '--' + '都是有库:' + str(check_all_have_stock(packages)),
                  ' 拆分:' + str(check_all_pack_id(packages)), ' 订单状态:' + str(check_order_states(orders, state_List)),
                  ' 包裹状态:' + str(check_packages_state(packages, pack_state_list)))
            if check_all_have_stock(packages) and check_all_pack_id(packages) and check_order_states(orders,
                                                                                                     state_List):
                if check_packages_state(packages, pack_state_list):
                    print('all conditions satisfied')
                    # hold_all_packages(packages)
                    cur_order_info = {'Order': order, 'OID': orders[0]['Master']['OriginalOrderID'],
                                      'orderZebraOutSkuList': get_stock_location_info(master_id, packages)}
                    result.append(cur_order_info)
        except Exception as e:
            error[str(ddh)] = "订单失败:" + str(e)
    if error == {}:
        error = '无失败订单'
    INFO = {'order_info' : result, 'Error': error}
    return INFO

if __name__ == '__main__':
    order = '#39457292918'
    ddh = order.split('#')[-1]
    data = 'filter=%7B%22ID%22%3A%7B%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%2C%22DisplayOrderID%22%3A%7B%22Value%22%3A%5B%22%23{0}%22%5D%2C%22Mode%22%3A0%7D%7D&pageNumber=1&pageSize=20'.format(
        ddh)
    r = requests.get(url=url, headers=headers, data=data)
    orders = r.json()['Results']
    # print(orders[0]['Master'])
    master_id = orders[0]['MasterID']
    # print(master_id)
    packages = r.json()['Results'][0]['Details']
    hold_all_packages(packages)
    # print(packages)
    # get_stock_location_info(master_id,packages)
    #
    # # orderslist = ['#39457333218', '#39457338818', '#39457333118', '#39457333918', '#39457334118','#39457334218']
    # orderslist = ['#39457311718']
    # print(get_data(orderslist))
