import xlrd
import pandas as pd
import datetime as dt
import requests
import os
import openpyxl as op
import os


def read_table(path):
    wb = op.load_workbook(path)
    ws = wb.active
    df = pd.DataFrame(ws.values)
    df = pd.DataFrame(df.iloc[1:].values, columns=df.iloc[0, :])
    return df


class retrieve_data(object):
    to_time = dt.datetime.now().date()
    from_time = dt.datetime(2020, 8, 1)
    path = os.getcwd() + '/'
    headers = {
        'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
        'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
        'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}

    def __init__(self, to_time, path=os.getcwd() + '/', from_time='2020-08-01'):
        self.from_time = dt.datetime.strptime(from_time, '%Y-%m-%d').date()
        self.to_time = dt.datetime.strptime(to_time, '%Y-%m-%d').date()
        self.path = path

    def get_dzj_data(self):
        url = 'https://erp.banmaerp.com/Stock/Quality/QualityExportHandler'
        data = 'filter=%7B%7D'
        r = requests.post(url=url, headers=self.headers, data=data)
        file_name = self.path + '待质检数据.xlsx'
        with open(file_name, 'wb') as file:
            file.write(r.content)
        dzj_data = read_table(file_name)
        os.remove(file_name)
        return dzj_data

    def get_pd_data(self):
        url = 'https://erp.banmaerp.com/Product/Platform/ExportSkuMappingHandler'
        data = 'filter=%7B%22CreateTime%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%7D'
        r = requests.post(url=url, headers=self.headers, data=data)
        file_name = self.path + 'SKU配对关系表.xlsx'
        with open(file_name, 'wb') as file:
            file.write(r.content)
        pd_data = read_table(file_name)
        os.remove(file_name)
        return pd_data

    def get_cp_data(self):
        data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A-1%7D%7D'
        url = 'https://erp.banmaerp.com/Shopify/Product/ExportHandler'
        r = requests.post(url=url, headers=self.headers, data=data)
        file_name_cp = self.path + '在线商品数据.xlsx'
        with open(file_name_cp, 'wb') as file:
            file.write(r.content)
        cp_data = read_table(file_name_cp)
        os.remove(file_name_cp)
        return cp_data

    def get_kc_data(self):
        url = 'https://erp.banmaerp.com/Stock/SelfInventory/ExportHandler'
        data = 'filter=%7B%22Quantity%22%3A%7B%22Sort%22%3A-1%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A10000%2C%22PageNumber%22%3A1%7D%7D'
        r = requests.post(url=url, headers=self.headers, data=data)
        file_name_kc = self.path + '库存数据.xlsx'
        with open(file_name_kc, 'wb') as file:
            file.write(r.content)
        kc_data = read_table(file_name_kc)
        os.remove(file_name_kc)
        return kc_data

    def get_dd_data(self):
        data_dd_by_day_list = []
        data_dd = None
        now = dt.datetime(year=dt.datetime.today().year, month=dt.datetime.today().month, day=dt.datetime.today().day)
        yesterday = now - dt.timedelta(days=1)
        one_month_ago = now - dt.timedelta(days=32)
        print(yesterday)
        print(one_month_ago)
        url = 'https://erp.banmaerp.com/Order/Order/ExportOrderHandler'
        data = 'filter=%7B%22ID%22%3A%7B%22Sort%22%3A-1%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22OriginalOrderTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.0000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.9999%22%7D%2C%22Addresses%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%7D&details%5B0%5D%5BFieldID%5D=40&details%5B0%5D%5BSort%5D=1&details%5B0%5D%5BFieldExportName%5D=%E8%AE%A2%E5%8D%95%E7%8A%B6%E6%80%81&details%5B1%5D%5BFieldID%5D=49&details%5B1%5D%5BSort%5D=2&details%5B1%5D%5BFieldExportName%5D=%E4%BB%98%E6%AC%BE%E6%97%B6%E9%97%B4&details%5B2%5D%5BFieldID%5D=68&details%5B2%5D%5BSort%5D=3&details%5B2%5D%5BFieldExportName%5D=%E6%95%B0%E9%87%8F&details%5B3%5D%5BFieldID%5D=70&details%5B3%5D%5BSort%5D=4&details%5B3%5D%5BFieldExportName%5D=%E5%8C%B9%E9%85%8DSKU&details%5B4%5D%5BFieldID%5D=253&details%5B4%5D%5BSort%5D=5&details%5B4%5D%5BFieldExportName%5D=%E7%BC%BA%E8%B4%A7%E6%95%B0%E9%87%8F&details%5B5%5D%5BFieldID%5D=221&details%5B5%5D%5BSort%5D=6&details%5B5%5D%5BFieldExportName%5D=%E5%B9%B3%E5%8F%B0SKU&details%5B6%5D%5BFieldID%5D=61&details%5B6%5D%5BSort%5D=7&details%5B6%5D%5BFieldExportName%5D=%E6%94%AF%E4%BB%98%E9%87%91%E9%A2%9D(CNY)&type=1'.format(
            one_month_ago.date(), yesterday.date())
        r = requests.post(url=url, headers=self.headers, data=data)
        file_name = self.path + '/{0}到{1}订单数据.xlsx'.format(one_month_ago, now)
        with open(file_name, 'wb') as file:
            file.write(r.content)
        data_dd_by_day_list.append(file_name)
        if data_dd is None:
            try:
                data_dd = read_table(file_name)
            except Exception as e:
                print(e)
        file_name_dd = self.path + '{0}到{1}订单数据.xlsx'.format(now, one_month_ago)
        data_dd.to_excel(file_name_dd)
        data_dd_by_day_list.append(file_name_dd)
        for dir_file in data_dd_by_day_list:
            os.remove(dir_file)
        return data_dd

    def get_cgd_data(self):
        # 请求采购单数据
        url = 'https://erp.banmaerp.com/Purchase/Sheet/ExportPurchaseHandler'
        data_cgd = None
        begin_day = '2020-08-01'
        begin_day = dt.datetime.strptime(begin_day, '%Y-%m-%d').date()
        diff_days = self.to_time - begin_day
        diff_days = int(diff_days.days) + 1
        months = diff_days / 30
        data_cgd_by_day_list = []
        temp_date = begin_day
        if months > 0:
            step = 30
            for single_date in (begin_day + dt.timedelta(n) for n in range(30, diff_days, step)):
                data = 'filter=%7B%22UpdateTime%22%3A%7B%22Sort%22%3A%22-1%22%7D%2C%22Pager%22%3A%7B%22PageSize%22%3A5000%7D%2C%22CreateTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.998%22%7D%7D'.format(
                    temp_date, single_date)
                r = requests.post(url=url, headers=self.headers, data=data)
                file_name = self.path + '{0}到{1}采购单数据.xlsx'.format(temp_date, single_date)
                temp_date = single_date
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
            temp_date, self.to_time)
        r = requests.post(url=url, headers=self.headers, data=data)
        file_name = self.path + '{0}到{1}采购单数据.xlsx'.format(temp_date, self.to_time)
        with open(file_name, 'wb') as file:
            file.write(r.content)
        data_cgd_by_day_list.append(file_name)
        if data_cgd is None:
            try:
                data_cgd = read_table(file_name)
            except Exception as e:
                print(e)
                pass
        elif temp_date != self.to_time:
            try:
                data_cgd_cur = read_table(file_name)
                data_cgd = pd.concat([data_cgd, data_cgd_cur], ignore_index=True)
            except Exception as e:
                print(e)
                pass
        for dir_file in data_cgd_by_day_list:
            os.remove(dir_file)
        return data_cgd


if __name__ == '__main__':
    data = retrieve_data('2021-02-19')
    data_cp = data.get_cp_data()
    data_kc = data.get_kc_data()
    data_cgd = data.get_cgd_data()
    data_dd = data.get_dd_data()
    data_pd = data.get_pd_data()
    data_dzj = data.get_dzj_data()
    sheet1 = data_dd[['匹配SKU', '付款时间', '订单状态', '支付金额(CNY)', '数量', '缺货数量']]
    sheet2 = data_cgd[(data_cgd['状态'] == '采购中') & (data_cgd['仓库'] == '自建仓-坑头')]
    sheet2 = sheet2[['本地SKU', '供应商', '物品数量', '到货物品数量', '预计到货时间']]
    sheet3 = data_kc[data_kc['仓库'] == '坑头']
    sheet3 = sheet3[['本地sku', '图片URL', '合格总量', '合格锁定量', '均价（￥）']]
    sheet4 = data_dzj[['本地SKU', '当前入库单SKU到货数量', '当前入库单SKU已质检数量']]
    sheet5 = data_cp[data_cp['店铺'] == 'shopcider']
    sheet5 = sheet5[['Sku', 'Sku图片', '售卖状态', '库存策略', '发布时间']]
    sheet6 = data_pd[['本地SKU', '平台SKU']]
    sheet1 = sheet1.reset_index(drop=True)
    sheet2 = sheet2.reset_index(drop=True)
    sheet3 = sheet3.reset_index(drop=True)
    sheet4 = sheet4.reset_index(drop=True)
    sheet5 = sheet5.reset_index(drop=True)
    sheet6 = sheet6.reset_index(drop=True)
    writer = pd.ExcelWriter('/Users/edz/Documents/{0}预期时间sop.xlsx'.format('2021-02-21'), engine='xlsxwriter')
    sheet1.to_excel(writer, sheet_name='order', index=False)
    sheet2.to_excel(writer, sheet_name='supply', index=False)
    sheet3.to_excel(writer, sheet_name='stock', index=False)
    sheet4.to_excel(writer, sheet_name='inspect', index=False)
    sheet5.to_excel(writer, sheet_name='online', index=False)
    sheet6.to_excel(writer, sheet_name='mapping', index=False)
    writer.save()
