import datetime as dt
import requests
import openpyxl as op
import pandas as pd


# 获取昨天订单数据

def read_table(path):
    wb = op.load_workbook(path)
    ws = wb.active
    df = pd.DataFrame(ws.values)
    df = pd.DataFrame(df.iloc[1:].values, columns=df.iloc[0, :])
    return df


today = dt.date.today()
yesterday = today - dt.timedelta(days=1)
path = '/Users/edz/Documents/'
url = 'https://erp.banmaerp.com/Order/Package/ExportHandler'
data = 'filter=%7B%22ID%22%3A%7B%22Sort%22%3A-1%7D%2C%22Details%22%3A%7B%22Filter%22%3A%7B%7D%7D%2C%22Tags%22%3A%7B%22Mode%22%3A0%7D%2C%22Pager%22%3A%7B%22PageNumber%22%3A1%2C%22PageSize%22%3A20%7D%2C%22DeliveryTime%22%3A%7B%22StartValue%22%3A%22{0}+00%3A00%3A00.0000%22%2C%22EndValue%22%3A%22{1}+23%3A59%3A59.9999%22%7D%2C%22Status%22%3A%7B%22Value%22%3A%5B4%5D%7D%2C%22InterceptStatus%22%3A%7B%22Value%22%3A%5B0%2C2%2C3%2C4%5D%7D%7D&details%5B0%5D%5BFieldID%5D=73&details%5B0%5D%5BSort%5D=1&details%5B0%5D%5BFieldExportName%5D=%E8%AE%A2%E5%8D%95%E5%8F%B7&details%5B1%5D%5BFieldID%5D=79&details%5B1%5D%5BSort%5D=2&details%5B1%5D%5BFieldExportName%5D=%E7%89%A9%E6%B5%81%E6%96%B9%E5%BC%8F&details%5B2%5D%5BFieldID%5D=80&details%5B2%5D%5BSort%5D=3&details%5B2%5D%5BFieldExportName%5D=%E7%89%A9%E6%B5%81%E5%8D%95%E5%8F%B7&details%5B3%5D%5BFieldID%5D=82&details%5B3%5D%5BSort%5D=4&details%5B3%5D%5BFieldExportName%5D=%E4%B8%8B%E5%8D%95%E6%97%B6%E9%97%B4&details%5B4%5D%5BFieldID%5D=84&details%5B4%5D%5BSort%5D=5&details%5B4%5D%5BFieldExportName%5D=%E5%8F%91%E8%B4%A7%E6%97%B6%E9%97%B4&details%5B5%5D%5BFieldID%5D=107&details%5B5%5D%5BSort%5D=6&details%5B5%5D%5BFieldExportName%5D=%E9%87%8D%E9%87%8F%EF%BC%88KG%EF%BC%89&details%5B6%5D%5BFieldID%5D=243&details%5B6%5D%5BSort%5D=7&details%5B6%5D%5BFieldExportName%5D=%E8%AE%A2%E5%8D%95%E9%87%91%E9%A2%9D(USD)&details%5B7%5D%5BFieldID%5D=94&details%5B7%5D%5BSort%5D=8&details%5B7%5D%5BFieldExportName%5D=%E6%94%B6%E4%BB%B6%E4%BA%BA%E5%9B%BD%E5%AE%B6&details%5B8%5D%5BFieldID%5D=1274&details%5B8%5D%5BSort%5D=9&details%5B8%5D%5BFieldExportName%5D=%E4%BA%A4%E6%98%93%E5%8F%B7&type=2'.format(
    yesterday, yesterday)
headers = {
    'content-type': 'application/x-www-form-urlencoded; charset=UTF-8',
    'user-agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_14_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/86.0.4240.80 Safari/537.36',
    'cookie': '.AspNetCore.Session=CfDJ8HFZt5KhGHxPrfAKn%2Fe35kaRpPerMJVnDOQnJCjicT8lyd81AtsUwStenh5nUMsWpyuS%2Bu38igf9ADjk2fhr6CYTk87TukhPs3Uqvid6CI4gSaSqYkM7fHDGw4xEnUKIIhoVh5nzaNU57l2OfpixmIgipBDXzggD1pciKOzkXQdc; Hm_lvt_9be79ac4f097e2a0be24ee6c088e921b=1603200345,1603247430; ERP.Token=eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJJRCI6Ijc1MjIiLCJOYW1lIjoi6Zu35pmT5pmoIiwiVXNlclR5cGUiOiIzIiwiT3duVXNlcklEIjoiNzA0MCIsImV4cCI6MTYzNDc5MzM3MSwiaXNzIjoiRVJQLmJhbm1hZXJwLmNvbSIsImF1ZCI6IkVSUC5iYW5tYWVycC5jb20ifQ.r5r1FrpMRa_yWr3qxuLnrJXUAZST_CC6V8nt2V-MbxM; Hm_lpvt_9be79ac4f097e2a0be24ee6c088e921b=1603257395'}

r = requests.post(url=url, headers=headers, data=data)
file_name = path + '{0}订单.xlsx'.format(yesterday)
with open(file_name, 'wb') as file:
    file.write(r.content)
data_dd = read_table(file_name)
print(data_dd.columns)
data_dd = data_dd.dropna()
data_dd = data_dd[~(
        data_dd['交易号'].str.startswith('pi') | data_dd['交易号'].str.startswith('.pi') | data_dd['交易号'].str.startswith(
    ',pi'))]
print(data_dd.head())

post_url = 'https://api-m.sandbox.paypal.com/v1/shipping/trackers-batch'
post_headers = {'X-PAYPAL-SECURITY-CONTEXT:Content-Type: application/json'}
data_dd = data_dd.reset_index()
print(data_dd.shape[0])
post_data = {'trackers': []}
start = 0
if data_dd.shape[0] > 20:
    while data_dd.shape[0] - start > 20:
        for i in range(start, start + 20):
            cur = {'transaction_id': data_dd.loc[i, '交易号'],
                   'tracking_number': data_dd.loc[i, '物流单号'],
                   'status': 'SHIPPED',
                   'carrier': 'OTHER',
                   'carrier_name_other': data_dd.loc[i, '物流方式']}
            post_data.get('trackers').append(cur)
        print(post_data.get('trackers'))
        print('------------------------------------------')
        print(len(post_data.get('trackers')))
        # response = requests.post(post_url, post_headers, post_data)
        post_data = {'trackers': []}
        start += 20
for i in range(start, data_dd.shape[0]):
    cur = {'transaction_id': data_dd.loc[i, '交易号'],
           'tracking_number': data_dd.loc[i, '物流单号'],
           'status': 'SHIPPED',
           'carrier': 'OTHER',
           'carrier_name_other': data_dd.loc[i, '物流方式']}
    post_data.get('trackers').append(cur)
# response = requests.post(post_url, post_headers, post_data)
print(post_data)
