import xlrd
import requests
import types
import  datetime as dt
s_date = dt.date (1899, 12, 31).toordinal() - 1
def getdate(date ):
    if isinstance(date , float ):
        date = int(date )
    d = dt.date .fromordinal(s_date + date )
    return d
tokenN=""
#读取联系人的客户名称
data=xlrd.open_workbook(u'赵志勇.xlsx')
table=data.sheet_by_name(u'Sheet1')
for l in range(378,899,1):
    print(l)
    contractView=table.row_values(l)
    costomerName=contractView[4]
    print(costomerName)

    contractName=contractView[0]
    contractPhone=contractView[1]
#处理联系人电话为多个的情况
    print(contractPhone)
    if contractPhone =='':
        contractPhone1 = contractPhone
    elif type(contractPhone)==type('a'):
        contractPhone1=contractPhone
    else:
        contractPhone1 = int(contractPhone)
    contractTNum=contractView[2]
    if contractTNum !='':
        contractTNum1=int(contractTNum)
    else:
        contractTNum1=contractTNum
#处理excel表取出日期的格式
    contractIDate=contractView[3]
    if contractIDate !='':
        contractIDate1=getdate(contractIDate)
        contractIDate2 = contractIDate1.strftime('%Y-%m-%d')
    else:
        contractIDate2=contractIDate

    print(contractPhone1)
    print(contractName)
    print(contractIDate2)


#查询客户的id
    url = "https://www.weibangong.com/api/customer/web/filter/my"
    headers = {"Content-Type":"application/json;charset=UTF-8","HZUID":"3","Authorization":token1}
    data={
        "category":"ALL",
        "orderField":"updatedAt",
        "orderDirection":"DESC",
        "offset":0,
        "limit":"50",
        "items":{"name":costomerName}}
    r = requests.post(url, json=data,headers=headers)
    print(r.json())
    if costomerName==r.json()['items'][0]['name']:
        customerId = r.json()['items'][0]['id']
    elif costomerName==r.json()['items'][1]['name']:
        customerId = r.json()['items'][1]['id']
    elif costomerName==r.json()['items'][2]['name']:
        customerId = r.json()['items'][2]['id']
    else:
        customerId = r.json()['items'][3]['id']






#读取联系人的字段信息




#创建联系人并关联客户
    url = "https://www.weibangong.com/api/contact/create"
    headers = {"Content-Type":"application/json;charset=UTF-8","HZUID":"3","Authorization":tokenN}
    data={"name":contractName,"phone":contractPhone1,"customerId":customerId,"customValueList":{"customValueItemList":[{"id":518844,"type":4,"value":contractTNum1},{"id":518924,"type":1,"value":contractIDate2}]}}
    r = requests.post(url, json=data,headers=headers)
    print(r.json())