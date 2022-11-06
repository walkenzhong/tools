import pandas as pd
from pandas import DataFrame
import os
import numpy as np

def read_one_file(file_name):
    data = pd.read_excel(file_name,sheet_name='漏洞信息')
    ip_address = file_name.replace('./data/','')
    ip_address = ip_address.replace('.xls','')
    data['ip地址']=ip_address
    data = data.ffill()
    return data

#漏洞表创建
def vulnerability_table_create(vulnerability_list):
    output_result = DataFrame(columns=['业务系统','IP','系统/主机名称','端口','漏洞名称','漏洞描述','风险等级','整改建议','发现渠道','系统负责人','计划修复日期','漏洞状态','备注'])#可自行添加需要的列名
    output_result['IP'] = vulnerability_list['ip地址']
    output_result['端口'] = vulnerability_list['端口']
    output_result['漏洞名称'] = vulnerability_list['漏洞名称']
    output_result['漏洞描述'] = vulnerability_list['详细描述']
    output_result['风险等级'] = vulnerability_list['风险等级']
    output_result['整改建议'] = vulnerability_list['解决办法']
    output_result.to_excel('漏洞表.xls',index = False)

#资产表创建，该处有个bug，未探测到端口开放的资产不会列入到资产表中
def asset_table_create(vulnerability_list):
    output_result = DataFrame(columns=['业务系统','IP','系统/主机名称','端口','服务','协议','系统负责人','区域'])#可自行添加需要的列名
    output_result['IP'] = vulnerability_list['ip地址']
    output_result['端口'] = vulnerability_list['端口']
    output_result['服务'] = vulnerability_list['服务']
    output_result['协议'] = vulnerability_list['协议']
    output_result = output_result.drop_duplicates(subset=['IP','端口'])
    output_result.to_excel('资产表.xls',index = False)

def main():
    if not os.path.exists('./data'):
        print('不存在data文件夹，请创建data文件夹并将主机报表放置在data文件夹内')
        return 0
    data_file_set = os.listdir('./data') #读取data下的所有文件名
    data_set = []
    for i in range(len(data_file_set)):
        data = read_one_file('./data/'+data_file_set[i])
        data_set.append(data)
    result = pd.concat(data_set)
    vulnerability_table_create(result)
    asset_table_create(result)
    print('success')

if __name__=="__main__":
    main()