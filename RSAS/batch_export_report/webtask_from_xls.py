import tldextract
import pandas as pd
import json

# 提取根域
def getRootDomain(url):
    extracted = tldextract.extract(url)
    root_domain = extracted.registered_domain
    return root_domain


def getCompanyInfo():
    """
    从文件中提取资产：IP、domain
    Returns:
        每个企业的资产，字典格式如下
        {'单位': {'export_ip': ['ip1', 'ip2'], 'domains': ['domain1', 'domain2']}
    """
    xls = pd.ExcelFile(r'企业信息文件')
    res = {}
    # 遍历所有表
    for sheet_name in xls.sheet_names:
        # 读取每个表格的数据
        df = pd.read_excel(xls, sheet_name)

        # 统计每个企业的IP和domain
        company_stats = df.groupby('企业名称').agg({'IP': 'unique', 'domain': 'unique'}).reset_index()

        # 统计结果
        for index, row in company_stats.iterrows():
            #表名-企业
            name = sheet_name + '-' + row['企业名称']
            # 提取不为空的域名
            domains = [domain for domain in row['domain'] if not pd.isnull(domain)]
            res[name] = {
                'export_ip': [ip for ip in row['IP'] if not pd.isnull(ip)],
                # 提取根域名后去重
                'domains': list(set(map(getRootDomain, domains)))
            }
    return res


def getSiteList():
    """
    web任务中每个域名对应一个site编号。梳理根域名对应的所有site编号
    Returns:
    根域名对应的所有site编号，字典格式
    """
    xls = pd.ExcelFile(r'域名与web任务site编号')
    domainfilter_site_list = {}
    for sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name)

        # 根据domain分组，并统计 domain
        # 注:此处域名site对应表中，我手动梳理了domain列（用getRootDomain()跑后手动加到表里的）,懒得写了
        domain_grouped = df.groupby('domain')['filter_site_list'].agg(lambda x: list(set(x.dropna())))

        for domain, filter_site_list in domain_grouped.items():
            domainfilter_site_list[domain] = list(filter_site_list)

    return domainfilter_site_list


def getTaskList():
    """
    企业资产中domain对应的所有site信息
    Returns:

    """
    companyDomain = getCompanyInfo()
    site_list = getSiteList()
    taskList = {}

    for com, data in companyDomain.items():
        site = []
        for domain in data['domains']:
            if domain in site_list:
                site += site_list[domain]
        taskList[com] = {
            'domains': data['domains'],
            'export_ip': data['export_ip'],
            'site': site
        }
        # break
    return taskList

