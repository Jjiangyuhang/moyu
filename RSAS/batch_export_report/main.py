# -*- coding: utf-8 -*-
import json
import os.path
import time
import requests, urllib3, re
import webtask_from_xls
from lxml import etree

import exportJSON
import getConfig
import zipfile


def login():
    r1 = requests.get(baseURL + 'accounts/login/', verify=False)
    xmlhtml = etree.HTML(r1.text)
    csrfmiddlewaretoken = xmlhtml.xpath('/html/body/div[2]/div/form/fieldset/input[4]/@value')[0]
    sessionid = re.findall(r'sessionid=(.+?);', r1.headers.get('Set-Cookie'))[0]
    csrftoken = re.findall(r'csrftoken=(.+?);', r1.headers.get('Set-Cookie'))[0]
    headers = {
        'Host': baseIP,
        'Cookie': 'sessionid=' + sessionid + '; csrftoken=' + csrftoken,
        'Cache-Control': 'max-age=0',
        'Sec-Ch-Ua': '" Not A;Brand";v="99", "Chromium";v="102", "Google Chrome";v="102"',
        'Sec-Ch-Ua-Mobile': '?0',
        'Sec-Ch-Ua-Platform': '"Windows"',
        'Upgrade-Insecure-Requests': '1',
        'Origin': baseURL,
        'Content-Type': 'application/x-www-form-urlencoded',
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/102.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
        'Sec-Fetch-Site': 'same-origin',
        'Sec-Fetch-Mode': 'navigate',
        'Sec-Fetch-User': '?1',
        'Sec-Fetch-Dest': 'document',
        'Referer': baseURL + 'accounts/login/',
        'Accept-Encoding': 'gzip, deflate',
        'Accept-Language': 'zh-CN,zh;q=0.9,en;q=0.8,und;q=0.7'
    }
    data = 'username=' + username + '&password=' + password + '&csrfmiddlewaretoken=' + csrfmiddlewaretoken
    r2 = session.post(baseURL + 'accounts/login_view/', headers=headers, data=data, verify=False)
    if '新建任务' in r2.text:
        return {
            'status': True,
            'data': {
                'csrftoken': requests.utils.dict_from_cookiejar(session.cookies)['csrftoken'],
                'sessionid': requests.utils.dict_from_cookiejar(session.cookies)['sessionid']
            }
        }
    else:
        return {
            'status': False,
            'data': '登录失败！'
        }


def createTask(reportName, export_ip):
    header = {
        "Cookie": f"csrftoken= {csrftoken}; sessionid={sessionid}; left_menustatue_NSFOCUSRSAS=0|2|{baseURL}/report/",
        "Sec-Ch-Ua": "\"Not.A/Brand\";v=\"8\", \"Chromium\";v=\"114\", \"Google Chrome\";v=\"114\"",
        "Accept": "*/*",
        "Content-Type": "application/x-www-form-urlencoded",
        "X-Requested-With": "XMLHttpRequest",
        "Sec-Ch-Ua-Mobile": "?0",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
        "Sec-Ch-Ua-Platform": "\"Windows\"",
        "Origin": baseURL,
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
        "Referer": baseURL + "report/",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,und;q=0.7"

    }
    r1 = session.get(baseURL + 'report/', verify=False,timeout=5)
    xmlhtml = etree.HTML(r1.text)
    csrfmiddlewaretoken = xmlhtml.xpath('//input[@name="csrfmiddlewaretoken"]/@value')[0]
    lists = ''
    # task_id=[str(num)for num in range(int(taskNum.split('-')[0]),int(taskNum.split('-')[1])+1)if(num!=74 and num!=75)]
    task_id='127,86'
    for i in export_ip:
        lists += 'filter_host_list=' + i + '&'
    data = """export_area=sys&filter_host_flag=yes&""" + lists + """report_type=html&report_type=doc&report_type=xls&report_content=summary&summary_template_id=1&summary_report_title=绿盟科技"远程安全评估系统"安全评估报告&report_content=host&host_template_id=101&single_report_title=绿盟科技"远程安全评估系统"安全评估报告-主机报表&multi_export_type=multi_sum&multi_report_name=""" + reportName + """&csrfmiddlewaretoken=""" + csrfmiddlewaretoken + """&from=report_export&task_id=""" + task_id
    r2 = requests.post(baseURL + 'report/export', headers=header, data=data.encode(), verify=False,proxies=proxy)
    res = json.loads(re.findall(r'[(](.*?)[)]', r2.text)[0])
    if res['result'] == 'error':
        print(reportName, '导出任务创建失败，没有相关的任务信息！', 'id:' + str(res['context']['report_id']))
    else:
        task_info[reportName] = {
            'report_id': res['context']['report_id'],
            'export_ip': export_ip,
            # 是否创建成功
            'isCreate': False,
            # 是否下载
            'isDone': False
        }
        print(reportName, '导出任务创建成功！', 'id:' + str(res['context']['report_id']))


def createWebTask(reportName, export_ip):
    header = {
        "Cookie": f"csrftoken= {csrftoken}; sessionid={sessionid}; left_menustatue_NSFOCUSRSAS=0|2|{baseURL}/report/",
        "Sec-Ch-Ua": "\"Not.A/Brand\";v=\"8\", \"Chromium\";v=\"114\", \"Google Chrome\";v=\"114\"",
        "Accept": "*/*",
        "Content-Type": "application/x-www-form-urlencoded",
        "X-Requested-With": "XMLHttpRequest",
        "Sec-Ch-Ua-Mobile": "?0",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/114.0.0.0 Safari/537.36",
        "Sec-Ch-Ua-Platform": "\"Windows\"",
        "Origin": baseURL,
        "Sec-Fetch-Site": "same-origin",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Dest": "empty",
        "Referer": baseURL + "report/",
        "Accept-Encoding": "gzip, deflate",
        "Accept-Language": "zh-CN,zh;q=0.9,en;q=0.8,und;q=0.7"

    }
    r1 = session.get(baseURL + 'report/', verify=False,timeout=5)
    xmlhtml = etree.HTML(r1.text)
    csrfmiddlewaretoken = xmlhtml.xpath('//input[@name="csrfmiddlewaretoken"]/@value')[0]
    lists = ''
    task_id=''
    for i in export_ip:
        lists += 'filter_site_list=' + i + '&'
    data = """export_area=web&filter_host_flag=yes&""" + lists + """report_type=html&report_type=doc&report_type=xls&report_content=websummary&summary_template_id=201&summary_report_title=绿盟科技"远程安全评估系统"安全评估报告&report_content=site&host_template_id=301&single_report_title=绿盟科技"远程安全评估系统"安全评估报告-主机报表&multi_export_type=multi_sum&multi_report_name=""" + reportName + """&csrfmiddlewaretoken=""" + csrfmiddlewaretoken + """&from=report_export&task_id=""" + task_id
    r2 = requests.post(baseURL + 'report/export', headers=header, data=data.encode(), verify=False,proxies=proxy)
    res = json.loads(re.findall(r'[(](.*?)[)]', r2.text)[0])
    if res['result'] == 'error':
        print(reportName, '导出任务创建失败，没有相关的任务信息！', 'id:' + str(res['context']['report_id']))
    else:
        task_info[reportName] = {
            'report_id': res['context']['report_id'],
            'export_ip': export_ip,
            # 是否创建成功
            'isCreate': False,
            # 是否下载
            'isDone': False
        }
        print(reportName, '导出任务创建成功！', 'id:' + str(res['context']['report_id']))

def dowfile(id, name):
    types = ['html','doc', 'xls']
    for t in types:

        url = 'https://' + baseIP + '/report/download/id/' + str(id) + '/type/'+t+'/'
        try:
            print(str(id) + '-' + name, '等待下载中！')

            while True:
                r = requests.get(url, timeout=30, verify=False, headers={
                    'Cookie': 'csrftoken=' + csrftoken + '; sessionid=' + sessionid + '; left_menustatue_NSFOCUSRSAS=0|1|https://' + baseIP + '/list/'}
                                 )
                if 'Not Found!' not in r.text:
                    Content_Disposition = r.headers['Content-Disposition']
                    compiler = re.compile(r'filename=(.*)')
                    filename = compiler.search(Content_Disposition).group(1).split("\"")[1]  # 正则提取文件名
                    dowfilename = filename.encode('iso-8859-1').decode('gbk')

                    with open(os.path.join(down_path, dowfilename), "wb") as code:
                        code.write(r.content)
                        # print(str(id) + '-' + name, '下载成功！')
                    break

        except Exception as e:
            fail_lists.append(name)
            print(e, '下载失败！')


def check_zip_files(directory):
    for filename in os.listdir(directory):
        filepath = os.path.join(directory, filename)
        if zipfile.is_zipfile(filepath):
            try:
                with zipfile.ZipFile(filepath) as zf:
                    zf.testzip()
            except zipfile.BadZipFile:
                print(f"{filename} 损坏")
        else:
            print(f"{filename} 不是zip文件")


def isCreateSuccess(task_info):
    for task in task_info:
        id=task['report_id']
        url = f'{baseURL}report/export/process_info/id/{id}/'
        header = {
            'Cookies': f'csrftoken={csrftoken};sessionid={sessionid}'
        }
        raw = session.get(url, verify=False, headers=header)
        res = json.loads(re.findall(r'[(](.*?)[)]', raw.text)[0])

        progress = res['context']['progress']
        if progress == 100:
            task['isCreate']=True
        else:
            print(task['report_id'],'创建失败')


if __name__ == "__main__":
    # 自签ssl防止报错
    urllib3.disable_warnings()
    # 初始化session
    session = requests.session()
    # 设置代理
    proxy = {
        # 'http': '127.0.0.1:8080',
        # 'https': '127.0.0.1:8080'
    }
    # 获取配置文件信息
    baseURL = getConfig.baseURL.decode()
    baseIP = getConfig.baseIP.decode()
    username = getConfig.username.decode()
    password = getConfig.password.decode()

    csrftoken = ""
    sessionid = ""

    fail_lists = []
    # 任务信息：名称、ip、是否创建成功、是否下载成功
    task_info = {}
    cur_dir = 'test'
    # 扫描任务编号
    taskNum = '68-82'
    # 下载路径
    down_path = os.path.join(cur_dir, 'web/')
    # 持久化存储文件路径
    task_info_path = os.path.join(cur_dir, 'task_info.json')
    # 登录
    while True:
        login = login()
        if login['status'] == True:
            csrftoken = login['data']['csrftoken']
            sessionid = login['data']['sessionid']
            break

    # 存在下载任务信息文件，则无需重复生成
    if os.path.exists(task_info_path):
        with open(task_info_path, 'r') as f:
            task_info = json.loads(f.read())
    else:
        print('Excel提取json数据中！')
        # tasklist = exportJSON.get_tasklist()
        tasklist=webtask_from_xls.getTaskList()

        print('Excel提取json数据成功，开始创建任务！')

        for task in tasklist:
            # 讲道理可以同时创建web和主机扫描导出任务，没试过
            # if tasklist[task]['export_ip'] :
            #     createTask(task, tasklist[task]['export_ip'])
            if tasklist[task]['site']:
                createWebTask(task,tasklist[task]['site'])
        print("创建任务请求发送完成，等待任务创建中")
        time.sleep(60 * 1)

        #判断是否创建成功
        # isCreateSuccess(task_info)

        # 将任务列表持久化存储，避免多次运行重复生成报告列表
        with open(task_info_path, 'w') as f:
            f.write(json.dumps(task_info))

    for task in task_info:
        # if(task['isCreate']):
        dowfile(task_info[task]['report_id'], task)

    print("*" * 20)
    print(fail_lists)