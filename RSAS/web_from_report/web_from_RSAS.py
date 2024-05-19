import sys
import os
import zipfile
from bs4 import BeautifulSoup

def undepress(path):
    # save_path=os.path.dirname(path) #目录路径
    try:
        file=zipfile.ZipFile(path)
        save_path=path[:-4]
        # 存在已解压文件夹不会重复解压
        file.extractall(save_path)
        print('解压完成')
        file.close()
    except FileNotFoundError:
        print("压缩包路径错误")
        sys.exit()
    return save_path

# 返回存在web服务的IP
def getIP(path):
    # 50188、50856
    ips=[]
    http_soup=""
    https_soup=""
    # 读取html文件内容，某些扫描任务可能没有http服务，目标文件不存在，需加判断
    try:
        https_raw=open(os.path.join(path,r'vulnHtml/50856.html'),'r',encoding='utf-8').read()
        https_soup = BeautifulSoup(https_raw, 'lxml')
    except FileNotFoundError:
        pass
    try:
        http_raw = open(os.path.join(path, r'vulnHtml/50188.html'), 'r', encoding='utf-8').read()
        http_soup = BeautifulSoup(http_raw, 'lxml')
    except FileNotFoundError:
        pass

    # 获取所有a标签
    if http_soup:
        http_links = http_soup.find_all("a")
        for link in http_links:
            ips.append(link.text)
    if https_soup:
        https_links = https_soup.find_all("a")
        # 提取
        for link in https_links:
            ips.append(link.text)

    print('提取WEB服务IP完成')
    return ips



# http://ip:port格式返回，
def getUrls(path,iplist):
    res=[]
    _path=os.path.join(path,r'host')
    for ip in iplist:
        cur_path=os.path.join(_path,ip+'.html')
        ip_content=open(cur_path,'r',encoding='utf-8').read()
        soup=BeautifulSoup(ip_content,'lxml')
        # 主机表格
        vuln_table=soup.find(id='vuln_list')
        # 遍历所有行
        vul_rows=vuln_table.find_all('tr')
        for i in range(2,len(vul_rows)):
            url=''
            tds= vul_rows[i].find_all('td')
            service=tds[2].text
            if service =='https':
                url=f'https://{ip}:{tds[0].text}'
            elif service =='http':
                url=f'http://{ip}:{tds[0].text}'
            if url:
                res.append(url)
        # vul_rows=vuln_table.find_all('tr')
        # print(vul_rows[5].find_all('td')[2].text)
    print('构造URL地址完成')
    return res
def save2txt(path,urls):
    with open(os.path.join(path,'url.txt'),'w+',encoding='utf-8')as file:
        for u in urls:
            file.write(u+'\n')
        print(f'结果已保存至{os.path.join(path,'url.txt')}')
        file.close()
if __name__ == '__main__':
    # 文件绝对路径
    path = sys.argv[1]
    # 解压后的文件路径
    save_path=undepress(path)
    # 存在web服务的ip
    ip_list=getIP(save_path)
    # 获取url
    urls=[]
    if ip_list:
        urls=getUrls(save_path,ip_list)
    if urls:
    # 结果保存到压缩包所在目录
        save2txt(os.path.dirname(save_path),urls)