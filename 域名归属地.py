# 域名解析后，查cname的归属公司，查IP是否被本网覆盖
import IPy
import datetime
import os
import pandas
import re
import requests
import time


def cname2company(the_url):
    inf = requests.get('https://icp.aizhan.com/' + the_url + '/').text
    company = re.findall(r'<td class="thead">网站名称</td>[\n].*?</td>', inf)
    time.sleep(5)
    if len(company):
        return company[0][42:-5]
    else:
        return '未备案'


def ip_cover(the_ip):
    judge_result = []

    for i in range(1, len(the_ip)):
        if (the_ip[i] in IPy.IP('45.116.208.0/22', make_net=True)) or \
           (the_ip[i] in IPy.IP('103.57.12.0/22', make_net=True)):
            judge_result.append('yes')
        else:
            judge_result.append('no')

    if ('yes' in judge_result) and ('no' in judge_result):
        return '半覆盖'
    else:
        if judge_result[0] == 'yes':
            return '全覆盖'
        elif judge_result[0] == 'no':
            return '未覆盖'


def data_nslookup(the_databases):
    os.system('chcp 65001')
    for row_count in range(len(the_databases)):
        data_val = the_databases.iat[row_count, 0]
        with os.popen(r'nslookup %s 45.116.211.211' % data_val, 'r') as f:
            text = f.read()

            match_name = re.findall(r'Name:    .*', text)
            if len(match_name) == 0:
                match_name = ['空占位符，无意义。无cname']
                the_company = '无法解析'
            else:
                the_company = cname2company(match_name[0][9:])

            match_ip = re.findall(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b', text)
            if len(match_ip) == 1:
                match_ip = '*无IP'
            else:
                the_cover = ip_cover(match_ip)

            the_databases.iat[row_count, 1] = match_name[0][9:] # Cname
            the_databases.iat[row_count, 2] = the_company       # 公司
            the_databases.iat[row_count, 3] = match_ip[1:]      # IP
            the_databases.iat[row_count, 4] = the_cover         # 覆盖情况
    return the_databases


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '基于应用的目标站点排名 (18-21点test.xlsx'
    all_databases = pandas.DataFrame()
    sor_file = pandas.read_excel(file_name)
    for sheet_name, databases in sor_file.items():
        new_databases = data_nslookup(databases)
        # new_databases.to_excel(file_name[:-5] + 'OK.xlsx', index=False)
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)