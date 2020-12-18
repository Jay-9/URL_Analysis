# -*- coding : utf-8 -*-
# coding: utf-8

import re
import time
import pandas
import datetime
import requests


def judge_ip(the_data):
    match_ip = re.findall(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b', the_data)
    if len(match_ip):
        return 'str_ip'


def url2info(the_url):
    html_info = requests.get('https://icp.aizhan.com/' + the_url + '/').content.decode()
    time.sleep(5)
    name_com = re.findall(r'<td class="thead">主办单位名称</td>[\n].*?</td>', html_info)
    name_web = re.findall(r'<td class="thead">网站名称</td>[\n].*?</td>', html_info)
    if len(name_com) and len(name_web):
        return name_com[0][44:-5], name_web[0][42:-5]
    else:
        return 0


def end_info(the_sor_file):
    the_sor_file = the_sor_file.astype(str)
    for x in range(len(the_sor_file)):
        if ':' in the_sor_file.iat[x, 0]:
            the_sor_file.iat[x, 0] = the_sor_file.iat[x, 0].split(':')[0]
        judge_result = judge_ip(the_sor_file.iat[x, 0])
        if judge_result == 'str_ip':
            the_sor_file.iat[x, 1] = '-'
            the_sor_file.iat[x, 2] = '-'
        else:
            query_info = url2info(the_sor_file.iat[x, 0])
            if query_info:
                the_sor_file.iat[x, 1] = query_info[0]
                the_sor_file.iat[x, 2] = query_info[1]
            else:
                the_sor_file.iat[x, 1] = '未备案'
                the_sor_file.iat[x, 2] = '未备案'
    return the_sor_file


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '分析基础数据.csv'
    sor_file = pandas.read_csv(file_name, encoding="gbk")
    end_file = end_info(sor_file)
    end_file.to_excel(file_name[:-5] + 'OK.xlsx', index=False)
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)
