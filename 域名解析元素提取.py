'''
源数据中'站点名称'删除端口号后，筛选域名及IP地址
域名做nslookup -q=ns解析，提取primary name
域名做nslookup解析，提取Cname（有排序）和及IP地址
输出excel包含’网址‘和’IP‘两个sheet表
'''


import datetime
import os
import pandas
import re


def data_clean(the_sor_data):
    the_ip_data = pandas.DataFrame()
    the_web_data = pandas.DataFrame()
    compile_ip = re.compile('^(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|[1-9])\.(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|\d)\.(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|\d)\.(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|\d)$')

    for row_num in range(len(the_sor_data)):
        if ':' in the_sor_data.iat[row_num, 0]:
            the_sor_data.iat[row_num, 0] = the_sor_data.iat[row_num, 0].split(':')[0]

    for row_num in range(len(the_sor_data)):
        if compile_ip.match(the_sor_data.iat[row_num, 0]):
            the_ip_data = the_ip_data.append(the_sor_data.iloc[row_num, :], ignore_index=True)
        else:
            the_web_data = the_web_data.append(the_sor_data.iloc[row_num, :], ignore_index=True)
    return the_web_data, the_ip_data


def nslookup_q(the_web_data):
    primary_name_list = []
    os.system('chcp 65001')
    for data in the_web_data['站点名称'].values:
        with os.popen(r'nslookup -q=ns %s 45.116.211.211' % data, 'r') as f:
            text = f.read()

            match_name = re.findall(r'primary name server = .*', text)
            if len(match_name) == 0:
                primary_name_list.append('无primary name')
            else:
                primary_name_list.append(match_name[0][22:])
    the_web_data['Primary Name（-q=ns）'] = primary_name_list
    return the_web_data


def nslookup(the_primary_web_data):
    ip_list = []
    aliases_list = []
    os.system('chcp 65001')

    for data in the_primary_web_data['站点名称'].values:
        with os.popen(r'nslookup %s 45.116.211.211' % data, 'r') as f:
            text = f.read()

            match_ip = re.findall(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b', text)
            if len(match_ip) == 1:
                ip_list.append('无IP')
            else:
                ip_list.append(match_ip[1:])

            match_aliases = re.findall(r'Aliases: [\s\S]*', text)
            match_name = re.findall(r'Name: .*', text)

            if len(match_name) == 0:
                aliases_list.append('无Cname')
            else:
                if len(match_aliases) != 0:
                    match_aliases[0] = match_aliases[0].expandtabs()  # 替换\t
                    match_aliases[0] = match_aliases[0].replace('          ', '')
                    match_aliases[0] = match_aliases[0].replace('\n\n', '\n')
                    aliases_list.append(match_aliases[0][10:] + match_name[0][9:])
                    if '\n' in aliases_list[0]:
                        aliases_list[0] = aliases_list[0].split('\n')
                else:
                    aliases_list.append(match_name[0][9:])

    the_primary_web_data['解析IP'] = ip_list
    the_primary_web_data['解析Cname'] = aliases_list
    return the_primary_web_data


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '基于应用的目标站点排名 (18-21点).xlsx'

    sor_data = pandas.read_excel(file_name)
    web_data, final_ip_data = data_clean(sor_data)

    primary_web_data = nslookup_q(web_data)
    final_web_data = nslookup(primary_web_data)

    writer = pandas.ExcelWriter(file_name[:-5] + 'OK.xlsx')
    final_web_data.to_excel(excel_writer=writer, sheet_name='网址', index=False)
    final_ip_data.to_excel(excel_writer=writer, sheet_name='IP', index=False)
    writer.save()
    writer.close()
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)
