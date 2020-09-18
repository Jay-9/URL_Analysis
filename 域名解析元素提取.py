'''
源数据中'站点名称_域名'的列删除端口号后，筛选域名及IP地址
1、域名做nslookup解析，提取Cname（有排序）和及解析IP（其中Cname再提取Cname_first和Cname_last）
2、域名做nslookup -q=ns解析，提取primary name，并对primary name进行nslookup解析ip
目标地址、目标AS域、下行流量、上行流量、总流量的列保留源数据
输出excel包含’网址‘和’IP‘两个sheet表
（50行数据约27秒）
'''

import datetime
import os
import pandas
import re


def data_clean(the_sor_data):
    ip_row = []
    the_ip_data = pandas.DataFrame()
    compile_ip = re.compile('^(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|[1-9])\.(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|\d)\.(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|\d)\.(1\d{2}|2[0-4]\d|25[0-5]|[1-9]\d|\d)$')

    for row_num in range(len(the_sor_data)):
        if ':' in the_sor_data.iat[row_num, 0]:
            the_sor_data.iat[row_num, 0] = the_sor_data.iat[row_num, 0].split(':')[0]

    for row_num in range(len(the_sor_data)):
        if compile_ip.match(the_sor_data.iat[row_num, 0]):
            the_ip_data = the_ip_data.append(the_sor_data.iloc[row_num, :], ignore_index=True)
            ip_row.append(row_num)

    the_sor_data.drop(index=ip_row, inplace=True)

    return the_sor_data, the_ip_data


def ns_lookup(the_clean_url_data):
    the_clean_url_data = the_clean_url_data.astype(str)
    os.system('chcp 65001')

    for one_url in the_clean_url_data['站点名称_域名']:
        one_url_row = the_clean_url_data[(the_clean_url_data.站点名称_域名 == one_url)].index.tolist()[0]
        with os.popen(r'nslookup %s 45.116.211.211' % one_url, 'r') as f:
            text = f.read()

            match_ip = re.findall(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b', text)
            if len(match_ip) == 1:
                the_clean_url_data.at[one_url_row, '解析IP'] = '无解析IP'
            else:
                the_clean_url_data.at[one_url_row, '解析IP'] = match_ip[1:]

            match_aliases = re.findall(r'Aliases: [\s\S]*', text)
            match_name = re.findall(r'Name: .*', text)
            if len(match_name) == 0:
                the_clean_url_data.at[one_url_row, '解析Cname'] = '无解析Cname'
            else:
                match_name[0] = match_name[0][9:]
                if len(match_aliases) != 0:
                    match_aliases[0] = match_aliases[0].expandtabs()  # 替换\t
                    match_aliases[0] = match_aliases[0].replace('          ', '')
                    match_aliases[0] = match_aliases[0].replace('\n\n', '\n')
                    match_aliases[0] = match_aliases[0].split('\n')
                    # match_aliases[0][0] = match_aliases[0][0][10:]
                    match_aliases[0] = match_aliases[0][1:]
                    match_aliases[0] = [x for x in match_aliases[0] if x != '']
                    the_clean_url_data.at[one_url_row, '解析Cname'] = match_aliases[0] + match_name
                else:
                    the_clean_url_data.at[one_url_row, '解析Cname'] = match_name

        final_cname = the_clean_url_data.at[one_url_row, '解析Cname']
        if len(final_cname) == 1:
            if final_cname[0][:-1] == the_clean_url_data.at[one_url_row, '站点名称_域名']:
                the_clean_url_data.at[one_url_row, '解析Cname'] = '无解析Cname'
            else:
                the_clean_url_data.at[one_url_row, '解析Cname_first'] = the_clean_url_data.at[one_url_row, '解析Cname']
        else:
            the_clean_url_data.at[one_url_row, '解析Cname_first'] = final_cname[0]
            the_clean_url_data.at[one_url_row, '解析Cname_last'] = final_cname[-1]

    return the_clean_url_data


def ns_primary_name(the_match_name):
    with os.popen(r'nslookup %s 45.116.211.211' % the_match_name, 'r') as f:
        text = f.read()
        match_ip = re.findall(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b', text)
        if len(match_ip) == 1:
            the_primary_name_ip = '无解析IP'
        else:
            the_primary_name_ip = match_ip[1:]
    return the_primary_name_ip


def ns_lookup_q(the_ns_lookup_web_data):
    os.system('chcp 65001')
    for one_url in the_ns_lookup_web_data['站点名称_域名']:
        one_url_row = the_ns_lookup_web_data[(the_ns_lookup_web_data.站点名称_域名 == one_url)].index.tolist()[0]
        with os.popen(r'nslookup -q=ns %s 45.116.211.211' % one_url, 'r') as f:
            text = f.read()

            match_name = re.findall(r'primary name server = .*', text)
            if len(match_name) == 0:
                the_ns_lookup_web_data.at[one_url_row, 'Primary Name（-q=ns）'] = '无primary name'
            else:
                the_ns_lookup_web_data.at[one_url_row, 'Primary Name（-q=ns）'] = match_name[0][22:]
                primary_name_ip = ns_primary_name(match_name[0][22:])
                the_ns_lookup_web_data.at[one_url_row, 'Primary Name解析IP'] = primary_name_ip

    return the_ns_lookup_web_data


def format_data(the_one_value):
    return the_one_value.replace("[", "").replace("]", "").replace("'", "").replace(",", "\n").replace(" ", "").replace("nan", "-")


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '基于应用的目标站点排名50.xlsx'
    sor_data = pandas.read_excel(file_name)

    clean_url_data, final_ip_data = data_clean(sor_data)
    ns_lookup_web_data = ns_lookup(clean_url_data)
    final_web_data = ns_lookup_q(ns_lookup_web_data)

    final_web_data = final_web_data.astype(str)
    format_final_web_data = final_web_data.applymap(format_data)

    writer = pandas.ExcelWriter(file_name[:-5] + '_OK.xlsx')
    columns = ['站点名称_域名', '解析Cname', '解析Cname_first', '解析Cname_last', '解析IP', 'Primary Name（-q=ns）', 'Primary Name解析IP', '（空列）', '目标地址', '目标AS域', '下行流量', '上行流量', '总流量']
    format_final_web_data.to_excel(excel_writer=writer, sheet_name='URL', index=False, columns=columns)
    final_ip_data.to_excel(excel_writer=writer, sheet_name='IP', index=False, columns=columns)
    writer.save()
    writer.close()
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)
