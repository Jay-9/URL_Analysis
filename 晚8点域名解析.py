import pandas
import datetime
import os
import re


def nslookup(the_sor_data):
    ip_list = []
    os.system('chcp 65001')

    for data in the_sor_data['域名'].values:
        with os.popen(r'nslookup %s 45.116.211.211' % data, 'r') as f:
            text = f.read()

            match_ip = re.findall(r'\b(?:[0-9]{1,3}\.){3}[0-9]{1,3}\b', text)
            if len(match_ip) == 1:
                ip_list.append('无IP')
            else:
                ip_list.append(match_ip[1:])

    the_sor_data['解析IP'] = ip_list
    return the_sor_data


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '晚8点域名解析.xlsx'

    sor_data = pandas.read_excel(file_name)
    final_data = nslookup(sor_data)

    final_data.to_excel(file_name[:-5] + 'OK.xlsx', index=False, header=False)
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)
