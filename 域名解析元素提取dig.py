import datetime
import os
import pandas
import re







if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '基于应用的目标站点排名dig25.xlsx'
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