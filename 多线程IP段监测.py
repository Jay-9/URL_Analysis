# 功能：判断IP是否在IP段范围内，并输出标记（可自动处理空行）---针对带Cname的源数据
import os, IPy, socket, datetime, multiprocessing, pandas   # 18m19s

'''
def linshi():       # 批量定义变量
    names = ['x', 'y', 'z']
    for name in names:
        exec('%s=pandas.DataFrame()' % name)
    print(z)
'''

def empty_domain(web_name):
    web_domain_ips = []
    try:
        ans = socket.getaddrinfo(web_name, None)
        for n in range(len(ans)):
            web_domain_ips.append(ans[n][4][0])
    except socket.gaierror:
        web_domain_ips.append('无法解析')
    return web_domain_ips


def empty_update(databases):
    # empty_index = databases.index[databases[1].isna().values == True][0]  #查找空行的行号
    for empty_index, empty_true in databases[1].isna().items():
        if empty_true:
            web_domain_ips = empty_domain(databases.iat[empty_index, 0])
            databases.iat[empty_index, 1] = web_domain_ips[0]   # 填写第一个解析结果
            for i in range(1, len(web_domain_ips)):
                if databases.shape[1] == 2:    # 没有Cname的填：域名、IP
                    insert_data_frame = pandas.Series([databases.iat[empty_index, 0], web_domain_ips[i]])
                    databases = databases.append(insert_data_frame, ignore_index=True)  # 最下面填写剩余解析结果
                elif databases.shape[1] == 3:  # 有Cname的填：域名、IP、Cname
                    insert_data_frame = pandas.Series([databases.iat[empty_index, 0], web_domain_ips[i],
                                                       databases.iat[empty_index, 2]])  # 域名、IP、Cname
                    databases = databases.append(insert_data_frame, ignore_index=True)  # 最下面填写剩余解析结果
    databases.sort_values(by=0, inplace=True)  # 排序
    databases.drop_duplicates(inplace=True)  # 去重
    databases = databases.reset_index(drop=True)    #重新索引
    return databases


def ip_match(databases, ip_segment):
    databases_columns = databases.shape[1]  # 数据总列数
    databases.insert(databases_columns, databases_columns, value='不匹配')    # 在最后一列插入匹配结果列，初始填“不匹配”(列索引、标签、值)
    for ip_index, ip_value in databases[1].items():
        for ip_segment_index, ip_segment_value in ip_segment[0].items():
            if ip_value != '无法解析' and (ip_value in IPy.IP(ip_segment_value, make_net=True)):
                databases.iat[ip_index, databases_columns] = ip_segment.iat[ip_segment_index, 1]
                break
    return databases


def begin(file_name):
    writer = pandas.ExcelWriter(file_name)
    sor_file = pandas.read_excel(file_name, sheet_name=None, header=None, index_col=None)
    for sheet_name, databases in sor_file.items():
        if sheet_name == 'IP段':
            ip_segment = databases
        else:
            databases = empty_update(databases)
            des_info = ip_match(databases, ip_segment)
            des_info.to_excel(writer, sheet_name=sheet_name, index=False, header=False)
    writer.save()


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    folder_file_path = 'D:\\PyTest\\IPCheck\\IP分拣\\'
    folder_file_names = os.listdir(folder_file_path)
    for folder_file_name in folder_file_names:
        if folder_file_name[-4:] != 'xlsx':
            folder_file_names.remove(folder_file_name)
    po = multiprocessing.Pool()
    for file_name in folder_file_names:
        po.apply_async(begin, args=(folder_file_path + file_name,))
    po.close()
    po.join()
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)
