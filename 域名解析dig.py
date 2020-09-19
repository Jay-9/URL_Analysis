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


def dig_url(the_clean_url_data):
    the_clean_url_data = the_clean_url_data.astype(str)
    os.chdir('C:/d')

    for one_url in the_clean_url_data['站点名称_域名']:
        one_url_row = the_clean_url_data[(the_clean_url_data.站点名称_域名 == one_url)].index.tolist()[0]
        with os.popen(r'dig @211.141.0.99 %s' % one_url, 'r') as f:
            all_text = f.read()

            match_answer_section = re.findall(r'ANSWER SECTION:[\s\S]*', all_text)
            if len(match_answer_section) != 0:
                answer_section = match_answer_section[0][:match_answer_section[0].find(';;') - 1]
                CNAME = re.findall(r'CNAME\t.*', answer_section)

                if len(CNAME) == 0:
                    the_clean_url_data.at[one_url_row, '解析Cname'] = '无解析Cname'
                elif len(CNAME) == 1:
                    the_clean_url_data.at[one_url_row, '解析Cname'] = CNAME[0][6:]
                elif len(CNAME) == 2:
                    the_clean_url_data.at[one_url_row, '解析Cname'] = CNAME[0][6:] + '\n' + CNAME[1][6:]
                    the_clean_url_data.at[one_url_row, '解析Cname_first'] = CNAME[0][6:]
                    the_clean_url_data.at[one_url_row, '解析Cname_last'] = CNAME[1][6:]
                else:
                    CNAME_value = ''
                    for one_CNAME in CNAME:
                        CNAME_value = CNAME_value + one_CNAME[6:] + '\n'
                    the_clean_url_data.at[one_url_row, '解析Cname'] = CNAME_value
                    the_clean_url_data.at[one_url_row, '解析Cname_first'] = CNAME[0][6:]
                    the_clean_url_data.at[one_url_row, '解析Cname_last'] = CNAME[-1][6:]

                A = re.findall(r'A\t.*', answer_section)
                if len(A) == 0:
                    the_clean_url_data.at[one_url_row, '解析IP'] = '无解析IP'
                else:
                    A_ip = ''
                    for one_ip in A:
                        A_ip = A_ip + one_ip[2:] + '\n'
                    the_clean_url_data.at[one_url_row, '解析IP'] = A_ip

            match_additional_section = re.findall(r'ADDITIONAL SECTION:[\s\S]*', all_text)
            if len(match_additional_section) != 0:
                additional_section = match_additional_section[0][:match_additional_section[0].find(';;') - 1]
                primary_name = re.findall(r'.*(?=\t\d*\tIN)', additional_section)
                primary_name = ['\n' if i == '' else i for i in primary_name]
                primary_value = ''
                for one_name in primary_name:
                    primary_value = primary_value + one_name
                the_clean_url_data.at[one_url_row, 'Primary Name（-q=ns）'] = primary_value

                primary_ip = re.findall(r'A\t.*', additional_section)
                ip_value = ''
                for one_ip in primary_ip:
                    ip_value = ip_value + one_ip[2:] + '\n'
                the_clean_url_data.at[one_url_row, 'Primary Name解析IP'] = ip_value

    return the_clean_url_data


def format_data(the_one_value):
    return the_one_value.replace("nan", "-")


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '基于应用的目标站点排名.xlsx'
    sor_data = pandas.read_excel(file_name)

    clean_url_data, final_ip_data = data_clean(sor_data)
    final_dataframe = dig_url(clean_url_data)
    format_final_dataframe = final_dataframe.applymap(format_data)

    os.chdir(os.path.abspath(os.path.dirname(__file__)))
    writer = pandas.ExcelWriter(file_name[:-5] + '_OK.xlsx')
    columns = ['站点名称_域名', '解析Cname', '解析Cname_first', '解析Cname_last', '解析IP', 'Primary Name（-q=ns）', 'Primary Name解析IP', '（空列）', '目标地址', '目标AS域', '下行流量', '上行流量', '总流量']
    format_final_dataframe.to_excel(excel_writer=writer, sheet_name='URL', index=False, columns=columns)
    final_ip_data.to_excel(excel_writer=writer, sheet_name='IP', index=False, columns=columns)
    writer.save()
    writer.close()
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)
