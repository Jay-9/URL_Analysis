import pandas
import datetime


def data_clean(the_sor_data):
    the_down_flow = []
    the_up_flow = []
    the_all_flow = []

    for row_num in range(len(the_sor_data)):
        x = the_sor_data.iat[row_num, 4]

        if type(x) is int:
            if x == 0:
                the_down_flow.append('0G')
            else:
                the_down_flow.append(str(x / 1000 / 1000 / 1000) + 'G')

        elif type(x) is str:
            if 'G' in x:
                the_down_flow.append(x)
            elif 'M' in x:
                the_down_flow.append(str(eval(x.split('M')[0]) / 1000) + 'G')
            elif 'K' in x:
                the_down_flow.append(str(eval(x.split('K')[0]) / 1000 / 1000) + 'G')
            else:
                print('str error')

        else:
            the_down_flow.append('空')

    for row_num in range(len(the_sor_data)):
        x = the_sor_data.iat[row_num, 5]

        if type(x) is int:
            if x == 0:
                the_up_flow.append('0G')
            else:
                the_up_flow.append(str(x / 1000 / 1000 / 1000) + 'G')

        elif type(x) is str:
            if 'G' in x:
                the_up_flow.append(x)
            elif 'M' in x:
                the_up_flow.append(str(eval(x.split('M')[0]) / 1000) + 'G')
            elif 'K' in x:
                the_up_flow.append(str(eval(x.split('K')[0]) / 1000 / 1000) + 'G')
            else:
                print('str error')

        else:
            the_up_flow.append('空')

    for row_num in range(len(the_sor_data)):
        x = the_sor_data.iat[row_num, 6]

        if type(x) is int:
            if x == 0:
                the_all_flow.append('0G')
            else:
                the_all_flow.append(str(x / 1000 / 1000 / 1000) + 'G')

        elif type(x) is str:
            if 'G' in x:
                the_all_flow.append(x)
            elif 'M' in x:
                the_all_flow.append(str(eval(x.split('M')[0]) / 1000) + 'G')
            elif 'K' in x:
                the_all_flow.append(str(eval(x.split('K')[0]) / 1000 / 1000) + 'G')
            else:
                print('str error')

        else:
            the_all_flow.append('空')

    the_sor_data['下行流量-新'] = the_down_flow
    the_sor_data['上行流量-新'] = the_up_flow
    the_sor_data['总流量-新'] = the_all_flow

    return the_sor_data


if __name__ == '__main__':
    start_time = datetime.datetime.now()
    file_name = '基于应用的目标站点排名 (10).xlsx'

    sor_data = pandas.read_excel(file_name)
    final_data = data_clean(sor_data)
    # print(final_data)
    final_data.to_csv(file_name[:-5] + 'OK.csv', index=False, encoding='utf_8_sig')
    print(start_time, '\t', datetime.datetime.now(), '\n', datetime.datetime.now() - start_time)