# coding utf-8

import xlrd as xl
import csv


def separate_by_time():
    _file_pos = r'C:\Users\78169\Desktop\b.csv'

    row_list = []
    with open(_file_pos, 'r', newline='') as r:
        _csv_reader = csv.reader(r)
        # start_time = next(_csv_reader)[0]
        # curr_time = start_time
        # start_index = 0
        # end_index = 0
        # for index, content in enumerate(r):
        #     end_index = index
        #     end_time = content[:19]
        #     curr_time = end_time
        #     if cal_time_diff(end_time, curr_time) > (3 * 60):
        #         _new = r'C:\Users\78169\Desktop\{}.csv'.format(csv_name_generator(start_time, end_time))
        #         with open(_new, 'w', newline='') as w:
        #             _csv_write = csv.writer(w)
        #             for i in range(start_index, end_index):
        #                 try:
        #                     _csv_write.writerow(next(_csv_reader))
        #                 except StopIteration:
        #                     pass
        #         w.close()
        while 1:
            try:
                row_list.append(next(_csv_reader))
            except StopIteration:
                break
    start_time = end_time = curr_time = row_list[0][0]
    start_index = end_index = 0
    for index, content in enumerate(row_list):
        end_time = content[0]
        if (cal_time_diff(end_time, curr_time) > (30)) or index == len(row_list) - 1:
            _new = r'C:\Users\78169\Desktop\new\{}.csv'.format(csv_name_generator(start_time, curr_time))
            with open(_new, 'w', newline='') as w:
                _csv_write = csv.writer(w)
                for i in range(start_index, end_index + 1):
                    _csv_write.writerow(row_list[i])
                if index == len(row_list) - 1:
                    _csv_write.writerow(row_list[-1])
            w.close()
            start_time = curr_time = end_time
            start_index = index
        else:
            curr_time = end_time
            end_index = index


def csv_name_generator(time1: str, time2: str):
    return time1.replace(':', '_') + "-" + time2.replace(':', '_')


def cal_time_diff(time1: str, time2: str):
    return (int(time1[:4]) - int(time2[:4])) * 365 * 24 * 60 * 60 + \
           (int(time1[5:7]) - int(time2[5:7])) * 30 * 24 * 60 * 60 + \
           (int(time1[8:10]) - int(time2[8:10]) * 24 * 60 * 60) + \
           (int(time1[11:13]) * 3600 + int(time1[14:16]) * 60 + int(time1[17:])) - \
           (int(time2[11:13]) * 3600 + int(time2[14:16]) * 60 + int(time2[17:]))


if __name__ == '__main__':
    file_pos = r'C:\Users\78169\Desktop\2019_08_15_08_37.xlsx'
    file_to_write = r'C:\Users\78169\Desktop\b.csv'
    with open(file_to_write, 'w', newline='') as f:
        csv_write = csv.writer(f)

        file = xl.open_workbook(file_pos)
        file = file.sheet_by_name('可储能装置电压数据')
        row_num = file.nrows
        col_num = file.ncols
        for i in range(1, row_num):
            if file.cell_value(i, 4) > 3:
                list_to_write = [file.cell_value(i, 0)]
                for j in range(8, col_num):
                    list_to_write.append(int(file.cell_value(i, j) * 1000))
                list_to_write += [25, file.cell_value(i, 4)]
                csv_write.writerow(list_to_write)
    f.close()
    separate_by_time()
