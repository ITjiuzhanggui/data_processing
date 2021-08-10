# --*-- conding:utf8 --*--
import os
import pandas as pd
from tools.path import cur_path
from down_automated_testing import down_automated_testing
from down_compatibility_testing import down_compatibility_testing
from down_destructive_testing import down_destructive_testing
from down_function_testing import down_function_testing
from down_performance_testing import down_performance_testing


def down_merge():
    with pd.ExcelWriter(dir + '/合并测试数据统计报告.xls') as writer:
        for i in origin_file_list:
            file_path = dir + '/' + i
            sheet_name = i[:-4]
            df = pd.read_excel(file_path)
            string = "".join(list(str(i) for i in df.index))
            if string.isdigit():
                df.to_excel(writer, sheet_name, index=False)
            else:
                df.to_excel(writer, sheet_name)
        print('Merge complete')


if __name__ == '__main__':
    down_automated_testing()
    down_compatibility_testing()
    down_destructive_testing()
    down_function_testing()
    down_performance_testing()
    cur_path()
    dir = cur_path()
    origin_file_list = os.listdir(dir)
    down_merge()
