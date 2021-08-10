# --*-- coding:utf8 --*--
from tools.setting import config
from tools.path import cur_path
from datetime import datetime
import xlrd
import pymysql
import re
import os


def upload_compatibility_testing(filename):
    file = xlrd.open_workbook(filename)
    cases = file.sheet_by_name(file.sheet_names()[0])

    try:
        db = pymysql.connect(**config)

    except Exception as e:
        print("could not connect to mysql server")

    for row in range(1, cases.nrows):
        firstL, secondL, thirdL, versionNo, content = cases.cell(row, 0).value, cases.cell(row, 1).value, \
                                                      cases.cell(row, 2).value, cases.cell(row, 3).value, \
                                                      cases.cell(row, 4).value
        # print((firstL, secondL, thirdL, versionNo, content))
        sql = f"select title from systems where title='{firstL}'"

        with db.cursor() as cursor:
            cursor.execute(sql)
            query = cursor.fetchall()

            if len(query) == 0:
                sql = f"insert into systems (parent_id, title, sort, state) VALUES (0, '{firstL}', 1, 1)"
                cursor.execute(sql)

            sql = f"select id from systems where title='{firstL}'"
            cursor.execute(sql)
            firstId = cursor.fetchone()[0]

            sql = f"select title from systems where title='{secondL}'"
            cursor.execute(sql)
            query = cursor.fetchall()

            if len(query) == 0:
                sql = f"insert into systems (parent_id, title, sort, state) VALUES({firstId}, '{secondL}', 1, 1)"
                cursor.execute(sql)

            sql = f"select id from systems where title='{secondL}'"
            cursor.execute(sql)
            secondId = cursor.fetchone()[0]

            sql = f"select title from systems where title='{thirdL}'"
            cursor.execute(sql)
            query = cursor.fetchall()

            if len(query) == 0:
                sql = f"insert into systems (parent_id, title, sort, state) VALUES('{secondId}', '{thirdL}', 1, 1)"
                cursor.execute(sql)

            sql = f"select id from systems where title='{thirdL}'"
            cursor.execute(sql)
            thirdId = cursor.fetchone()[0]

            sql = f"select id from compatibility_testing where pri_classification={firstId} " \
                  f"and sec_classification={secondId} " \
                  f"and thr_classification={thirdId} " \
                  f"and test_version='{versionNo}' " \
                  f"and require_content='{content}'"

            cursor.execute(sql)
            query = cursor.fetchone()
            if not query:
                cases.cell(row, 5).value = datetime.strptime(str(int(cases.cell(row, 5).value)), '%Y%m%d')
                cases.cell(row, 6).value = datetime.strptime(str(int(cases.cell(row, 6).value)), '%Y%m%d')

                if '%' not in str(cases.cell(row, 8).value):
                    strTemp = str(int(cases.cell(row, 8).value * 100)) + '%'

                # sqls = ",".join([str(cases.cell(row, col).value) if not re.search(r'[A-Za-z]|[\u4e00-\u9fa5]|%|-', str(
                #     cases.cell(row, col).value)) else "'{}'".format(cases.cell(row, col).value) for col in
                #                  range(9, cases.ncols)])

                sql = f"insert into compatibility_testing (id, pri_classification, sec_classification, " \
                      f"thr_classification, test_version, require_content, test_start_date, " \
                      f"test_end_date, consume, pass_rate, test_package_name, auto_cases_num, " \
                      f"excu_allcase_num, test_freq_num, test_freq_content, question_machine, " \
                      f"question_version, install_failed, tart_failed, flash_back, no_response, " \
                      f"accidental_termination, stuck, abnormal_function, ui_exception, bugs_num, " \
                      f"close_bug_num, open_bug_num, open_bug_content, bug_description, bug_analysis, bug_resolvent, " \
                      f"developer, tester, upater, upate_datetime) values(0, {firstId}, {secondId}, {thirdId}," \
                      f"'{cases.cell(row, 3).value}', '{cases.cell(row, 4).value}', '{cases.cell(row, 5).value}', " \
                      f"'{cases.cell(row, 6).value}', {cases.cell(row, 7).value}, '{strTemp}'," \
                      f"'{cases.cell(row, 9).value}',{cases.cell(row, 10).value},{cases.cell(row, 11).value}," \
                      f"{cases.cell(row, 12).value},'{cases.cell(row, 13).value}','{cases.cell(row, 14).value}'," \
                      f"'{cases.cell(row, 15).value}',{cases.cell(row, 16).value},{cases.cell(row, 17).value}," \
                      f"{cases.cell(row, 18).value},{cases.cell(row, 19).value},{cases.cell(row, 20).value}," \
                      f"{cases.cell(row, 21).value},{cases.cell(row, 22).value},{cases.cell(row, 23).value}," \
                      f"{cases.cell(row, 24).value},{cases.cell(row, 25).value},{cases.cell(row, 26).value}," \
                      f"'{cases.cell(row, 27).value}','{cases.cell(row, 28).value}','{cases.cell(row, 29).value}'," \
                      f"'{cases.cell(row, 30).value}','{cases.cell(row, 31).value}','{cases.cell(row, 32).value}'," \
                      f"'{cases.cell(row, 33).value}','{cases.cell(row, 6).value}')"
                cursor.execute(sql)
        db.commit()
    cursor.close()
    db.close()
    print('upload of compatibility test report Successfully！')


if __name__ == '__main__':
    cur_path()
    CURPATH = cur_path()
    filename = os.path.join(CURPATH, "兼容性测试数据统计.xls")
    upload_compatibility_testing(filename)
