# --*-- coding:utf8 --*--
from tools.setting import config
from tools.path import cur_path
from xlwt import Workbook
import os
import pymysql
import threading


def down_destructive_testing():
    try:
        db = pymysql.connect(**config)

    except Exception as e:
        print("could not connect to mysql server")

    sql = f"select * from destructive_testing;"

    with db.cursor() as cursor:
        cursor.execute(sql)
        query = cursor.fetchall()
    ws = Workbook(encoding='utf-8')

    def saveFile(query):
        if query:
            w = ws.add_sheet(u"破坏性测试数据统计")
            w.write(0, 0, u"一级分类")
            w.write(0, 1, u"二级分类")
            w.write(0, 2, u"三级分类")
            w.write(0, 3, u"版本号")
            w.write(0, 4, u"需求内容")
            w.write(0, 5, u"破坏性测试开始日期")
            w.write(0, 6, u"破坏性测试结束日期")
            w.write(0, 7, u"总场景数")
            w.write(0, 8, u"破坏性测试总轮次")
            w.write(0, 9, u"破坏性测试各轮次详情")
            w.write(0, 10, u"新设计破坏性场景数")
            w.write(0, 11, u"执行破坏性场景数")
            w.write(0, 12, u"致命缺陷数量")
            w.write(0, 13, u"严重缺陷数量")
            w.write(0, 14, u"一般缺陷数量")
            w.write(0, 15, u"提示缺陷数量")
            w.write(0, 16, u"缺陷总数")
            w.write(0, 17, u"已关闭缺陷数量")
            w.write(0, 18, u"未关闭缺陷数量")
            w.write(0, 19, u"未关闭缺陷说明")
            w.write(0, 20, u"缺陷详细描述")
            w.write(0, 21, u"缺陷分析（出现问题的原因）")
            w.write(0, 22, u"缺陷解决描述（缺陷怎么解决的）")
            w.write(0, 23, u"开发人")
            w.write(0, 24, u"测试人")
            w.write(0, 25, u'最后修改人')
            w.write(0, 26, u'最新修改时间')

            excel_row = 1
            for obj in query:
                data_prl = obj[1]
                data_sec = obj[2]
                data_thr = obj[3]

                sql = f"select title from systems where id = {data_prl}"
                with db.cursor() as cursor:
                    cursor.execute(sql)
                    query = cursor.fetchall()
                data_prl = query[0]

                sql = f"select title from systems where id = {data_sec}"
                with db.cursor() as cursor:
                    cursor.execute(sql)
                    query = cursor.fetchall()
                data_sec = query[0]

                sql = f"select title from systems where id = {data_thr}"
                with db.cursor() as cursor:
                    cursor.execute(sql)
                    query = cursor.fetchall()
                data_thr = query[0]

                test_version = obj[4]
                require_content = obj[5]
                test_start_date = obj[6]
                test_end_date = obj[7]
                auto_cases_num = obj[8]
                test_freq_num = obj[9]
                test_freq_content = obj[10]
                add_autocase_num = obj[11]
                excu_allcase_num = obj[12]
                blocker_bug_num = obj[13]
                critical_bug_num = obj[14]
                major_bug_num = obj[15]
                minor_bug_num = obj[16]
                bugs_num = obj[17]
                close_bug_num = obj[18]
                open_bug_num = obj[19]
                open_bug_content = obj[20]
                bug_description = obj[21]
                bug_analysis = obj[22]
                bug_resolvent = obj[23]
                developer = obj[24]
                tester = obj[25]
                updater = obj[26]
                update_datetime = obj[27]

                w.write(excel_row, 0, data_prl)
                w.write(excel_row, 1, data_sec)
                w.write(excel_row, 2, data_thr)
                w.write(excel_row, 3, test_version)
                w.write(excel_row, 4, require_content)
                w.write(excel_row, 5, test_start_date.strftime('%Y%m%d'))
                w.write(excel_row, 6, test_end_date.strftime('%Y%m%d'))
                w.write(excel_row, 7, auto_cases_num)
                w.write(excel_row, 8, test_freq_num)
                w.write(excel_row, 9, test_freq_content)
                w.write(excel_row, 10, add_autocase_num)
                w.write(excel_row, 11, excu_allcase_num)
                w.write(excel_row, 12, blocker_bug_num)
                w.write(excel_row, 13, critical_bug_num)
                w.write(excel_row, 14, major_bug_num)
                w.write(excel_row, 15, minor_bug_num)
                w.write(excel_row, 16, bugs_num)
                w.write(excel_row, 17, close_bug_num)
                w.write(excel_row, 18, open_bug_num)
                w.write(excel_row, 19, open_bug_content)
                w.write(excel_row, 20, bug_description)
                w.write(excel_row, 21, bug_analysis)
                w.write(excel_row, 22, bug_resolvent)
                w.write(excel_row, 23, developer)
                w.write(excel_row, 24, tester)
                w.write(excel_row, 25, updater)
                w.write(excel_row, 26, update_datetime.strftime('%Y-%m-%d %H:%M:%S'))
                excel_row += 1

    thread = threading.Thread(target=saveFile, args=(query,))
    thread.start()
    thread.join()
    cur_path()
    CURPATH = cur_path()
    ex_path = os.path.join(CURPATH, "破坏性测试数据统计.xls")
    ws.save(ex_path)
    cursor.close()
    db.close()
    print('Download of destructive test report Successfully！')


if __name__ == '__main__':
    down_destructive_testing()
