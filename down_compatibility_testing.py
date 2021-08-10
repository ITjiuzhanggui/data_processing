# --*-- coding:utf8 --*--
from tools.setting import config
from tools.path import cur_path
from xlwt import Workbook
import os
import pymysql
import threading


def down_compatibility_testing():
    try:
        db = pymysql.connect(**config)

    except Exception as e:
        print("could not connect to mysql server")

    sql = f"select * from compatibility_testing;"

    with db.cursor() as cursor:
        cursor.execute(sql)
        query = cursor.fetchall()
    ws = Workbook(encoding='utf-8')

    def saveFile(query):
        if query:
            w = ws.add_sheet(u"兼容性测试数据统计")
            w.write(0, 0, u"一级分类")
            w.write(0, 1, u"二级分类")
            w.write(0, 2, u"三级分类")
            w.write(0, 3, u"版本号")
            w.write(0, 4, u"需求内容")
            w.write(0, 5, u"兼容性测试开始日期")
            w.write(0, 6, u"兼容性测试结束日期")
            w.write(0, 7, u"配额消耗")
            w.write(0, 8, u"通过率")
            w.write(0, 9, u"测试包名称")
            w.write(0, 10, u"总场景数")
            w.write(0, 11, u"执行兼容性场景数")
            w.write(0, 12, u"兼容性测试总轮次")
            w.write(0, 13, u"兼容性测试各轮次详情")
            w.write(0, 14, u"主要问题机型")
            w.write(0, 15, u"哪些版本出现过此问题")
            w.write(0, 16, u"安装失败")
            w.write(0, 17, u"启动失败")
            w.write(0, 18, u"闪退")
            w.write(0, 19, u"无响应")
            w.write(0, 20, u"意外终止")
            w.write(0, 21, u"卡死")
            w.write(0, 22, u"功能异常")
            w.write(0, 23, u"UI异常")
            w.write(0, 24, u"缺陷总数")
            w.write(0, 25, u"已关闭缺陷数量")
            w.write(0, 26, u"未关闭缺陷数量")
            w.write(0, 27, u"未关闭缺陷说明")
            w.write(0, 28, u"缺陷详细描述")
            w.write(0, 29, u"缺陷分析（出现问题的原因）")
            w.write(0, 30, u"缺陷解决描述（缺陷怎么解决的）")
            w.write(0, 31, u"开发人")
            w.write(0, 32, u"测试人")
            w.write(0, 33, u'最后修改人')
            w.write(0, 34, u'最新修改时间')

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
                connume = obj[8]
                pass_rate = obj[9]
                test_package_name = obj[10]
                auto_cases_num = obj[11]
                excu_allcase_num = obj[12]
                test_freq_num = obj[13]
                test_freq_content = obj[14]
                question_machine = obj[15]
                question_version = obj[16]
                install_failed = obj[17]
                tart_failed = obj[18]
                flash_back = obj[19]
                no_response = obj[20]
                accidental_termination = obj[21]
                stuck = obj[22]
                abnormal_function = obj[23]
                ui_exception = obj[24]
                bugs_num = obj[25]
                close_bug_num = obj[26]
                open_bug_num = obj[27]
                open_bug_content = obj[28]
                bug_description = obj[29]
                bug_analysis = obj[30]
                bug_resolvent = obj[31]
                developer = obj[32]
                tester = obj[33]
                upater = obj[34]
                upate_datetime = obj[35]

                w.write(excel_row, 0, data_prl)
                w.write(excel_row, 1, data_sec)
                w.write(excel_row, 2, data_thr)
                w.write(excel_row, 3, test_version)
                w.write(excel_row, 4, require_content)
                w.write(excel_row, 5, test_start_date.strftime('%Y%m%d'))
                w.write(excel_row, 6, test_end_date.strftime('%Y%m%d'))
                w.write(excel_row, 7, connume)
                w.write(excel_row, 8, pass_rate)
                w.write(excel_row, 9, test_package_name)
                w.write(excel_row, 10, auto_cases_num)
                w.write(excel_row, 11, excu_allcase_num)
                w.write(excel_row, 12, test_freq_num)
                w.write(excel_row, 13, test_freq_content)
                w.write(excel_row, 14, question_machine)
                w.write(excel_row, 15, question_version)
                w.write(excel_row, 16, install_failed)
                w.write(excel_row, 17, tart_failed)
                w.write(excel_row, 18, flash_back)
                w.write(excel_row, 19, no_response)
                w.write(excel_row, 20, accidental_termination)
                w.write(excel_row, 21, stuck)
                w.write(excel_row, 22, abnormal_function)
                w.write(excel_row, 23, ui_exception)
                w.write(excel_row, 24, bugs_num)
                w.write(excel_row, 25, close_bug_num)
                w.write(excel_row, 26, open_bug_num)
                w.write(excel_row, 27, open_bug_content)
                w.write(excel_row, 28, bug_description)
                w.write(excel_row, 29, bug_analysis)
                w.write(excel_row, 30, bug_resolvent)
                w.write(excel_row, 31, developer)
                w.write(excel_row, 32, tester)
                w.write(excel_row, 33, upater)
                w.write(excel_row, 34, upate_datetime.strftime('%Y-%m-%d %H:%M:%S'))
                excel_row += 1

    thread = threading.Thread(target=saveFile, args=(query,))
    thread.start()
    thread.join()
    cur_path()
    CURPATH = cur_path()
    ex_path = os.path.join(CURPATH, "兼容性测试数据统计.xls")
    ws.save(ex_path)
    cursor.close()
    db.close()
    print('Download of compatibility test report Successfully！')


if __name__ == '__main__':
    down_compatibility_testing()
