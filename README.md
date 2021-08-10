安装第三方库之前保证是 Python3 版本

# 安装python库
$ pip3 install -r requirements.txt

# 配置数据库连接信息，修改setting.py文件

# 从数据库下载自动化测试数据统计.xls，执行down_automated_testing.py
$ python3 down_automated_testing.py

# 从数据库下载功能测试数据统计.xls，执行down_function_testing.py
$ python3 down_function_testing.py

# 从数据库下载功性能测试数据统计.xls，执行down_performance_testing.py
$ python3 down_performance_testing.py

# 从数据库下载兼容性测试数据统计.xls，执行down_performance_testing.py
$ python3 down_compatibility_testing.py

# 从数据库下载破坏性测试数据统计.xls，执行down_performance_testing.py
$ python3 down_destructive_testing.py

# 上传自动化测试数据统计.xls到数据库，执行upload_automated_testing.py文件
$ python3 upload_automated_testing.py

# 上传功能测试数据统计.xls到数据库，执行upload_function_testing.py文件
$ python3 upload_function_testing.py

# 上传性能测试数据统计.xls到数据库，执行upload_performance_testing.py文件
$ python3 upload_performance_testing.py

# 上传兼容性测试数据统计.xls到数据库，执行upload_compatibility_testing.py文件
$ python3 upload_compatibility_testing.py

# 上传破坏性测试数据统计.xls到数据库，执行upload_destructive_testing.py文件
$ python3 upload_destructive_testing.py

# 注意：
   执行下载功能操作时会自动生成当天日期的文件夹，下载的excel文件在该文件夹中，执行上传时也需要在当天日期的文件夹中修改要上传的excel文件。
