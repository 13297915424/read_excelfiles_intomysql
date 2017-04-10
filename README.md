读取文件夹以及其子文件夹下的所有excel文件，把每个excel文件的每个sheet页作为sheet表存入mysql数据库中。
依赖模块有： mysql.connector和xlrd
需要输入数据库配置和文件夹路径示例：
xpath="/home/user/下载/词云项目/data/未命名文件夹/未命名文件夹"
#xpath='C:/Users/flyminer/Desktop/新建文件夹'
sql_path={"host":"localhost",
          "user":"root",
          "port":3306,
          "password":"123456",
          "db":"词云项目2",
          "charset" : 'utf8'
          }
a=Excel_Msql(sql_path,xpath)
a.datahelper()
