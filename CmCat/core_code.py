# 导入必要的插件
import io
import sys
import xlrd
import xlwt
import folium
import pymysql

# 导入UI类
from CmCat_Ui.cmcat_main import Ui_mainWin
from CmCat_Ui.cmcat_map import Ui_childMap

# 导入QT类
from PyQt5 import QtWidgets
from PyQt5.Qt import QFileDialog, QTime
from PyQt5.QtGui import QIcon
from PyQt5.QtCore import pyqtSignal
from PyQt5.QtChart import QChartView, QLineSeries, QValueAxis, QCategoryAxis
from PyQt5.QtWidgets import QApplication, QWidget, QVBoxLayout, QMessageBox, QMainWindow
from PyQt5.QtWebEngineWidgets import QWebEngineView


# 主窗口类
class LandingPanel(QMainWindow, Ui_mainWin, object):
    TableName = None
    File = None
    conn = None
    conn_signal = pyqtSignal(str, str, str, str, str, str, str)

    # 打开主UI界面
    def __init__(self):
        super().__init__()
        self.webview = None
        self.data = None
        self.m = None
        self.map = None
        self.charView = None
        self.y_Aix = None
        self.x_Aix = None
        self.series_1 = None
        self.chart = None
        self.setupUi(self)

        # 为按钮添加图标
        icon_cdb = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\cdb.png')
        icon_dc = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\dc.png')
        icon_open = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\open.png')
        icon_di = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\di.png')
        icon_map = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map.png')
        icon_qy = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\qy.png')
        icon_num_qy = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\num_qy.png')
        icon_1 = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\1.png')
        icon_2 = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\2.png')
        icon_te = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\te.png')
        icon_chart = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\chart.png')
        icon_4g = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\4g.png')
        icon_5g = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\5g.png')
        icon_clear = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\clear.png')
        icon_bb = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\bb.png')
        icon_e = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\E.png')
        icon_exit = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\exit.png')
        icon_cluster = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\cluster.png')
        icon_warn = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\warn.png')
        icon_warn_b = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\warn_b.png')
        icon_ms = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map_show.png')
        icon_mn = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map_num.png')
        icon_mll = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map_ll.png')
        icon_bs = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\bs.png')
        icon_cell = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\cell.png')

        # 为控件设置图标
        self.pushButton_connect.setIcon(icon_cdb)
        self.pushButton_disconnected.setIcon(icon_dc)
        self.pushButton_impOpen.setIcon(icon_open)
        self.pushButton_imp.setIcon(icon_di)
        self.pushButton_openqy.setIcon(icon_open)
        self.pushButton_qy.setIcon(icon_qy)
        self.pushButton_c.setIcon(icon_num_qy)
        self.pushButton_export.setIcon(icon_te)
        self.pushButton_export_2.setIcon(icon_te)
        self.radioButton_ll.setIcon(icon_1)
        self.radioButton_r.setIcon(icon_2)
        self.pushButton_cp.setIcon(icon_num_qy)
        self.pushButton_openqy_2.setIcon(icon_open)
        self.pushButton_qy_2.setIcon(icon_qy)
        self.pushButton_export_2.setIcon(icon_te)
        self.pushButton_f.setIcon(icon_chart)
        self.pushButton_export_3.setIcon(icon_te)
        self.pushButton_export_4.setIcon(icon_4g)
        self.pushButton_export_5.setIcon(icon_5g)
        self.pushButton_export_6.setIcon(icon_warn_b)
        self.pushButton_clear.setIcon(icon_clear)
        self.pushButton_clear_2.setIcon(icon_clear)
        self.pushButton_clear_3.setIcon(icon_clear)
        self.pushButton_clear_4.setIcon(icon_clear)
        self.pushButton_clear_5.setIcon(icon_clear)
        self.pushButton_clear_6.setIcon(icon_clear)
        self.checkBox_bb.setIcon(icon_bb)
        self.checkBox_emos.setIcon(icon_e)
        self.action_openMap.setIcon(icon_map)
        self.action_exit.setIcon(icon_exit)
        self.action_chart.setIcon(icon_chart)
        self.checkBox_cluster.setIcon(icon_cluster)
        self.pushButton_cluster.setIcon(icon_num_qy)
        self.pushButton_cCluster.setIcon(icon_clear)
        self.checkBox_warn.setIcon(icon_warn)
        self.pushButton_warn.setIcon(icon_num_qy)
        self.pushButton_show.setIcon(icon_ms)
        self.checkBox_num.setIcon(icon_mn)
        self.checkBox_ll.setIcon(icon_mll)
        self.checkBox_bs.setIcon(icon_bs)
        self.checkBox_cell.setIcon(icon_cell)

        # 按钮状态初始化
        self.pushButton_impOpen.setEnabled(False)
        self.pushButton_imp.setEnabled(False)
        self.pushButton_disconnected.setEnabled(False)
        self.pushButton_c.setEnabled(False)
        self.pushButton_f.setEnabled(False)
        self.pushButton_openqy.setEnabled(False)
        self.pushButton_qy.setEnabled(False)
        self.pushButton_openqy_2.setEnabled(False)
        self.pushButton_qy_2.setEnabled(False)
        self.radioButton_ll.setEnabled(False)
        self.radioButton_r.setEnabled(False)
        self.comboBox_r.setEnabled(False)
        self.pushButton_cp.setEnabled(False)
        self.checkBox_bb.setEnabled(False)
        self.checkBox_emos.setEnabled(False)
        self.checkBox_cluster.setEnabled(False)
        self.pushButton_cluster.setEnabled(False)
        self.checkBox_warn.setEnabled(False)
        self.pushButton_warn.setEnabled(False)
        self.pushButton_show.setEnabled(False)
        self.checkBox_num.setEnabled(False)
        self.checkBox_ll.setEnabled(False)
        self.checkBox_bs.setEnabled(False)
        self.checkBox_cell.setEnabled(False)

        # 数据库连接与选择文件并导入
        self.pushButton_connect.clicked.connect(self.mysql_connect)
        self.pushButton_impOpen.clicked.connect(self.open)
        self.pushButton_imp.clicked.connect(self.to_mysql)
        self.listWidget_impTb.clicked.connect(self.set_to_btn)
        self.pushButton_disconnected.clicked.connect(self.mysql_disconnect)

        # 手动输入查询
        self.pushButton_c.clicked.connect(self.load_h)
        self.pushButton_cp.clicked.connect(self.load_fzy)
        self.pushButton_cp.clicked.connect(self.load_lte)
        self.pushButton_cp.clicked.connect(self.load_nr)

        # 批量查询
        self.pushButton_openqy.clicked.connect(self.open_qy)
        self.pushButton_qy.clicked.connect(self.qy)
        self.pushButton_qy.clicked.connect(self.mark_show)
        self.listWidget_qy.clicked.connect(self.set_to_btn_qy)
        self.pushButton_openqy_2.clicked.connect(self.open_qy_2)
        self.pushButton_qy.clicked.connect(self.logic_op)
        self.pushButton_qy_2.clicked.connect(self.qy_2)
        self.listWidget_qy_2.clicked.connect(self.set_to_btn_qy_2)
        self.pushButton_qy_2.clicked.connect(self.load_lte_qy)
        self.pushButton_qy_2.clicked.connect(self.load_nr_qy)
        self.pushButton_qy_2.clicked.connect(self.logic_op)
        self.pushButton_qy_2.clicked.connect(self.mark_show_ll_r)
        self.pushButton_cp.clicked.connect(self.mark_show_ll_r)
        self.pushButton_show.clicked.connect(self.mark_show)

        # 家宽与emos工单的选择
        self.checkBox_bb.clicked.connect(self.emos_false)
        self.checkBox_emos.clicked.connect(self.bb_false)

        # 基站与小区的选择
        self.checkBox_bs.clicked.connect(self.cell_false)
        self.checkBox_cell.clicked.connect(self.bs_false)

        # 所有的清空与导出控件
        self.pushButton_clear.clicked.connect(self.clear_tableWidget_c)
        self.pushButton_export.clicked.connect(self.export_tableWidget_c)
        self.pushButton_clear_2.clicked.connect(self.clear_tableWidget_fc)
        self.pushButton_export_2.clicked.connect(self.export_tableWidget_fc)
        self.pushButton_clear_3.clicked.connect(self.clear_tableWidget_ac)
        self.pushButton_export_3.clicked.connect(self.export_tableWidget_ac)
        self.pushButton_clear_4.clicked.connect(self.clear_tableWidget_lte)
        self.pushButton_export_4.clicked.connect(self.export_tableWidget_lte)
        self.pushButton_clear_5.clicked.connect(self.clear_tableWidget_nr)
        self.pushButton_export_5.clicked.connect(self.export_tableWidget_nr)
        self.pushButton_clear_6.clicked.connect(self.clear_tableWidget_warn)
        self.pushButton_export_6.clicked.connect(self.export_tableWidget_warn)
        self.pushButton_warn.clicked.connect(self.load_warn)
        self.tableWidget_nr.itemClicked.connect(self.load_op)
        self.pushButton_cluster.clicked.connect(self.cluster)
        self.pushButton_cCluster.clicked.connect(self.cluster_clear)
        self.pushButton_f.clicked.connect(self.plot_chart)
        self.pushButton_qy.clicked.connect(self.set_to_btn_qy_2)

        # 数据库连接设置
        self.db_host = self.lineEdit_name.text()
        self.db_port = self.lineEdit_port.text()
        self.db_user = self.lineEdit_username.text()
        self.db_passwd = self.lineEdit_password.text()
        self.db_name = self.lineEdit_dbname.text()
        self.db_charset = "utf8"
        self.addr = self.lineEdit_addr.text()
        # 通过信号传递数据库连接设置参数
        self.map = Map()
        self.conn_signal.connect(self.map.get_data_list)
        self.conn_signal.emit(self.db_host, self.db_port, self.db_user, self.db_passwd, self.db_name, self.db_charset,
                              self.addr)

    # 连接数据库
    def mysql_connect(self):
        try:
            self.conn = pymysql.connect(host=self.db_host,
                                        port=int(self.db_port),
                                        user=self.db_user,
                                        passwd=self.db_passwd,
                                        db=self.db_name,
                                        charset=self.db_charset)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, "error", str(e))

        else:
            self.label_stateBar.setText("数据库连接成功！")
            self.pushButton_impOpen.setEnabled(True)
            self.pushButton_disconnected.setEnabled(True)
            self.pushButton_c.setEnabled(True)
            self.pushButton_cp.setEnabled(True)
            self.pushButton_openqy.setEnabled(True)
            self.pushButton_openqy_2.setEnabled(True)
            self.radioButton_ll.setEnabled(True)
            self.checkBox_bb.setEnabled(True)
            self.checkBox_emos.setEnabled(True)
            self.radioButton_ll.setChecked(True)
            self.checkBox_cluster.setEnabled(True)
            self.checkBox_cluster.setChecked(True)
            self.pushButton_cluster.setEnabled(True)
            self.checkBox_warn.setEnabled(True)
            self.pushButton_warn.setEnabled(True)
            self.pushButton_show.setEnabled(True)
            self.checkBox_num.setEnabled(True)
            self.checkBox_ll.setEnabled(True)
            self.checkBox_num.setChecked(True)
            self.checkBox_ll.setChecked(True)
            self.checkBox_bs.setEnabled(True)
            self.checkBox_cell.setEnabled(True)
            self.checkBox_cell.setChecked(True)

    # 断开数据库连接
    def mysql_disconnect(self):
        try:
            self.conn.close()
            self.label_stateBar.setText("数据库连接已断开！")
            self.pushButton_imp.setEnabled(False)
            self.pushButton_disconnected.setEnabled(False)
            self.pushButton_impOpen.setEnabled(False)
            self.pushButton_openqy.setEnabled(False)
            self.pushButton_qy.setEnabled(False)
            self.pushButton_c.setEnabled(False)
            self.pushButton_openqy_2.setEnabled(False)
            self.pushButton_qy_2.setEnabled(False)
            self.pushButton_openqy.setEnabled(False)
            self.pushButton_qy.setEnabled(False)
            self.pushButton_cp.setEnabled(False)
            self.pushButton_f.setEnabled(False)
            self.radioButton_ll.setEnabled(False)
            self.checkBox_bb.setEnabled(False)
            self.checkBox_emos.setEnabled(False)
            self.pushButton_cluster.setEnabled(False)
            self.checkBox_cluster.setEnabled(False)
            self.checkBox_warn.setEnabled(False)
            self.pushButton_warn.setEnabled(False)
            self.pushButton_show.setEnabled(False)
            self.checkBox_num.setEnabled(False)
            self.checkBox_ll.setEnabled(False)
            self.checkBox_bs.setEnabled(False)
            self.checkBox_cell.setEnabled(False)

        except BaseException as e:
            print(str(e))

    # 通过号码手动输入查询与批量查询模块
    # 打开文件夹
    def open(self):
        self.pushButton_impOpen.setEnabled(False)
        self.listWidget_impTb.clear()
        filepath, _ = QFileDialog.getOpenFileName(self,
                                                  '选中文件',
                                                  'E:\\PyCharm\\PythonProjects\\CmCat_MySQL',
                                                  'Excel files(*.xls *.xlsx)')
        self.lineEdit_impPath.setText(filepath)

        try:

            # 使用xlrd通过只读方式打开文件
            self.File = xlrd.open_workbook(filepath)

            # 获取所有sheet
            self.TableName = self.File.sheet_names()
            self.listWidget_impTb.addItems(self.TableName)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, "tip", "打开文件未成功")

    # 更新状态栏与导入按钮状态
    def set_to_btn(self):
        self.label_stateBar.setText("")
        self.pushButton_imp.setEnabled(True)

    # 导入数据库
    def to_mysql(self):
        start_time = QTime.currentTime()

        # 建立数据库游标
        cur = self.conn.cursor()

        # 获取选中的表
        sheet_name = self.File.sheet_by_name(self.listWidget_impTb.currentItem().text())

        # 获取第一行数据
        row1 = sheet_name.row_values(0)
        if self.listWidget_impTb.currentItem().text() == self.comboBox_select.currentText():
            table_name = self.listWidget_impTb.currentItem().text()
        if self.listWidget_impTb.currentItem().text() != self.comboBox_select.currentText():
            table_name = self.comboBox_select.currentText()

        # 整理建表SQL语句，默认varchar
        sql_create_table = "CREATE TABLE IF NOT EXISTS `" + table_name + "`("
        sql_create_table_word = ""
        for n in range(len(row1)-1):

            # 处理第一行数据为空的情况
            if row1[n] != "":
                sql_create_table_word += "`" + str(row1[n]) + "`" \
                                         + "varchar(31),"
            else:
                sql_create_table_word += "`" + str(n) + "`" \
                                         + "varchar(31),"

        # 处理第一行数据为空的情况
        if row1[-1] != "":
            sql_create_table += sql_create_table_word + "`" + str(row1[-1]) \
                                + "`" + "varchar(31) )DEFAULT CHARSET=utf8;"
        else:
            sql_create_table += sql_create_table_word + "`" + str(len(row1)-1) + "`" \
                                + " varchar(31) )DEFAULT CHARSET=utf8;"

        # 执行SQL语句建表，如果表不存在
        try:
            cur.execute(sql_create_table)

        except BaseException as e:
            QMessageBox.about(self, "error", str(e))
            print(e)

        # 整理插入语句
        sql_insert_array = []
        sql_insert_table = "INSERT INTO `" + table_name + "`() Values ("
        x = 1
        while x < sheet_name.nrows:
            sql_insert_word = ""
            for y in range(len(sheet_name.row_values(x))-1):
                sql_insert_word += "'" + str(sheet_name.row_values(x)[y]) + "',"
            sql_insert_table1 = sql_insert_table + sql_insert_word + "'" + str(sheet_name.row_values(x)[-1]) + "')"
            sql_insert_array.append(sql_insert_table1)
            x = x + 1

        # 写入数据库
        try:
            for n in sql_insert_array:
                cur.execute(str(n))

        except BaseException as e:
            print(e)
            QMessageBox.about(self, "error", str(e))

        finally:
            self.conn.commit()
            cur.close()

        # 记录结束时间
        end_time = QTime.currentTime()
        time = QTime.msecsTo(start_time, end_time) / 1000
        s = "成功导入：" + str(sheet_name.nrows) + "条数据，" + "用时：" + str(time) + "秒"
        self.label_stateBar.setText(s)

        # 只能导入一次
        self.pushButton_imp.setEnabled(False)

        # 打开文件路径可用
        self.pushButton_impOpen.setEnabled(True)

    # 打开查询文件
    def open_qy(self):
        self.listWidget_qy.clear()
        filepath, _ = QFileDialog.getOpenFileName(self,
                                                  '选中文件',
                                                  'E:\\PyCharm\\PythonProjects\\CmCat_Query',
                                                  'Excel files(*.xls *.xlsx)')
        self.lineEdit_qy.setText(filepath)

        try:
            # 使用xlrd通过只读方式打开文件
            self.File = xlrd.open_workbook(filepath)

            # 获取所有sheet
            self.TableName = self.File.sheet_names()
            self.listWidget_qy.addItems(self.TableName)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, "error", "打开文件未成功")

    # 启用查询按钮
    def set_to_btn_qy(self):
        self.pushButton_qy.setEnabled(True)

    # 批量查询
    def qy(self):
        try:
            cursor = self.conn.cursor()
            sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
            sql_select_sum = ''
            for i in range(1, len(sheet_name.col_values(6)) - 1):
                sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
            sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
            sql_select = "select `电话接入时间`, `接触流水号`, `故障号码`, `国家 省/市 区县`, `经度`, `纬度`, `经纬度`" \
                         "from `投诉-广义咨询` where 故障号码 in (" + sql_select_sum + ")"
            cursor.execute(sql_select)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_c.setRowCount(row)  # 设置表格行数
            self.tableWidget_c.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_c.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            if self.checkBox_bb.isChecked():
                sql_select_sum = ''
                for i in range(1, len(sheet_name.col_values(6)) - 1):
                    sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
                sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
                sql_select = "select `工单号`, `建单时间`, `联系电话`, `投诉地点`" \
                             "from `投诉-家宽` where 联系电话 in (" + sql_select_sum + ")"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_fc.setRowCount(row)  # 设置表格行数
                self.tableWidget_fc.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_fc.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            if self.checkBox_emos.isChecked():
                sql_select_sum = ''
                for i in range(1, len(sheet_name.col_values(6)) - 1):
                    sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
                sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
                sql_select = "select `工单号`, `建单时间`, `联系电话`, `投诉地点`" \
                             "from `投诉-emos工单` where 联系电话 in (" + sql_select_sum + ")"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_fc.setRowCount(row)  # 设置表格行数
                self.tableWidget_fc.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_fc.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            if (not self.radioButton_ll.isChecked()) and (not self.radioButton_r.isChecked()):
                QMessageBox.about(self, 'tip', '请选择查询方式')

            if self.radioButton_ll.isChecked():
                cursor = self.conn.cursor()
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `电话接入时间`, `接触流水号`, `故障号码`, `国家 省/市 区县`, `经度`, `纬度`, `经纬度`" \
                               "from `投诉-广义咨询`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_ac.setRowCount(row)  # 设置表格行数
                self.tableWidget_ac.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_ac.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            if self.radioButton_ll.isChecked():
                cursor = self.conn.cursor()
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `基站名称`, `小区名称`, `经度`, `纬度` from `工参-lte`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_lte.setRowCount(row)  # 设置表格行数
                self.tableWidget_lte.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_lte.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            if self.radioButton_ll.isChecked():
                cursor = self.conn.cursor()
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `基站名称`, `小区名称`, `GNSS经度`, `GNSS纬度` from `工参-nr`" \
                               "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (GNSS经度 like '" + str(
                            sheet_name.col_values(31)[i]) + "%'" + \
                                         "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_nr.setRowCount(row)  # 设置表格行数
                self.tableWidget_nr.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_nr.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

        except BaseException as e:
            print(e)
            QMessageBox.about(self, 'error', '文件表格数据不正确')

    # 输入号码查询历史投诉
    def load_h(self):
        try:
            cursor = self.conn.cursor()
            sql_select = "select `电话接入时间`, `接触流水号`, `故障号码`, `国家 省/市 区县`, `经度`, `纬度`, `经纬度`" \
                         "from `投诉-广义咨询` where 故障号码 in (" + self.lineEdit_num.text() + " )"
            cursor.execute(sql_select)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_c.setRowCount(row)  # 设置表格行数
            self.tableWidget_c.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_c.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

        except BaseException as e:
            print(e)
            if self.lineEdit_num.text() == '':
                QMessageBox.about(self, 'error', '请输入查询号码')
            else:
                QMessageBox.about(self, 'error', '查询结果为空')

    # 清空查询结果
    def clear_tableWidget_c(self):
        self.tableWidget_c.clear()

    # 导出为文件
    def export_tableWidget_c(self):
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("广义咨询全量分析表")

        # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
        for row in range(self.tableWidget_c.rowCount()):
            for col in range(self.tableWidget_c.columnCount()):
                worksheet.write(row, col, self.tableWidget_c.item(row, col).text())

        # 保存表格至excel
        workbook.save("E:\\PyCharm\\PythonProjects\\CmCat_Export\\广义咨询全量分析表.xls")
        QMessageBox.about(self, 'tip', '导出成功')

    # 清空查询结果
    def clear_tableWidget_fc(self):
        self.tableWidget_fc.clear()

    # 导出为文件
    def export_tableWidget_fc(self):
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("家庭宽带投诉表")

        # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
        for row in range(self.tableWidget_fc.rowCount()):
            for col in range(self.tableWidget_fc.columnCount()):
                worksheet.write(row, col, self.tableWidget_fc.item(row, col).text())

        # 保存表格至excel
        workbook.save("E:\\PyCharm\\PythonProjects\\CmCat_Export\\投诉全量工单明细表.xls")
        QMessageBox.about(self, 'tip', '导出成功')

    # 清空查询结果
    def clear_tableWidget_ac(self):
        self.tableWidget_ac.clear()

    # 导出为文件
    def export_tableWidget_ac(self):
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("广义咨询全量分析表_经纬度")

        # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
        for row in range(self.tableWidget_ac.rowCount()):
            for col in range(self.tableWidget_ac.columnCount()):
                worksheet.write(row, col, self.tableWidget_fc.item(row, col).text())

        # 保存表格至excel
        workbook.save("E:\\PyCharm\\PythonProjects\\CmCat_Export\\广义咨询全量分析表_经纬度.xls")
        QMessageBox.about(self, 'tip', '导出成功')

    # 打开文件经纬度批量查询
    def open_qy_2(self):
        self.listWidget_qy_2.clear()
        filepath, _ = QFileDialog.getOpenFileName(self,
                                                  '选中文件',
                                                  'E:\\PyCharm\\PythonProjects\\CmCat_Query',
                                                  'Excel files(*.xls *.xlsx)')
        self.lineEdit_qy_2.setText(filepath)

        try:
            # 使用xlrd通过只读方式打开文件
            self.File = xlrd.open_workbook(filepath)

            # 获取所有sheet
            self.TableName = self.File.sheet_names()
            self.listWidget_qy_2.addItems(self.TableName)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, 'error', '打开文件未成功')

    # 查询按钮可用
    def set_to_btn_qy_2(self):
        self.pushButton_qy_2.setEnabled(True)

    # 根据文件批量查询
    def qy_2(self):
        try:
            if (not self.radioButton_ll.isChecked()) and (not self.radioButton_r.isChecked()):
                QMessageBox.about(self, 'tip', '请选择查询方式')
            if self.radioButton_ll.isChecked():
                cursor = self.conn.cursor()
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `电话接入时间`, `接触流水号`, `故障号码`, `国家 省/市 区县`, `经度`, `纬度`, `经纬度`" \
                               "from `投诉-广义咨询`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_ac.setRowCount(row)  # 设置表格行数
                self.tableWidget_ac.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_ac.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

        except BaseException as e:
            print(e)
            QMessageBox.about(self, 'error', '文件表格数据不正确')

    # 聚类
    def cluster(self):
        plot_begin = self.dateEdit_bCluster.date().toString("yyyy-MM-dd")
        plot_end = self.dateEdit_eCluster.date().toString("yyyy-MM-dd")
        cursor = self.conn.cursor()
        sql_count = "SELECT `国家 省/市 区县`,count( * ) AS COUNT FROM `投诉-广义咨询` " \
                    "WHERE SUBSTRING(电话接入时间,1,10)>= '" + plot_begin + \
                    "' and SUBSTRING(电话接入时间,1,10)<='" + plot_end + \
                    "' GROUP BY `国家 省/市 区县` ORDER BY COUNT DESC"
        cursor.execute(sql_count)
        res = cursor.fetchall()
        row = cursor.rowcount
        col = len(res[0])
        self.tableWidget_cluster.setRowCount(row)  # 设置表格行数
        self.tableWidget_cluster.setColumnCount(col)  # 设置表格列数
        for i in range(0, row):  # 遍历行
            for j in range(0, col):  # 遍历列
                self.tableWidget_cluster.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

    def cluster_clear(self):
        self.tableWidget_cluster.clear()

    def emos_false(self):
        self.checkBox_emos.setChecked(False)

    def bb_false(self):
        self.checkBox_bb.setChecked(False)

    def bs_false(self):
        self.checkBox_bs.setChecked(False)

    def cell_false(self):
        self.checkBox_cell.setChecked(False)

    # 查询告警
    def load_warn(self):
        if self.checkBox_warn.isChecked():
            try:
                cursor = self.conn.cursor()
                sql_select = "select * from `告警-lst_almaf` where NAME = '" + self.lineEdit_addr.text() + "';"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_warn.setRowCount(row)  # 设置表格行数
                self.tableWidget_warn.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_warn.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            except BaseException as e:
                print(e)
                QMessageBox.about(self, 'tip', '查询结果为空')

    # 清空查询结果
    def clear_tableWidget_warn(self):
        self.tableWidget_warn.clear()

    # 导出为文件
    def export_tableWidget_warn(self):
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("广义咨询全量分析表")

        # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
        for row in range(self.tableWidget_warn.rowCount()):
            for col in range(self.tableWidget_warn.columnCount()):
                worksheet.write(row, col, self.tableWidget_warn.item(row, col).text())

        # 保存表格至excel
        workbook.save("E:\\PyCharm\\PythonProjects\\CmCat_Export\\告警表.xls")
        QMessageBox.about(self, 'tip', '导出成功')

    # 查询4g工参
    def load_lte_qy(self):
        if self.radioButton_ll.isChecked():
            cursor = self.conn.cursor()
            sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
            sql_select_sum = ""
            sql_select_1 = "select `基站名称`, `小区名称`, `经度`, `纬度` from `工参-lte`" \
                           "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                           "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
            for i in range(2, len(sheet_name.col_values(31))):
                if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                    sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                     "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
            sql_select = sql_select_1 + sql_select_sum
            cursor.execute(sql_select)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_lte.setRowCount(row)  # 设置表格行数
            self.tableWidget_lte.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_lte.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

    # 清空查询结果
    def clear_tableWidget_lte(self):
        self.tableWidget_lte.clear()

    # 导出为文件
    def export_tableWidget_lte(self):
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("LTE网络工参表")

        # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
        for row in range(self.tableWidget_lte.rowCount()):
            for col in range(self.tableWidget_lte.columnCount()):
                worksheet.write(row, col, self.tableWidget_lte.item(row, col).text())

        # 保存表格至excel
        workbook.save("E:\\PyCharm\\PythonProjects\\CmCat_Export\\LTE网络工参表.xls")
        QMessageBox.about(self, 'tip', '导出成功')

    # 批量查询5g工参
    def load_nr_qy(self):
        if self.radioButton_ll.isChecked():
            cursor = self.conn.cursor()
            sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
            sql_select_sum = ""
            sql_select_1 = "select `基站名称`, `小区名称`, `GNSS经度`, `GNSS纬度` from `工参-nr`" \
                           "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                           "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
            for i in range(2, len(sheet_name.col_values(31))):
                if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                    sql_select_sum = sql_select_sum + "or (GNSS经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                     "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
            sql_select = sql_select_1 + sql_select_sum
            cursor.execute(sql_select)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_nr.setRowCount(row)  # 设置表格行数
            self.tableWidget_nr.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_nr.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

    # 清空查询结果
    def clear_tableWidget_nr(self):
        self.tableWidget_nr.clear()

    # 导出为文件
    def export_tableWidget_nr(self):
        # 创建新的workbook（其实就是创建新的excel）
        workbook = xlwt.Workbook(encoding='ascii')

        # 创建新的sheet表
        worksheet = workbook.add_sheet("NR网络工参表")

        # 往表格写入内容,这里可以通过嵌套两层For循环历便表格内容后写入到excel文件中
        for row in range(self.tableWidget_nr.rowCount()):
            for col in range(self.tableWidget_nr.columnCount()):
                worksheet.write(row, col, self.tableWidget_nr.item(row, col).text())

        # 保存表格至excel
        workbook.save("E:\\PyCharm\\PythonProjects\\CmCat_Export\\NR网络工参表.xls")
        QMessageBox.about(self, 'tip', '导出成功')

    def logic_op(self):
        self.radioButton_r.setEnabled(True)
        self.comboBox_r.setEnabled(True)

    # 根据截取经纬度数据至两位小数进行模糊匹配或进行范围查询
    def load_fzy(self):
        # 栅格查询
        if (not self.radioButton_ll.isChecked()) and (not self.radioButton_r.isChecked()):
            QMessageBox.about(self, 'tip', '请选择查询方式')
        if self.radioButton_ll.isChecked():
            lng_list = self.lineEdit_lng.text().split(',')
            lat_list = self.lineEdit_lat.text().split(',')
            cursor = self.conn.cursor()
            sql_select_1 = "select `电话接入时间`, `接触流水号`, `故障号码`, `国家 省/市 区县`, `经度`, `纬度`, `经纬度`" \
                           "from `投诉-广义咨询`" \
                           "where (经度 like '" + lng_list[0] + "%'" + \
                           "and 纬度 like '" + lat_list[0] + "%')"
            sql_select_sum = ""
            for i in range(1, len(lng_list)):
                sql_select_sum = sql_select_sum + "or (经度 like '" + lng_list[i] + "%'" + \
                                 "and 纬度 like '" + lat_list[i] + "%')"
            cursor.execute(sql_select_1 + sql_select_sum)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_ac.setRowCount(row)  # 设置表格行数
            self.tableWidget_ac.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_ac.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

        # 范围查询
        if self.radioButton_r.isChecked():
            try:
                radius = self.comboBox_r.currentText()
                cursor = self.conn.cursor()
                sql_select = "select `电话接入时间`, `接触流水号`, `故障号码`, `国家 省/市 区县`, `经度`, `纬度`, `经纬度` " \
                             "from `投诉-广义咨询` " \
                             "where (acos(sin((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*sin((convert(纬度,float)*3.1415)/180) " \
                             "+ cos((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*cos((convert(纬度,float)*3.1415)/180) " \
                             "* cos((" + self.lineEdit_lng.text() + "*3.1415)/180" + \
                             "-(convert(经度,float)*3.1415)/180))*6370.996)<" + radius + \
                             " and 经度!='' and  纬度!='';"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_ac.setRowCount(row)  # 设置表格行数
                self.tableWidget_ac.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_ac.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            except BaseException as e:
                print(e)
                QMessageBox.about(self, 'tip', '查询结果为空')

    # 批量查询4g工参
    def load_lte(self):
        if self.radioButton_ll.isChecked():
            lng_list = self.lineEdit_lng.text().split(',')
            lat_list = self.lineEdit_lat.text().split(',')
            cursor = self.conn.cursor()
            sql_select_1 = "select `基站名称`, `小区名称`, `经度`, `纬度` from `工参-lte`" \
                           "where (经度 like '" + lng_list[0] + "%'" + \
                           "and 纬度 like '" + lat_list[0] + "%')"
            sql_select_sum = ""
            for i in range(1, len(lng_list)):
                sql_select_sum = sql_select_sum + "or (经度 like '" + lng_list[i] + "%'" + \
                                 "and 纬度 like '" + lat_list[i] + "%')"
            cursor.execute(sql_select_1 + sql_select_sum)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_lte.setRowCount(row)  # 设置表格行数
            self.tableWidget_lte.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_lte.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

        # 范围查询
        if self.radioButton_r.isChecked():
            try:
                radius = self.comboBox_r.currentText()
                cursor = self.conn.cursor()
                sql_select = "select `基站名称`, `小区名称`, `经度`, `纬度` from `工参-lte`" \
                             "where (acos(sin((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*sin((convert(纬度,float)*3.1415)/180) " \
                             "+ cos((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*cos((convert(纬度,float)*3.1415)/180) " \
                             "* cos((" + self.lineEdit_lng.text() + "*3.1415)/180" + \
                             "-(convert(经度,float)*3.1415)/180))*6370.996)<" + radius + \
                             " and 经度!='' and  纬度!='';"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_lte.setRowCount(row)  # 设置表格行数
                self.tableWidget_lte.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_lte.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            except BaseException as e:
                print(e)
                QMessageBox.about(self, 'tip', '查询结果为空')

    # 查询5g工参
    def load_nr(self):
        if self.radioButton_ll.isChecked():
            lng_list = self.lineEdit_lng.text().split(',')
            lat_list = self.lineEdit_lat.text().split(',')
            cursor = self.conn.cursor()
            sql_select_1 = "select `基站名称`, `小区名称`, `GNSS经度`, `GNSS纬度` from `工参-nr`" \
                           "where (GNSS经度 like '" + lng_list[0] + "%'" + \
                           "and GNSS纬度 like '" + lat_list[0] + "%')"
            sql_select_sum = ""
            for i in range(1, len(lng_list)):
                sql_select_sum = sql_select_sum + "or (GNSS经度 like '" + lng_list[i] + "%'" + \
                                 "and GNSS纬度 like '" + lat_list[i] + "%')"
            cursor.execute(sql_select_1 + sql_select_sum)
            res = cursor.fetchall()
            row = cursor.rowcount
            col = len(res[0])
            self.tableWidget_nr.setRowCount(row)  # 设置表格行数
            self.tableWidget_nr.setColumnCount(col)  # 设置表格列数
            for i in range(0, row):  # 遍历行
                for j in range(0, col):  # 遍历列
                    self.tableWidget_nr.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

        # 范围查询
        if self.radioButton_r.isChecked():
            try:
                radius = self.comboBox_r.currentText()
                cursor = self.conn.cursor()
                sql_select = "select `基站名称`, `小区名称`, `GNSS经度`, `GNSS纬度` from `工参-nr`" \
                             "where (acos(sin((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*sin((convert(GNSS纬度,float)*3.1415)/180) " \
                             "+ cos((" + self.lineEdit_lat.text() + "*3.1415)/180)" \
                             "*cos((convert(GNSS纬度,float)*3.1415)/180) " \
                             "* cos((" + self.lineEdit_lng.text() + "*3.1415)/180" \
                             "-(convert(GNSS经度,float)*3.1415)/180))*6370.996)<" + radius + \
                             " and GNSS经度!='' and  GNSS纬度!='';"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount
                col = len(res[0])
                self.tableWidget_nr.setRowCount(row)  # 设置表格行数
                self.tableWidget_nr.setColumnCount(col)  # 设置表格列数
                for i in range(0, row):  # 遍历行
                    for j in range(0, col):  # 遍历列
                        self.tableWidget_nr.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))

            except BaseException as e:
                print(e)
                QMessageBox.about(self, 'tip', '查询结果为空')

    # 点选站址后文本控件状态
    def load_op(self, item):
        addr_text = self.tableWidget_nr.item(item.row(), item.column()).text()
        addr_row = self.tableWidget_nr.currentRow()
        lng_text = self.tableWidget_nr.item(addr_row, 2).text()
        lat_text = self.tableWidget_nr.item(addr_row, 3).text()
        self.lineEdit_lng.setText(lng_text)
        self.lineEdit_lat.setText(lat_text)
        self.lineEdit_addr.setText(addr_text)
        self.radioButton_r.setChecked(True)
        self.radioButton_ll.setChecked(False)
        self.radioButton_ll.setChecked(False)
        self.pushButton_f.setEnabled(True)
        self.checkBox_warn.setChecked(True)

    def plot_chart(self):
        try:
            if not self.checkBox_bs.isChecked() and not self.checkBox_cell.isChecked():
                QMessageBox.about(self, 'tip', '请选择业务量显示方式')

            if not self.checkBox_bs.isChecked() and self.checkBox_cell.isChecked():
                plot_begin = self.dateEdit_begin.date().toString("yyyy-MM-dd")
                plot_end = self.dateEdit_end.date().toString("yyyy-MM-dd")
                cursor_b = self.conn.cursor()
                cursor_t = self.conn.cursor()
                self.series_1 = QLineSeries()  # 将类QLineSeries实例化
                sql_select_b = "select `总业务量` from `业务-总业务量` where 小区名称 = '" + self.lineEdit_addr.text() + "'"
                sql_select_t = "select `日期` from `业务-总业务量` where 小区名称 = '" + self.lineEdit_addr.text() + "'"
                cursor_b.execute(sql_select_b)
                cursor_t.execute(sql_select_t)
                res_b = cursor_b.fetchall()
                res_t = cursor_t.fetchall()
                row_t = cursor_t.rowcount
                self.series_1.setName("Line Chart")  # 折线命名
                self.x_Aix = QCategoryAxis()  # 定义y轴，实例化
                begin = -1
                end = -1
                for i in range(0, row_t):
                    if str(res_t[i][0]) == plot_begin:
                        begin = i
                    if str(res_t[i][0]) == plot_end:
                        end = i
                for i in range(begin, end + 1):
                    self.x_Aix.append(str(res_t[i][0]), i - begin)
                max_num = float(res_b[begin][0])
                min_num = float(res_b[begin][0])
                for i in range(begin, end + 1):
                    self.series_1.append(i - begin, float(res_b[i][0]))  # 折线添加坐标点清单
                    if float(res_b[i][0]) > max_num:
                        max_num = float(res_b[i][0])
                    if float(res_b[i][0]) < min_num:
                        min_num = float(res_b[i][0])
                self.x_Aix.setRange(0.00, float(end - begin))
                self.x_Aix.setTitleText('Date')
                self.y_Aix = QValueAxis()  # 定义y轴，实例化
                self.y_Aix.setRange(min_num, max_num)
                self.y_Aix.setLabelFormat("%0.2f")
                self.y_Aix.setTitleText('KBit')
                self.charView = QChartView(self.widget_chart)  # 定义charView，父窗体类型为widget
                self.charView.setGeometry(0, 0, self.widget_chart.width(),
                                          self.widget_chart.height())  # 设置charView位置、大小
                self.charView.chart().addSeries(self.series_1)  # 添加折线
                self.charView.chart().setAxisX(self.x_Aix)  # 设置x轴属性
                self.charView.chart().setAxisY(self.y_Aix)  # 设置y轴属性
                self.charView.show()  # 显示charView

            if not self.checkBox_cell.isChecked() and self.checkBox_bs.isChecked():
                plot_begin = self.dateEdit_begin.date().toString("yyyy-MM-dd")
                plot_end = self.dateEdit_end.date().toString("yyyy-MM-dd")
                cursor_b = self.conn.cursor()
                cursor_t = self.conn.cursor()
                cursor_count = self.conn.cursor()
                self.series_1 = QLineSeries()  # 将类QLineSeries实例化
                sql_select_b = "select `总业务量` from `业务-总业务量` where 基站名称 = '" + self.lineEdit_addr.text() + "'"
                sql_select_t = "select `日期` from `业务-总业务量` where 基站名称 = '" + self.lineEdit_addr.text() + "'"
                sql_count = "SELECT `日期`,count( * ) AS COUNT FROM `业务-总业务量` " \
                            "where 基站名称 = '" + self.lineEdit_addr.text() + \
                            "' GROUP BY `日期` ORDER BY COUNT DESC"
                cursor_b.execute(sql_select_b)
                cursor_t.execute(sql_select_t)
                cursor_count.execute(sql_count)
                res_b = cursor_b.fetchall()
                res_count = cursor_count.fetchall()  # 用来横轴显示时间
                row_t = cursor_t.rowcount
                row_count = cursor_count.rowcount
                step = int(row_t/row_count)  # 用来跳转索引
                res_b_tmp = []
                tmp = 0
                for i in range(0, row_t, step):
                    for j in range(i, i+step):
                        tmp = tmp + float(res_b[j][0])
                    res_b_tmp.append(tmp)
                    tmp = 0
                self.series_1.setName("Line Chart")  # 折线命名
                self.x_Aix = QCategoryAxis()  # 定义y轴，实例化
                begin = -1
                end = -1
                for i in range(0, row_count):
                    if str(res_count[i][0]) == plot_begin:
                        begin = i
                    if str(res_count[i][0]) == plot_end:
                        end = i
                print(begin, end)
                for i in range(begin, end + 1):
                    self.x_Aix.append(str(res_count[i][0]), i - begin)
                max_num = res_b_tmp[begin]
                min_num = res_b_tmp[begin]
                for i in range(begin, end + 1):
                    self.series_1.append(i - begin, res_b_tmp[i])  # 折线添加坐标点清单
                    if res_b_tmp[i] > max_num:
                        max_num = res_b_tmp[i]
                    if res_b_tmp[i] < min_num:
                        min_num = res_b_tmp[i]
                self.x_Aix.setRange(0.00, float(end - begin))
                self.x_Aix.setTitleText('Date')
                self.y_Aix = QValueAxis()  # 定义y轴，实例化
                self.y_Aix.setRange(float(min_num), float(max_num))
                self.y_Aix.setLabelFormat("%0.2f")
                self.y_Aix.setTitleText('KBit')
                self.charView = QChartView(self.widget_chart)  # 定义charView，父窗体类型为widget
                self.charView.setGeometry(0, 0, self.widget_chart.width(),
                                          self.widget_chart.height())  # 设置charView位置、大小
                self.charView.chart().addSeries(self.series_1)  # 添加折线
                self.charView.chart().setAxisX(self.x_Aix)  # 设置x轴属性
                self.charView.chart().setAxisY(self.y_Aix)  # 设置y轴属性
                self.charView.show()  # 显示charView

        except BaseException as e:
            print(e)
            QMessageBox.about(self, 'tip', '查询结果为空')

    # # 查询业务量
    # def load_tfc(self):
    #     try:
    #         cursor = self.conn.cursor()
    #         sql_select = "select * from `业务量-业务量-含总业务量` where 基站名称 = '" + self.lineEdit_addr.text() + "'"
    #         cursor.execute(sql_select)
    #         res = cursor.fetchall()
    #         row = cursor.rowcount
    #         col = len(res[0])
    #         self.tableWidget_f.setRowCount(row)  # 设置表格行数
    #         self.tableWidget_f.setColumnCount(col)  # 设置表格列数
    #         for i in range(0, row):  # 遍历行
    #             for j in range(0, col):  # 遍历列
    #                 self.tableWidget_f.setItem(i, j, QtWidgets.QTableWidgetItem(str(res[i][j])))
    #
    #     except BaseException as e:
    #         print(e)
    #         QMessageBox.about(self, '提示', '查询结果为空')

    def mark_show(self):
        try:
            cursor = self.conn.cursor()

            # 选中标注方式
            if (not self.checkBox_num.isChecked()) and (not self.checkBox_ll.isChecked()):
                QMessageBox.about(self, 'tip', '请选择显示内容')

            if self.checkBox_num.isChecked() and not self.checkBox_ll.isChecked():
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ''
                for i in range(1, len(sheet_name.col_values(6)) - 1):
                    sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
                sql_select_sum = sql_select_sum + str(
                    int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
                sql_select = "select `经度`, `纬度`, `故障号码` " \
                             "from `投诉-广义咨询` where 故障号码 in (" + sql_select_sum + ") and (经度 !='' and 纬度 != '');"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(res[0][1]), float(res[0][0])],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row):
                    folium.Marker(
                        location=[float(res[i][1]), float(res[i][0])],
                        popup='故障号码：' + str(res[i][2]),
                        icon=folium.Icon(icon='cny', color='blue')
                    ).add_to(self.m)
                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

            if self.checkBox_ll.isChecked() and not self.checkBox_num.isChecked():

                # 投诉查询
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度`, `常驻小区` from `投诉-广义咨询`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(res[0][1]), float(res[0][0])],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row):
                    folium.Marker(
                        location=[float(res[i][1]), float(res[i][0])],
                        popup='常驻小区：' + str(res[i][2]),
                        icon=folium.Icon(icon='cny', color='red')
                    ).add_to(self.m)

                # 4G宏站查询
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度`, `小区名称` from `工参-lte`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_lte = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_lte[i][1]), float(res_lte[i][0])],
                        popup='小区名称：' + str(res_lte[i][2]),
                        icon=folium.Icon(icon='cny', color='purple')
                    ).add_to(self.m)

                # 5G宏站查询
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `GNSS经度`, `GNSS纬度`, `小区名称` from `工参-nr`" \
                               "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (GNSS经度 like '" +\
                                         str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_nr = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_nr[i][1]), float(res_nr[i][0])],
                        popup='小区名称：' + str(res_nr[i][2]),
                        icon=folium.Icon(icon='cny', color='green')
                    ).add_to(self.m)

                # 更新地图
                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

            if self.checkBox_num.isChecked() and self.checkBox_ll.isChecked():

                # 号码查询投诉
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ''
                for i in range(1, len(sheet_name.col_values(6)) - 1):
                    sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
                sql_select_sum = sql_select_sum + str(
                    int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
                sql_select = "select `经度`, `纬度`, `故障号码`" \
                             "from `投诉-广义咨询` where 故障号码 in (" + sql_select_sum + ") and (经度 !='' and 纬度 != '');"
                cursor.execute(sql_select)
                res_num = cursor.fetchall()
                row_num = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(res_num[0][1]), float(res_num[0][0])],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row_num):
                    folium.Marker(
                        location=[float(res_num[i][1]), float(res_num[i][0])],
                        popup='故障号码：' + str(res_num[i][2]),
                        icon=folium.Icon(icon='cny', color='blue')
                    ).add_to(self.m)

                # 经纬度查询投诉
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度`, `故障号码`" \
                               "from `投诉-广义咨询`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_ll = cursor.fetchall()
                row_ll = cursor.rowcount
                for i in range(row_ll):
                    folium.Marker(
                        location=[float(res_ll[i][1]), float(res_ll[i][0])],
                        popup='故障号码：' + str(res_ll[i][2]),
                        icon=folium.Icon(icon='cny', color='red')
                    ).add_to(self.m)

                # 4G宏站
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度`, `小区名称` from `工参-lte`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_lte = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_lte[i][1]), float(res_lte[i][0])],
                        popup='小区名称：' + str(res_lte[i][2]),
                        icon=folium.Icon(icon='cny', color='purple')
                    ).add_to(self.m)

                # 5G宏站
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `GNSS经度`, `GNSS纬度`, `小区名称` from `工参-nr`" \
                               "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (GNSS经度 like '" + \
                                         str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_nr = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_nr[i][1]), float(res_nr[i][0])],
                        popup='小区名称：' + str(res_nr[i][2]),
                        icon=folium.Icon(icon='cny', color='green')
                    ).add_to(self.m)

                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, 'error', '文件表格数据不正确')

    def mark_show_ll_r(self):
        cursor = self.conn.cursor()

        if self.radioButton_ll.isChecked():
            # 投诉查询
            sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
            sql_select_sum = ""
            sql_select_1 = "select `经度`, `纬度`, `故障号码` from `投诉-广义咨询`" \
                           "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                           "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
            for i in range(2, len(sheet_name.col_values(31))):
                if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                    sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                     "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
            sql_select = sql_select_1 + sql_select_sum
            cursor.execute(sql_select)
            res = cursor.fetchall()
            row = cursor.rowcount

            # 选取中心点
            self.gridLayout_map.removeWidget(self.webview)
            self.m = folium.Map(
                location=[float(res[0][1]), float(res[0][0])],
                tiles='Stamen Terrain',
                zoom_start=14
            )
            for i in range(row):
                folium.Marker(
                    location=[float(res[i][1]), float(res[i][0])],
                    popup='故障号码：' + str(res[i][2]),
                    icon=folium.Icon(icon='cny', color='red')
                ).add_to(self.m)

            # 4G宏站查询
            sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
            sql_select_sum = ""
            sql_select_1 = "select `经度`, `纬度`, `小区名称` from `工参-lte`" \
                           "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                           "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
            for i in range(2, len(sheet_name.col_values(31))):
                if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                    sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                     "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
            sql_select = sql_select_1 + sql_select_sum
            cursor.execute(sql_select)
            res_lte = cursor.fetchall()
            row = cursor.rowcount
            for i in range(row):
                folium.Marker(
                    location=[float(res_lte[i][1]), float(res_lte[i][0])],
                    popup='小区名称：' + str(res_lte[i][2]),
                    icon=folium.Icon(icon='cny', color='purple')
                ).add_to(self.m)

            # 5G宏站查询
            sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
            sql_select_sum = ""
            sql_select_1 = "select `GNSS经度`, `GNSS纬度`, `小区名称` from `工参-nr`" \
                           "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                           "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
            for i in range(2, len(sheet_name.col_values(31))):
                if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                    sql_select_sum = sql_select_sum + "or (GNSS经度 like '" + \
                                     str(sheet_name.col_values(31)[i]) + "%'" + \
                                     "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
            sql_select = sql_select_1 + sql_select_sum
            cursor.execute(sql_select)
            res_nr = cursor.fetchall()
            row = cursor.rowcount
            for i in range(row):
                folium.Marker(
                    location=[float(res_nr[i][1]), float(res_nr[i][0])],
                    popup='小区名称：' + str(res_nr[i][2]),
                    icon=folium.Icon(icon='cny', color='green')
                ).add_to(self.m)

            # 更新地图
            self.data = io.BytesIO()
            self.m.save(self.data, close_file=False)
            self.webview = QWebEngineView()
            self.webview.setHtml(self.data.getvalue().decode())
            self.gridLayout_map.addWidget(self.webview)

        # 范围查询
        if self.radioButton_r.isChecked():
            try:
                radius = self.comboBox_r.currentText()
                cursor = self.conn.cursor()

                # 范围投诉
                sql_select = "select `经度`, `纬度`, `故障号码` " \
                             "from `投诉-广义咨询` " \
                             "where (acos(sin((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*sin((convert(纬度,float)*3.1415)/180) " \
                             "+ cos((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*cos((convert(纬度,float)*3.1415)/180) " \
                             "* cos((" + self.lineEdit_lng.text() + "*3.1415)/180" + \
                             "-(convert(经度,float)*3.1415)/180))*6370.996)<" + radius + \
                             " and 经度!='' and  纬度!='';"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(self.lineEdit_lat.text()), float(self.lineEdit_lng.text())],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row):
                    folium.Marker(
                        location=[float(res[i][1]), float(res[i][0])],
                        popup='故障号码' + str(res[i][2]),
                        icon=folium.Icon(icon='cny', color='red')
                    ).add_to(self.m)

                # 范围4G
                sql_select = "select `经度`, `纬度`, `小区名称` from `工参-lte`" \
                             "where (acos(sin((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*sin((convert(纬度,float)*3.1415)/180) " \
                             "+ cos((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*cos((convert(纬度,float)*3.1415)/180) " \
                             "* cos((" + self.lineEdit_lng.text() + "*3.1415)/180" + \
                             "-(convert(经度,float)*3.1415)/180))*6370.996)<" + radius + \
                             " and 经度!='' and  纬度!='';"
                cursor.execute(sql_select)
                res_lte = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_lte[i][1]), float(res_lte[i][0])],
                        popup='小区名称：' + str(res_lte[i][2]),
                        icon=folium.Icon(icon='cny', color='purple')
                    ).add_to(self.m)

                # 范围5G
                sql_select = "select `GNSS经度`, `GNSS纬度`, `小区名称` from `工参-nr`" \
                             "where (acos(sin((" + self.lineEdit_lat.text() + "*3.1415)/180)" + \
                             "*sin((convert(GNSS纬度,float)*3.1415)/180) " \
                             "+ cos((" + self.lineEdit_lat.text() + "*3.1415)/180)" \
                             "*cos((convert(GNSS纬度,float)*3.1415)/180) " \
                             "* cos((" + self.lineEdit_lng.text() + "*3.1415)/180" \
                             "-(convert(GNSS经度,float)*3.1415)/180))*6370.996)<" + radius + \
                             " and GNSS经度!='' and  GNSS纬度!='';"
                cursor.execute(sql_select)
                res_nr = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_nr[i][1]), float(res_nr[i][0])],
                        popup='小区名称：' + str(res_nr[i][2]),
                        icon=folium.Icon(icon='cny', color='purple')
                    ).add_to(self.m)

                # 更新地图
                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

            except BaseException as e:
                print(e)
                QMessageBox.about(self, 'error', '查询结果为空')

    # 关闭程序
    def closeEvent(self, q_close_event):
        res = QMessageBox.question(self, "msg", "是否关闭CmCat？请确保数据正常保存。")
        if res == QMessageBox.Yes:
            self.mysql_disconnect()
            q_close_event.accept()
        else:
            q_close_event.ignore()


# 弹出新窗口显示地图
class Map(QWidget, Ui_childMap):
    def __init__(self):
        super().__init__()
        self.m = None
        self.webview = None
        self.data = None
        self.TableName = None
        self.File = None
        self.conn = None
        self.db_host = None
        self.db_port = None
        self.db_user = None
        self.db_passwd = None
        self.db_name = None
        self.db_charset = None
        self.db_addr = None
        self.setupUi(self)

        # 导入图标
        icon_cdb = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\cdb.png')
        icon_ms = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map_show.png')
        icon_bqy = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\b_qy.png')
        icon_rqy = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\r_qy.png')
        icon_mn = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map_num.png')
        icon_mll = QIcon('E:\\PyCharm\\PythonProjects\\CmCat\\CmCat_Icon\\map_ll.png')

        # 为控件设置图标
        self.pushButton_connect.setIcon(icon_cdb)
        self.pushButton_openqy.setIcon(icon_bqy)
        self.pushButton_openqy_2.setIcon(icon_rqy)
        self.pushButton_show.setIcon(icon_ms)
        self.checkBox_num.setIcon(icon_mn)
        self.checkBox_ll.setIcon(icon_mll)
        self.pushButton_show.setEnabled(False)

        # 数据库连接与地图标注点显示
        self.pushButton_connect.clicked.connect(self.mysql_connect)
        self.pushButton_openqy.clicked.connect(self.open_qy)
        self.pushButton_openqy_2.clicked.connect(self.open_qy_2)
        self.listWidget_qy.clicked.connect(self.set_to_btn_qy)
        self.listWidget_qy_2.clicked.connect(self.set_to_btn_qy_2)
        self.pushButton_show.clicked.connect(self.mark_show)

        # 地图显示
        self.layout = QVBoxLayout()
        self.setLayout(self.layout)

    # 通过函数接收传递数值
    def get_data_list(self, db_host, db_port, db_user, db_passwd, db_name, db_charset, db_addr):
        self.db_host = db_host
        self.db_port = db_port
        self.db_user = db_user
        self.db_passwd = db_passwd
        self.db_name = db_name
        self.db_charset = db_charset
        self.db_addr = db_addr
        print("Debug:" + self.db_host + ","
                       + self.db_port + ","
                       + self.db_user + ","
                       + self.db_passwd + ","
                       + self.db_name + ","
                       + self.db_charset + ","
                       + self.db_addr)

    def mysql_connect(self):
        try:
            self.conn = pymysql.connect(host=self.db_host,
                                        port=int(self.db_port),
                                        user=self.db_user,
                                        passwd=self.db_passwd,
                                        db=self.db_name,
                                        charset=self.db_charset)
        except BaseException as e:
            print(e)
            QMessageBox.about(self, "error", str(e))

        else:
            self.label_stateBar.setText("数据库连接成功！")

    # 打开查询文件
    def open_qy(self):
        self.listWidget_qy.clear()
        filepath, _ = QFileDialog.getOpenFileName(self,
                                                  '选中文件',
                                                  'E:\\PyCharm\\PythonProjects\\CmCat_Query',
                                                  'Excel files(*.xls *.xlsx)')
        self.lineEdit_qy.setText(filepath)

        try:
            # 使用xlrd通过只读方式打开文件
            self.File = xlrd.open_workbook(filepath)

            # 获取所有sheet
            self.TableName = self.File.sheet_names()
            self.listWidget_qy.addItems(self.TableName)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, "error", "打开文件未成功")

    # 启用地图显示按钮
    def set_to_btn_qy(self):
        self.pushButton_show.setEnabled(True)
        self.checkBox_num.setChecked(True)

    # 打开查询文件
    def open_qy_2(self):
        self.listWidget_qy_2.clear()
        filepath, _ = QFileDialog.getOpenFileName(self,
                                                  '选中文件',
                                                  'E:\\PyCharm\\PythonProjects\\CmCat_Query',
                                                  'Excel files(*.xls *.xlsx)')
        self.lineEdit_qy_2.setText(filepath)

        try:

            # 使用xlrd通过只读方式打开文件
            self.File = xlrd.open_workbook(filepath)

            # 获取所有sheet
            self.TableName = self.File.sheet_names()
            self.listWidget_qy_2.addItems(self.TableName)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, "error", "打开文件未成功")

    # 启用地图显示按钮
    def set_to_btn_qy_2(self):
        self.pushButton_show.setEnabled(True)
        self.checkBox_ll.setChecked(True)

    def mark_show(self):
        try:
            cursor = self.conn.cursor()

            # 选中标注方式
            if (not self.checkBox_num.isChecked()) and (not self.checkBox_ll.isChecked()):
                QMessageBox.about(self, 'tip', '请选择查询方式')

            if self.checkBox_num.isChecked() and not self.checkBox_ll.isChecked():
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ''
                for i in range(1, len(sheet_name.col_values(6)) - 1):
                    sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
                sql_select_sum = sql_select_sum + str(
                    int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
                sql_select = "select `经度`, `纬度`" \
                             "from `投诉-广义咨询` where 故障号码 in (" + sql_select_sum + ") and (经度 !='' and 纬度 != '');"
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(res[0][1]), float(res[0][0])],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row):
                    folium.Marker(
                        location=[float(res[i][1]), float(res[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='blue')
                    ).add_to(self.m)
                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

            if self.checkBox_ll.isChecked() and not self.checkBox_num.isChecked():

                # 投诉查询
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度` from `投诉-广义咨询`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res = cursor.fetchall()
                row = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(res[0][1]), float(res[0][0])],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row):
                    folium.Marker(
                        location=[float(res[i][1]), float(res[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='red')
                    ).add_to(self.m)

                # 4G宏站查询
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度` from `工参-lte`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_lte = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_lte[i][1]), float(res_lte[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='purple')
                    ).add_to(self.m)

                # 5G宏站查询
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `GNSS经度`, `GNSS纬度` from `工参-nr`" \
                               "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (GNSS经度 like '" +\
                                         str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_nr = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_nr[i][1]), float(res_nr[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='green')
                    ).add_to(self.m)

                # 更新地图
                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

            if self.checkBox_num.isChecked() and self.checkBox_ll.isChecked():

                # 号码查询投诉
                sheet_name = self.File.sheet_by_name(self.listWidget_qy.currentItem().text())
                sql_select_sum = ''
                for i in range(1, len(sheet_name.col_values(6)) - 1):
                    sql_select_sum = sql_select_sum + str(int(sheet_name.col_values(6)[i])) + ","
                sql_select_sum = sql_select_sum + str(
                    int(sheet_name.col_values(6)[len(sheet_name.col_values(6)) - 1]))
                sql_select = "select `经度`, `纬度`" \
                             "from `投诉-广义咨询` where 故障号码 in (" + sql_select_sum + ") and (经度 !='' and 纬度 != '');"
                cursor.execute(sql_select)
                res_num = cursor.fetchall()
                row_num = cursor.rowcount

                # 选取中心点
                self.gridLayout_map.removeWidget(self.webview)
                self.m = folium.Map(
                    location=[float(res_num[0][1]), float(res_num[0][0])],
                    tiles='Stamen Terrain',
                    zoom_start=14
                )
                for i in range(row_num):
                    folium.Marker(
                        location=[float(res_num[i][1]), float(res_num[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='blue')
                    ).add_to(self.m)

                # 经纬度查询投诉
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度`" \
                               "from `投诉-广义咨询`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_ll = cursor.fetchall()
                row_ll = cursor.rowcount
                for i in range(row_ll):
                    folium.Marker(
                        location=[float(res_ll[i][1]), float(res_ll[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='red')
                    ).add_to(self.m)

                # 4G宏站
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `经度`, `纬度` from `工参-lte`" \
                               "where (经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and 纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (经度 like '" + str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and 纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_lte = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_lte[i][1]), float(res_lte[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='purple')
                    ).add_to(self.m)

                # 5G宏站
                sheet_name = self.File.sheet_by_name(self.listWidget_qy_2.currentItem().text())
                sql_select_sum = ""
                sql_select_1 = "select `GNSS经度`, `GNSS纬度` from `工参-nr`" \
                               "where (GNSS经度 like '" + str(sheet_name.col_values(31)[1]) + "%'" + \
                               "and GNSS纬度 like '" + str(sheet_name.col_values(32)[1]) + "%')"
                for i in range(2, len(sheet_name.col_values(31))):
                    if str(sheet_name.col_values(31)[i]) != '' and str(sheet_name.col_values(32)[i]) != '':
                        sql_select_sum = sql_select_sum + "or (GNSS经度 like '" + \
                                         str(sheet_name.col_values(31)[i]) + "%'" + \
                                         "and GNSS纬度 like '" + str(sheet_name.col_values(32)[i]) + "%')"
                sql_select = sql_select_1 + sql_select_sum
                cursor.execute(sql_select)
                res_nr = cursor.fetchall()
                row = cursor.rowcount
                for i in range(row):
                    folium.Marker(
                        location=[float(res_nr[i][1]), float(res_nr[i][0])],
                        popup='',
                        icon=folium.Icon(icon='cny', color='green')
                    ).add_to(self.m)

                self.data = io.BytesIO()
                self.m.save(self.data, close_file=False)
                self.webview = QWebEngineView()
                self.webview.setHtml(self.data.getvalue().decode())
                self.gridLayout_map.addWidget(self.webview)

        except BaseException as e:
            print(e)
            QMessageBox.about(self, 'error', '文件表格数据不正确')


# # 弹出业务量折线图
# class Chart(QWidget, Ui_childChart):
#     def __init__(self):
#         super().__init__()
#         self.addr = None
#         self.charView = None
#         self.y_Aix = None
#         self.x_Aix = None
#         self.series_1 = None
#         self.conn = None
#         self.db_charset = None
#         self.db_name = None
#         self.db_passwd = None
#         self.db_port = None
#         self.db_user = None
#         self.db_host = None
#         self.setupUi(self)
#
#     def get_data_list(self, db_host, db_port, db_user, db_passwd, db_name, db_charset, addr):
#         self.db_host = db_host
#         self.db_port = db_port
#         self.db_user = db_user
#         self.db_passwd = db_passwd
#         self.db_name = db_name
#         self.db_charset = db_charset
#         self.conn = pymysql.connect(host=self.db_host,
#                                     port=int(self.db_port),
#                                     user=self.db_user,
#                                     passwd=self.db_passwd,
#                                     db=self.db_name,
#                                     charset=self.db_charset)
#         self.addr = addr
#         cursor_b = self.conn.cursor()
#         cursor_t = self.conn.cursor()
#         self.series_1 = QLineSeries()  # 将类QLineSeries实例化
#         sql_select_b = "select `总业务量` from `业务量-业务量-含总业务量` where 基站名称 = '" + self.addr + "'"
#         sql_select_t = "select `日期` from `业务量-业务量-含总业务量` where 基站名称 = '" + self.addr + "'"
#         cursor_b.execute(sql_select_b)
#         cursor_t.execute(sql_select_t)
#         res_b = cursor_b.fetchall()
#         res_t = cursor_t.fetchall()
#         row_b = cursor_b.rowcount
#         row_t = cursor_t.rowcount
#
#         max_num = float(res_b[0][0])
#         min_num = float(res_b[0][0])
#         for i in range(row_b):
#             self.series_1.append(i, float(res_b[i][0]))  # 折线添加坐标点清单
#             if float(res_b[i][0]) > max_num:
#                 max_num = float(res_b[i][0])
#             if float(res_b[i][0]) < min_num:
#                 min_num = float(res_b[i][0])
#         self.series_1.setName("总业务量变化折线图")  # 折线命名
#         self.x_Aix = QCategoryAxis()  # 定义y轴，实例化
#         for i in range(0, row_t):
#             self.x_Aix.append(str(res_t[i][0]), i)
#         self.x_Aix.setRange(0.00, float(row_t) - 1.00)
#         self.x_Aix.setTitleText('日期')
#         self.y_Aix = QValueAxis()  # 定义y轴，实例化
#         self.y_Aix.setRange(min_num, max_num)
#         self.y_Aix.setLabelFormat("%0.2f")
#         self.y_Aix.setTitleText('业务量（千比特）')
#         self.charView = QChartView(self.widget_chart)  # 定义charView，父窗体类型为widget
#         self.charView.setGeometry(0, 0, self.widget_chart.width(), self.widget_chart.height())  # 设置charView位置、大小
#         self.charView.chart().addSeries(self.series_1)  # 添加折线
#         self.charView.chart().setAxisX(self.x_Aix)  # 设置x轴属性
#         self.charView.chart().setAxisY(self.y_Aix)  # 设置y轴属性
#         self.charView.chart().setTitle("总业务量")  # 设置标题
#         self.charView.show()  # 显示charView


if __name__ == "__main__":
    app = QApplication(sys.argv)
    landing_panel = LandingPanel()
    landing_panel.show()
    child_map = Map()
    landing_panel.action_openMap.triggered.connect(child_map.show)
    landing_panel.action_exit.triggered.connect(app.quit)
    sys.exit(app.exec_())
    # child_chart = Chart()
    # landing_panel.pushButton_f.clicked.connect(child_chart.show)
