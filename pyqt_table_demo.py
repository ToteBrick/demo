#!/usr/bin/env python
# -*- coding:utf-8 -*-
import sys
from PyQt5.QtWidgets import *
from PyQt5.QtGui import *
from PyQt5.QtCore import Qt


class MainWindow(QWidget):

    def __init__(self):
        super().__init__()
        self.setWindowTitle('xx system')
        self.resize(1224, 450)
        self.txt_asin = None
        # position center
        qr = self.frameGeometry()
        cp = QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)

        layout = QVBoxLayout()

        self.init_header(layout)
        # title
        self.init_title(layout)
        # table
        self.init_table(layout)
        # footer
        self.init_footer(layout)
        self.setLayout(layout)

    def init_footer(self, layout):
        footer_layout = QHBoxLayout()
        self.label_status = label_status = QLabel("未检测", self)
        footer_layout.addWidget(label_status)
        footer_layout.addStretch()
        btn_reset = QPushButton('重新初始化')
        btn_reset.clicked.connect(self.event_reset_event)
        footer_layout.addWidget(btn_reset)
        btn_recheck = QPushButton('重新检测')
        footer_layout.addWidget(btn_recheck)
        btn_delete = QPushButton('删除检测项')
        footer_layout.addWidget(btn_delete)
        btn_alert = QPushButton('SMTP报警')
        btn_alert.clicked.connect(self.event_alert_click)
        footer_layout.addWidget(btn_alert)
        btn_proxy = QPushButton('代理设置')
        footer_layout.addWidget(btn_proxy)
        layout.addLayout(footer_layout)
        # layout.addStretch()  # 弹簧，将 元素顶上去

    def init_table(self, layout):
        table_layout = QHBoxLayout()
        # self.model = QStandardItemModel(0, 8)
        table_header = [
            {'field': 'asin', 'text': 'ASIN', 'width': 120},
            {'field': 'title', 'text': '标题', 'width': 150},
            {'field': 'url', 'text': 'URL', 'width': 400},
            {'field': 'price', 'text': '底价', 'width': 100},
            {'field': 'success', 'text': '成功次数', 'width': 100},
            {'field': 'fail', 'text': '失败次数', 'width': 100},
            {'field': 'status', 'text': '状态', 'width': 100},
            {'field': 'frequency', 'text': '频率', 'width': 120},
        ]
        # 设置表头
        title_list = [i['text'] for i in table_header]
        # self.model.setHorizontalHeaderLabels(title_list)
        # self.table_widget = QTableWidget(5, 8)
        self.table_widget = QTableWidget(0, 8)
        self.table_widget.setHorizontalHeaderLabels(title_list)
        # self.tableView = QTableView()
        data_list = [
            ['bxxx', 'amd', 'www.amazon.com', 300, 0, 166, 1, 5],
            ['bxxx', 'yyy', 'www.baidu.com', 200, 0, 186, 2, 6],
        ]
        # 渲染数据
        current_row_count = self.table_widget.rowCount()  # current row
        for i, data in enumerate(data_list):
            self.table_widget.insertRow(current_row_count)
            current_row_count += 1
            for j, element in enumerate(data):
                # data_item = QStandardItem(str(element))
                # self.model.setItem(i, j, cell)
                cell = QTableWidgetItem(str(element))
                if j in [0, 4, 5, 6]:
                    cell.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)  # NO EDIT

                self.table_widget.setItem(i, j, cell)

        # self.tableView.setModel(self.model)
        # 设置宽度
        for idx, info in enumerate(table_header):
            self.table_widget.setColumnWidth(idx, info.get('width'))

        table_layout.addWidget(self.table_widget)

        print(current_row_count)
        # # # 选中的数据
        # indexs=self.tableView.selectionModel().selection().indexes()
        # print('===', indexs)

        layout.addLayout(table_layout)

        self.table_widget.setContextMenuPolicy(Qt.CustomContextMenu)  ######允许右键产生子菜单  # 设置策略为自定义菜单

        self.table_widget.customContextMenuRequested.connect(self.generateMenu)  ####右键菜单  # 菜单内容回应信号槽

    def generateMenu(self, pos):
        row_num = -1
        for i in self.table_widget.selectionModel().selection().indexes():
            row_num = i.row()

        if row_num < 2:
            menu = QMenu()
            item1 = menu.addAction(u"选项一")
            item2 = menu.addAction(u"选项二")
            item3 = menu.addAction(u"选项三")
            action = menu.exec_(self.table_widget.mapToGlobal(pos))
            if action == item1:
                print('您选了选项一，当前行文字内容是：', self.table_widget.item(row_num, 0).text(),
                      self.table_widget.item(row_num, 1).text(), self.table_widget.item(row_num, 2).text())

            elif action == item2:
                print('您选了选项二，当前行文字内容是：', self.table_widget.item(row_num, 0).text(),
                      self.table_widget.item(row_num, 1).text(), self.table_widget.item(row_num, 2).text())

            elif action == item3:
                print('您选了选项三，当前行文字内容是：', self.table_widget.item(row_num, 0).text(),
                      self.table_widget.item(row_num, 1).text(), self.table_widget.item(row_num, 2).text())
            else:
                return

    def init_title(self, layout):
        form_layout = QHBoxLayout()
        self.txt_asin = QLineEdit()  # 输入框
        self.txt_asin.setPlaceholderText("please input")
        btn_add = QPushButton("添加")
        btn_add.clicked.connect(self.event_add_click)
        form_layout.addWidget(self.txt_asin)
        form_layout.addWidget(btn_add)
        layout.addLayout(form_layout)

    def init_header(self, layout):
        header_layout = QHBoxLayout()
        btn_start = QPushButton("开始")
        # btn_start.setFixedHeight(100) # height
        btn_stop = QPushButton("停止")
        header_layout.addWidget(btn_start)
        header_layout.addWidget(btn_stop)
        header_layout.addStretch()
        layout.addLayout(header_layout)

    def event_reset_event(self):
        selected_row_list = self.table_widget.selectionModel().selectedRows()
        if not selected_row_list:
            QMessageBox.warning(self, 'error', 'not select data')
            return

        for row_obj in selected_row_list:
            index = row_obj.row()
            print('select index:', index)
            cell_status = QTableWidgetItem('init')
            cell_status.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)  #  no edit
            self.table_widget.setItem(index, 6, cell_status)

    def event_add_click(self):
        text = self.txt_asin.text()
        print(text)
        if not text:
            QMessageBox.warning(self, 'error', 'please input shop name')
        asin, price = text.split('=')
        price = float(price)
        new_row_list = [asin, "", "", price, 0, 0, 0, 5]
        current_row_count = self.table_widget.rowCount()  # current row
        self.table_widget.insertRow(current_row_count)
        for j, element in enumerate(new_row_list):
            cell = QTableWidgetItem(str(element))
            if j in [0, 4, 5, 6]:
                cell.setFlags(Qt.ItemIsSelectable | Qt.ItemIsEnabled)  # NO EDIT
            self.table_widget.setItem(current_row_count, j, cell)
        from utils.threads import NewTaskThreads
        thread = NewTaskThreads(current_row_count, asin, self)
        thread.success.connect(self.init_task_success_callback)
        thread.error.connect(self.init_task_error_callback)
        thread.start()

    def event_alert_click(self):
        from utils.diaglog import AlertDiaglog
        diaglog = AlertDiaglog()
        diaglog.setWindowModality(Qt.ApplicationModal)
        diaglog.exec_()

    def init_task_success_callback(self, index, asin, title, url):
        # 参数为thread 信号里的4个参数
        print(index, asin, title, url)
        current_title = QTableWidgetItem(title)
        self.table_widget.setItem(index, 1, current_title)

        current_url = QTableWidgetItem(url)
        self.table_widget.setItem(index, 2, current_url)
        pass

    def init_task_error_callback (self, index, asin, title, url):
        print(index, asin, title, url)
        pass


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
