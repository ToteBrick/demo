#!/usr/bin/env python
# -*- coding:utf-8 -*-

from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt


class AlertDiaglog(QDialog):
    def __init__(self, *args, **kwargs):
        super(AlertDiaglog, self).__init__()
        self.field_dict = {}
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('报警设置')
        self.resize(300, 270)
        layout = QVBoxLayout()

        form_data_list = [
            {'title': 'SMTP Server', 'field': 'smtp'},
            {'title': 'Sender', 'field': 'from'},
            {'title': 'password', 'field': 'pwd'},
            {'title': 'Reciver(xx;yy)', 'field': 'to'},
        ]

        for item in form_data_list:
            lbl = QLabel()
            lbl.setText(item.get('title'))
            layout.addWidget(lbl)
            text = QLineEdit()
            self.field_dict[item.get('field')] = text
            layout.addWidget(text)

        btn_save = QPushButton('save')
        btn_save.clicked.connect(self.event_save_click)
        layout.addWidget(btn_save, 0, Qt.AlignRight)
        layout.addStretch(1)
        self.setLayout(layout)

    def event_save_click(self):
        data_dict = {}
        for key, field in self.field_dict.items():
            value = field.text().strip()
            if not value:
                QMessageBox.warning(self, 'error', 'exists empty filed!')
                return
            data_dict[key] = value
        print(data_dict)
        self.close()
        pass
