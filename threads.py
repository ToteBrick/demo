#!/usr/bin/env python
# -*- coding:utf-8 -*-
from PyQt5.QtCore import QThread, pyqtSignal


class NewTaskThreads(QThread):
    success = pyqtSignal(int, str, str, str)
    error = pyqtSignal(int, str, str, str)

    def __init__(self, current_index, asin, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.current_index = current_index
        self.asin = asin

    def run(self):
        # 触发信号
        self.success.emit(self.current_index, self.asin, 'x', 'www.baidu2.com')
        self.error.emit(self.current_index, 's', 'a', 'r')
