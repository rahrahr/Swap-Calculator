import json
import re
from PyQt5 import QtWidgets, uic
import traceback
import sys
sys.path.append("..")


class SwapUi(QtWidgets.QMainWindow):
    def __init__(self):
        super(SwapUi, self).__init__()
        uic.loadUi("swap.ui", self)
        self.setCentralWidget(self.widget)
        self.centralWidget().layout().setContentsMargins(*([50]*4))
        self.send_swap_order.clicked.connect(self.sendSwapOrder)
        self.pushButton.clicked.connect(self.calculate)

        accounts = json.load(open('trader.json', encoding='utf-8'))
        self.list_type.clear()
        self.list_type.addItems(accounts.keys())
        self.list_type.currentTextChanged.connect(self.on_list_type_change)

        self.account_list.clear()
        key = self.list_type.currentText()
        value = accounts[key]
        self.account_list.addItems(value)

        accounts = json.load(open('counterparty.json', encoding='utf-8'))
        self.list_type_2.clear()
        self.list_type_2.addItems(accounts.keys())
        self.list_type_2.currentTextChanged.connect(self.on_list_type_change_2)

        self.account_list_2.clear()
        key = self.list_type_2.currentText()
        value = accounts[key]
        self.account_list_2.addItems(value)

    def on_list_type_change(self):
        self.account_list.clear()
        key = self.list_type.text()
        value = json.load(open('trader.json', encoding='utf-8'))[key].keys()
        self.account_list.addItems(value)

    def on_list_type_change_2(self):
        self.account_list_2.clear()
        key = self.list_type_2.text()
        value = json.load(
            open('counterparty.json', encoding='utf-8'))[key].keys()
        self.account_list.addItems(value)

    def calculate(self):
        pass

    def sendSwapOrder(self):
        pass

    def to_excel(self):
        pass