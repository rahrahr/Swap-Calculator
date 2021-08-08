import json
import re
from PyQt5 import QtWidgets, uic
import traceback
import sys
import swap_utils
import swap_calculator
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
        try:
            name = self.floating_leg.currentText()
            direction = self.trade_direction.currentText()
            nominal = float(self.nominal.text())
            trade_date = self.trade_date.text()
            maturity = self.tenor.currentText()
            first_reset_date = self.first_reset_date.text()
            fixed_rate = float(self.fixed_leg.text()) / 100
            payment_frequency = self.fixed_tenor.currentText()
            fixed_daycount = self.fixed_accrual_method.currentText()
            spread = int(self.bps.text())
            reset_frequency = self.reset_tenor.currentText()
            floating_daycount = self.floating_accrual_method.currentText()

            swap = swap_utils.create_swap(name, direction,
                                          nominal,
                                          trade_date,
                                          maturity,
                                          first_reset_date,
                                          fixed_rate,
                                          payment_frequency,
                                          fixed_daycount,
                                          spread,
                                          reset_frequency,
                                          floating_daycount)

            calc_date = self.now_date.text()
            npv, dv01 = swap_calculator.calculate_vanilla(
                swap, name, calc_date)
            self.reference_price.setText('{:.2f}'.format(npv))
            self.dv01.setText('{:.2f}'.format(dv01))

        except:
            QtWidgets.QMessageBox().about(self, '错误信息', traceback.format_exc())
            return False

    def sendSwapOrder(self):
        pass

    def to_excel(self):
        pass
