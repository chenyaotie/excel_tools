# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'untitled.ui'
#
# Created by: PyQt4 UI code generator 4.11.4
#
# WARNING! All changes made in this file will be lost!

import os
import threading
import time
from PyQt4 import QtCore, QtGui
from PyQt4.QtCore import SIGNAL, QString
from PyQt4.QtGui import QWidget, QFileDialog

from excel.process import process
from excel.util.log import get_logger
from excel.config.config import Config

try:
    _fromUtf8 = QtCore.QString.fromUtf8
except AttributeError:
    def _fromUtf8(s):
        return s

try:
    _encoding = QtGui.QApplication.UnicodeUTF8


    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig, _encoding)
except AttributeError:
    def _translate(context, text, disambig):
        return QtGui.QApplication.translate(context, text, disambig)

is_run = False

LOG = get_logger(__name__)


class Ui_Dialog(QWidget):
    def __init__(self):
        QWidget.__init__(self)
        self.config = Config()

    def setupUi(self, Dialog):
        Dialog.setObjectName(_fromUtf8(u"财务报表统计"))
        Dialog.resize(632, 466)
        self.groupBox = QtGui.QGroupBox(Dialog)
        self.groupBox.setGeometry(QtCore.QRect(0, 0, 553, 210))
        self.groupBox.setObjectName(_fromUtf8("groupBox"))
        self.formLayoutWidget = QtGui.QWidget(self.groupBox)
        self.formLayoutWidget.setGeometry(QtCore.QRect(10, 20, 500, 191))
        self.formLayoutWidget.setObjectName(_fromUtf8("formLayoutWidget"))
        self.formLayout = QtGui.QFormLayout(self.formLayoutWidget)
        self.formLayout.setFieldGrowthPolicy(
            QtGui.QFormLayout.AllNonFixedFieldsGrow)
        self.formLayout.setContentsMargins(-1, 0, -1, -1)
        self.formLayout.setHorizontalSpacing(6)
        self.formLayout.setObjectName(_fromUtf8("formLayout"))

        self.pushButton = QtGui.QPushButton(Dialog)
        self.pushButton.setGeometry(QtCore.QRect(555, 180, 75, 23))
        self.pushButton.setObjectName(_fromUtf8("pushButton"))
        self.pushButton.clicked.connect(self.run)

        self.timesheetLabel = QtGui.QLabel(self.formLayoutWidget)
        self.timesheetLabel.setObjectName(_fromUtf8("timesheetLabel"))
        self.formLayout.setWidget(
            0, QtGui.QFormLayout.LabelRole, self.timesheetLabel)
        self.timesheetLineEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.timesheetLineEdit.setEnabled(True)
        self.timesheetLineEdit.setObjectName(_fromUtf8("timesheetLineEdit"))
        self.formLayout.setWidget(
            0,
            QtGui.QFormLayout.FieldRole,
            self.timesheetLineEdit)

        self.Li_run_zhong_xin_label = QtGui.QLabel(self.formLayoutWidget)
        self.Li_run_zhong_xin_label.setObjectName(
            _fromUtf8("Li_run_zhong_xin_label"))
        self.formLayout.setWidget(1,
                                  QtGui.QFormLayout.LabelRole,
                                  self.Li_run_zhong_xin_label)
        self.Li_Run_ReportEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.Li_Run_ReportEdit.setObjectName(_fromUtf8("LineEdit"))
        self.formLayout.setWidget(
            1,
            QtGui.QFormLayout.FieldRole,
            self.Li_Run_ReportEdit)

        self.proeject_query_label = QtGui.QLabel(self.formLayoutWidget)
        self.proeject_query_label.setObjectName(
            _fromUtf8("proeject_query_label"))
        self.formLayout.setWidget(
            2,
            QtGui.QFormLayout.LabelRole,
            self.proeject_query_label)
        self.proeject_queryEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.proeject_queryEdit.setObjectName(_fromUtf8("LineEdit_2"))
        self.formLayout.setWidget(
            2,
            QtGui.QFormLayout.FieldRole,
            self.proeject_queryEdit)

        self.bCLabel = QtGui.QLabel(self.formLayoutWidget)
        self.bCLabel.setObjectName(_fromUtf8("bCLabel"))
        self.formLayout.setWidget(3, QtGui.QFormLayout.LabelRole, self.bCLabel)
        self.bCLineEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.bCLineEdit.setObjectName(_fromUtf8("bCLineEdit"))
        self.formLayout.setWidget(
            3, QtGui.QFormLayout.FieldRole, self.bCLineEdit)

        self.formLayout_3 = QtGui.QFormLayout()
        self.formLayout_3.setObjectName(_fromUtf8("formLayout_3"))
        self.bCSheetLabel = QtGui.QLabel(self.formLayoutWidget)
        self.bCSheetLabel.setObjectName(_fromUtf8("bCSheetLabel"))
        self.formLayout_3.setWidget(
            0, QtGui.QFormLayout.LabelRole, self.bCSheetLabel)

        self.bCSheetLineEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.bCSheetLineEdit.setInputMask(_fromUtf8(""))
        self.bCSheetLineEdit.setEchoMode(QtGui.QLineEdit.Normal)
        self.bCSheetLineEdit.setAlignment(QtCore.Qt.AlignCenter)
        self.bCSheetLineEdit.setObjectName(_fromUtf8("bCSheetLineEdit"))
        self.formLayout_3.setWidget(
            0, QtGui.QFormLayout.FieldRole, self.bCSheetLineEdit)

        self.formLayout.setLayout(
            4, QtGui.QFormLayout.FieldRole, self.formLayout_3)
        self.weightLabel = QtGui.QLabel(self.formLayoutWidget)
        self.weightLabel.setObjectName(_fromUtf8("weightLabel"))
        self.formLayout.setWidget(
            5, QtGui.QFormLayout.LabelRole, self.weightLabel)
        self.weightEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.weightEdit.setObjectName(_fromUtf8("weight"))
        self.formLayout.setWidget(
            5, QtGui.QFormLayout.FieldRole, self.weightEdit)

        self.formLayout.setLayout(
            6, QtGui.QFormLayout.FieldRole, self.formLayout_3)
        self.outputpathLabel = QtGui.QLabel(self.formLayoutWidget)
        self.outputpathLabel.setObjectName(_fromUtf8("outputpathLabel"))
        self.formLayout.setWidget(
            7, QtGui.QFormLayout.LabelRole, self.outputpathLabel)
        self.outputpathEdit = QtGui.QLineEdit(self.formLayoutWidget)
        self.outputpathEdit.setObjectName(_fromUtf8("LineEdit_3"))
        self.formLayout.setWidget(
            7, QtGui.QFormLayout.FieldRole, self.outputpathEdit)
        self.groupBox_2 = QtGui.QGroupBox(Dialog)
        self.groupBox_2.setGeometry(QtCore.QRect(5, 210, 620, 250))

        self.groupBox_2.setObjectName(_fromUtf8("groupBox_2"))
        self.textBrowser = QtGui.QTextBrowser(self.groupBox_2)
        self.textBrowser.setGeometry(QtCore.QRect(5, 14, 610, 232))
        self.textBrowser.setObjectName(_fromUtf8("textBrowser"))

        self.toolButton_time_sheet = QtGui.QToolButton(Dialog)
        self.toolButton_time_sheet.setGeometry(QtCore.QRect(510, 19, 40, 20))
        self.toolButton_time_sheet.setObjectName(
            _fromUtf8("toolButton_time_sheet"))
        self.toolButton_lirun = QtGui.QToolButton(Dialog)
        self.toolButton_lirun.setGeometry(QtCore.QRect(510, 45, 40, 20))
        self.toolButton_lirun.setObjectName(_fromUtf8("toolButton_lirun"))
        self.toolButton_project_qurey = QtGui.QToolButton(Dialog)
        self.toolButton_project_qurey.setGeometry(
            QtCore.QRect(510, 74, 40, 20))
        self.toolButton_project_qurey.setObjectName(
            _fromUtf8("toolButton_project_qurey"))
        self.toolButton_bc = QtGui.QToolButton(Dialog)
        self.toolButton_bc.setGeometry(QtCore.QRect(510, 100, 40, 20))
        self.toolButton_bc.setObjectName(_fromUtf8("toolButton_bc"))

        self.toolButton_weight = QtGui.QToolButton(Dialog)
        self.toolButton_weight.setGeometry(QtCore.QRect(510, 150, 40, 20))
        self.toolButton_weight.setObjectName(_fromUtf8("toolButton_weight"))

        #
        self.toolButton_output = QtGui.QToolButton(Dialog)
        self.toolButton_output.setGeometry(QtCore.QRect(510, 180, 40, 20))
        self.toolButton_output.setObjectName(_fromUtf8("toolButton_output"))

        self.connect(
            self.toolButton_time_sheet,
            SIGNAL("clicked()"),
            lambda: self.brower_file(
                self.timesheetLineEdit))
        self.connect(
            self.toolButton_lirun,
            SIGNAL("clicked()"),
            lambda: self.brower_file(
                self.Li_Run_ReportEdit))
        self.connect(self.toolButton_project_qurey, SIGNAL("clicked()"),
                     lambda: self.brower_file(self.proeject_queryEdit))
        self.connect(
            self.toolButton_bc,
            SIGNAL("clicked()"),
            lambda: self.brower_file(
                self.bCLineEdit))
        self.connect(
            self.toolButton_weight,
            SIGNAL("clicked()"),
            lambda: self.brower_file(
                self.weightEdit))
        self.connect(
            self.toolButton_output,
            SIGNAL("clicked()"),
            lambda: self.brower_dir(
                self.outputpathEdit))

        self.retranslateUi(Dialog)

        QtCore.QMetaObject.connectSlotsByName(Dialog)

    def retranslateUi(self, Dialog):
        Dialog.setWindowTitle(_translate("Dialog", "财务报表统计", None))
        self.groupBox.setTitle(_translate("Dialog", "路径输入框", None))
        self.timesheetLabel.setText(_translate("Dialog", "TimeSheet报表", None))
        self.Li_run_zhong_xin_label.setText(
            _translate("Dialog", "利润中心报表", None))
        self.proeject_query_label.setText(
            _translate("Dialog", "项目综合查询报表", None))
        self.bCLabel.setText(_translate("Dialog", "BC报表", None))
        self.bCSheetLabel.setText(
            _translate(
                "Dialog",
                "BC报表中的sheet页名称                ",
                None))
        self.bCSheetLineEdit.setText(_translate("Dialog", "2019-07", None))
        self.weightLabel.setText(_translate("Dialog", "权重表", None))
        self.outputpathLabel.setText(_translate("Dialog", "汇总后输出路径", None))
        self.groupBox_2.setTitle(_translate("Dialog", "输出", None))
        self.pushButton.setText(_translate("Dialog", "Run", None))
        self.toolButton_time_sheet.setText(_translate("Dialog", "...", None))
        self.toolButton_lirun.setText(_translate("Dialog", "...", None))
        self.toolButton_project_qurey.setText(
            _translate("Dialog", "...", None))
        self.toolButton_bc.setText(_translate("Dialog", "...", None))
        self.toolButton_weight.setText(_translate("Dialog", "...", None))
        self.toolButton_output.setText(_translate("Dialog", "...", None))

    def initial_config(self):
        self.bc_sheet_name = unicode(self.bCSheetLineEdit.text())
        self.bc_report_path = unicode(self.bCLineEdit.text())
        self.li_run_report_path = unicode(self.Li_Run_ReportEdit.text())
        self.proeject_query_report_path = unicode(
            self.proeject_queryEdit.text())
        self.timesheet_report_path = unicode(self.timesheetLineEdit.text())
        self.weight_path = unicode(self.weightEdit.text())
        self.output_path = os.path.join(
            unicode(self.outputpathEdit.text()), "result.xls")

        self.config.set_timesheet_report_path(self.timesheet_report_path)
        self.config.set_bc_report_path(self.bc_report_path)
        self.config.set_bc_sheet_name(self.bc_sheet_name)
        self.config.set_profit_report_path(self.li_run_report_path)
        self.config.set_project_query_report_path(self.proeject_query_report_path)
        self.config.set_weight_path(self.weight_path)
        self.config.set_output_path(self.output_path)
        self.config.set_textBrowser(self.textBrowser)

    def brower_file(self, line_edit):
        path = QFileDialog.getOpenFileName(self, 'Open file',
                                           '.', "*")
        if path:
            line_edit.setText(unicode(path))

    def brower_dir(self, line_edit):
        dir_path = QFileDialog.getExistingDirectory(
            self, "select directory ", "/")
        if dir_path:
            line_edit.setText(dir_path)

    def is_empty(self, path):
        if not path:
            return False
        return True

    def run(self):

        if self.get_is_run():
            self.textBrowser.clear()
            time.sleep(0.2)
            self.textBrowser.insertPlainText(u"正在处理，请稍等...\n")
            return

        self.initial_config()
        result = map(self.is_empty,
                     [self.proeject_query_report_path,
                      self.bc_report_path,
                      self.timesheet_report_path,
                      self.li_run_report_path,
                      self.bc_sheet_name,
                      self.output_path])
        if False in result:
            msg = u"输入框不能为空\n"
            self.textBrowser.insertPlainText(msg)
            LOG.info(msg)
            return


        t = threading.Thread(target=self.caculate)
        t.setDaemon(True)
        t.start()
        self.textBrowser.clear()
        self.textBrowser.insertPlainText(u"请稍后...\n")
        self.set_is_run(True)

    def caculate(self):
        try:
            main_process = process.Caculating()
            main_process.update()

            self.textBrowser.append("completed.")
            self.textBrowser.append(u"文件已经输出到：%s" % self.output_path)
            LOG.info(u"计算完成，文件路径：%s" % self.output_path)
        except Exception as e:
            LOG.info(e)
            if type(e) != unicode or type(e)!=str:
                e=str(e)
            self.textBrowser.append(e)
            self.set_is_run(False)

    def set_is_run(self, boolean):
        global is_run
        is_run = boolean

    def get_is_run(self):
        global is_run
        return is_run


if __name__ == "__main__":
    import sys

    app = QtGui.QApplication(sys.argv)
    Dialog = QtGui.QDialog()
    ui = Ui_Dialog()
    ui.setupUi(Dialog)
    Dialog.show()
    app.exec_()
    # sys.exit(app.exec_())
