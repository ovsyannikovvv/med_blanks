#!/usr/bin/env python
# coding=cp1251

from PyQt4.QtCore import *
from PyQt4.QtGui import *
import win32com.client
import pywintypes
import win32file

import sys
#reload(sys)
encoding = "cp1251"
if hasattr(sys, "setdefaultencoding"):
    sys.setdefaultencoding(encoding)

blanks_lst = [
    [0, [u"�� ���"]],
    [0, [u"���"]],
    [0, [u"��� ������"]],
    [0, [u"����"]],
    [0, [u"��� ���"]],
    [0, [u"��� ���"]],
    [0, [u"����"]],
    [0, [u"���"]],
    [0, [u"���"]],
    [0, [u"�����������"]],
    [0, [u"��� ���������� ������"]],
    [0, [u"��� �� ���������� + �\\�"]],
    [0, [u"��� �� ���� �����"]],
    [0, [u"����� ������ ����"]],
    [1, [u"������ �����", u"�����", u"���", u"���������", u"Hb", u"������������", u"������������� �������",
         u"����������� �����������"]],
    [0, [u"������ ���� �� �����������", u"� 1 �� ���� ����������"]],
    [0, [u"������ ����� �� RW"]],
    [0, [u"������������"]],
    [0, [u"�������� ������"]],
    [1, [u"����������", u"���", u"���������� �����", u"���", u"�����������",
     u"�������� �-��"]],
    [1, [u"�������������", u"����������", u"�� ���������������",
     u"������������� ������ � ��������", u"����������", u"���", u"���"]],
    [1, [u"���� �� ������������� � �������", u"1 ������", u"2 ������"]],
    [2, [u"����� �����\n+ �������� �������"]],
    [2, [u"������������� ������ �����", u"����� �����", u"�������", u"��������� �����", u"��������� ������", u"�����",
         u"����������", u"�������� ���������", u"��������", u"������� �������", u"FE-���������", u"���������",
         u"���", u"���", u"�����", u"�������", u"������"]],
    [2, [u"������������� ������ ����� ", u"���\n����\n���\n���"]],
    [3, [u"��������������", u"������� ������ � 2� ���������", u"������� ��������� �������",
        u"������ ��������� �������", u"�������� ��������", u"������� �������������� �������",
        u"������ �������������� �������", u"������������� ��������", u"������� ��������� �������",
        u"������ ��������� �������", u"�������� ��������", u"������� ��������� �������",
        u"������ ��������� �������", u"�������� ��������", u"����� ������ ����",
        u"����� ����� ����", u"������ ���", u"����� ����� ����", u"����� ������ ����",
        u"����� ����", u"������� ������ ������������", u"�������� ������ ������������",
        u"����������� ������ ������������", u"�������"]]
]


class MedBlanksSettings(QDialog):
    def __init__(self, parent=None):
        super(MedBlanksSettings, self).__init__(parent)
        self.setWindowTitle(u"���������")

        okButton = QPushButton("&OK")
        cancelButton = QPushButton(u"������")
        self.connect(okButton, SIGNAL("clicked()"), self.onOk)
        self.connect(cancelButton, SIGNAL("clicked()"),
            self, SLOT("reject()"))

        buttonLayout = QHBoxLayout()
        buttonLayout.addStretch()
        buttonLayout.addWidget(okButton)
        buttonLayout.addWidget(cancelButton)

        self.layout = QVBoxLayout()
        self.layout.addLayout(buttonLayout)
        self.setLayout(self.layout)

        save_layout = QHBoxLayout()
        path_label = QLabel(u"���� ��� ����������:")
        self.save_path_edit = QLineEdit()
        self.save_path_edit.setEnabled(False)
        change_path_button = QPushButton(u"��������")
        save_layout.addWidget(path_label)
        save_layout.addWidget(self.save_path_edit)
        save_layout.addWidget(change_path_button)

        self.layout.addLayout(save_layout)

        self.checks_dict = dict()
        self.append_cab_to_layout([u"������� �������"])
        already_added = set()
        for blank_column, blank_lst_item in blanks_lst:
            if blank_lst_item[0] not in already_added:
                self.append_cab_to_layout(blank_lst_item)
                already_added.add(blank_lst_item[0])
        self.load_settings()

        self.connect(change_path_button, SIGNAL("clicked()"), self.change_path)

    def change_path(self):
        sel_dir = QFileDialog.getExistingDirectory(self,
                                                   u"�������� ����� ��� ����������",
                                                   self.save_path_edit.text())
        sel_dir = unicode(sel_dir)
        if sel_dir:
            self.save_path_edit.setText(sel_dir)

    def append_cab_to_layout(self, blank_lst_item):
        sub_layout = QHBoxLayout()

        blank_name = blank_lst_item[0]
        blank_check = QLabel(blank_name)
        blank_edit = QLineEdit()
        sub_layout.addWidget(blank_check)
        sub_layout.addWidget(blank_edit)
        self.layout.addLayout(sub_layout)
        self.checks_dict[blank_name] = blank_edit

    def load_settings(self):
        settings = QSettings()
        def_save_path = unicode(QDir.tempPath()).replace("/", "\\")
        save_val = settings.value("save_path")
        if save_val:
            sss = save_val.toString()
            if sss:
                self.save_path_edit.setText(sss)
            else:
                self.save_path_edit.setText(def_save_path)
        else:
            self.save_path_edit.setText(def_save_path)

        for checkname, checkedit in self.checks_dict.items():
            val = settings.value(checkname)
            if val:
                checkedit.setText(val.toString())

    def save_settings(self):
        settings = QSettings()
        settings.setValue("save_path", self.save_path_edit.text())
        for checkname, checkedit in self.checks_dict.items():
            settings.setValue(checkname, checkedit.text())

    def onOk(self):
        self.save_settings()
        self.accept()


class MedBlanksUI(QWidget):
    def __init__(self, parent=None):
        super(MedBlanksUI, self).__init__(parent)

        # ���������
        enter_data_layout = QVBoxLayout()

        data_layout = QGridLayout()
        fio_label = QLabel(u"���")
        self.fio_edit = QLineEdit(u"")
        date_label = QLabel(u"�\\�")
        self.date_edit = QLineEdit(u"")
        m_label = QLabel(u"�\\�")
        self.m_edit = QLineEdit(u"")
        polis_label = QLabel(u"�����")
        self.polis_edit = QLineEdit(u"")
        diagnose_label = QLabel(u"�������")
        self.diagnose_edit = QLineEdit(u"")
        patient_label = QLabel(u"� ��������")
        self.patient_edit = QLineEdit(u"")

        data_layout.addWidget(fio_label, 0, 0)
        data_layout.addWidget(self.fio_edit, 0, 1)

        data_layout.addWidget(date_label, 0, 2)
        data_layout.addWidget(self.date_edit, 0, 3)
        data_layout.addWidget(m_label, 1, 0)
        data_layout.addWidget(self.m_edit, 1, 1)

        data_layout.addWidget(polis_label, 1, 2)
        data_layout.addWidget(self.polis_edit, 1, 3)
        data_layout.addWidget(diagnose_label, 2, 0)
        data_layout.addWidget(self.diagnose_edit, 2, 1)
        data_layout.addWidget(patient_label, 2, 2)
        data_layout.addWidget(self.patient_edit, 2, 3)
        enter_data_layout.addLayout(data_layout)

        blanks_layouts = [QVBoxLayout(), QVBoxLayout(), QVBoxLayout(), QVBoxLayout()]

        self.checks_lst = []
        for blank_column, blank_lst_item in blanks_lst:
            blank_name = blank_lst_item[0]
            if len(blank_lst_item) == 1:
                blank_check = QCheckBox(blank_name)
                self.checks_lst.append([blank_check])
                blanks_layouts[blank_column].addWidget(blank_check)
            else:
                grouper = QGroupBox(blank_name)
                grouper_layout = QVBoxLayout()
                # buttons_layout = QHBoxLayout()
                # button_plus = QPushButton("+")
                # button_minus = QPushButton("-")
                # buttons_layout.addWidget(button_plus)
                # buttons_layout.addWidget(button_minus)
                # grouper_layout.addLayout(buttons_layout)

                sub_checks_lst = [grouper]
                for subcheck in blank_lst_item[1:]:
                    blank_check = QCheckBox(subcheck)
                    sub_checks_lst.append(blank_check)
                    grouper_layout.addWidget(blank_check)
                grouper.setChecked(False)
                grouper.setLayout(grouper_layout)

                self.checks_lst.append(sub_checks_lst)
                blanks_layouts[blank_column].addWidget(grouper)

        layout = QHBoxLayout()
        for lay in blanks_layouts:
            lay.addStretch()
            layout.addLayout(lay)
        enter_data_layout.addLayout(layout)

        buts_layout = QVBoxLayout()
        self.settings_button = QPushButton(u"\n���������\n")
        buts_layout.addWidget(self.settings_button)
        self.create_button = QPushButton(u"\n�������\n")
        buts_layout.addWidget(self.create_button)
        layout.addLayout(buts_layout)

        self.setLayout(enter_data_layout)
        self.clear_ui()

        self.connect(self.settings_button, SIGNAL("clicked()"), self.show_settings)
        self.connect(self.create_button, SIGNAL("clicked()"), self.create_blanks)

        self.update_checks_withs_cabs()
        QTimer.singleShot(0, self.center)

    def show_settings(self):
        dialog = MedBlanksSettings()
        if dialog.exec_():
            self.update_checks_withs_cabs()

    def group_toggled(self):
        pass

    def center(self):
        frame_gm = self.frameGeometry()
        screen = QApplication.desktop().screenNumber(QApplication.desktop().cursor().pos())
        center_point = QApplication.desktop().screenGeometry(screen).center()
        frame_gm.moveCenter(center_point)
        self.move(frame_gm.topLeft())

    def fill_cell(self, cell, doc, blank):
        cell.Range.Select()
        cell.Range.Text = ""

        blnk_name, blnk_cab = self.parse_check_name(blank[0])

        doc.ActiveWindow.Selection.Font.Size = 12
        #doc.ActiveWindow.Selection.TypeText(u"������� �" + blnk_cab)
        doc.ActiveWindow.Selection.TypeText(blnk_cab)
        doc.ActiveWindow.Selection.TypeParagraph()
        doc.ActiveWindow.Selection.Font.Size = 14
        doc.ActiveWindow.Selection.TypeText(blnk_name)
        doc.ActiveWindow.Selection.ParagraphFormat.Alignment = 1
        doc.ActiveWindow.Selection.TypeParagraph()
        doc.ActiveWindow.Selection.Font.Size = 12
        doc.ActiveWindow.Selection.ParagraphFormat.Alignment = 0
        # doc.ActiveWindow.Selection.TypeText(u"���, �\�")
        doc.ActiveWindow.Selection.TypeText(unicode(self.fio_edit.text() + " " + self.date_edit.text()))
        doc.ActiveWindow.Selection.TypeParagraph()
        # doc.ActiveWindow.Selection.TypeText(u"�\�, �����")
        doc.ActiveWindow.Selection.TypeText(unicode(self.m_edit.text() + " " + self.polis_edit.text()))
        doc.ActiveWindow.Selection.TypeParagraph()
        #doc.ActiveWindow.Selection.TypeText(u"�������")
        doc.ActiveWindow.Selection.TypeText(unicode(self.diagnose_edit.text()))
        doc.ActiveWindow.Selection.TypeParagraph()
        #doc.ActiveWindow.Selection.TypeText(u"� ��������")
        doc.ActiveWindow.Selection.TypeText(u"�"+unicode(self.patient_edit.text()))
        doc.ActiveWindow.Selection.TypeParagraph()
        doc.ActiveWindow.Selection.TypeParagraph()
        doc.ActiveWindow.Selection.Font.Size = 14
        for subblank in blank[1]:
            doc.ActiveWindow.Selection.TypeText(unicode(subblank))
            doc.ActiveWindow.Selection.TypeParagraph()

        #doc.ActiveWindow.Selection.TypeParagraph()
        #doc.ActiveWindow.Selection.TypeParagraph()
        #doc.ActiveWindow.Selection.TypeText(self.cab)

    def get_selected_blanks_info(self):
        ret_dict = dict()

        for check_lst in self.checks_lst:
            try:
                blankName = check_lst[0].text()
            except AttributeError:
                blankName = check_lst[0].title()

            if len(check_lst) == 1:
                if check_lst[0].isChecked():
                    ret_dict[blankName] = []
            else:
                for sub_check in check_lst[1:]:
                    if sub_check.isChecked():
                        ret_dict.setdefault(blankName, []).append(sub_check.text())
                        # if blankName not in ret_dict:
                        #     ret_dict[blankName] = []
                        # ret_dict[blankName].append(sub_check.text())

        return ret_dict

    def update_checks_withs_cabs(self):
        settings = QSettings()
        cab = settings.value(u"������� �������")
        self.cab = cab.toString()
        self.setWindowTitle(u"������� �" + self.cab)
        self.save_path = unicode(QDir.tempPath()).replace("/", "\\")
        save_val = settings.value("save_path")
        if save_val:
            sss = save_val.toString()
            if sss:
                self.save_path = sss

        for check_lst in self.checks_lst:
            # ��������� ������� ���������
            check_name = self.parse_check_name(check_lst[0])[0]
            cab = settings.value(check_name)
            check_name += " (" + cab.toString() + ")"

            try:
                check_lst[0].setText(check_name)
            except AttributeError:
                check_lst[0].setTitle(check_name)

    def clear_ui(self):
        self.fio_edit.setText(u"")
        self.date_edit.setText(u"")
        self.m_edit.setText(u"")
        self.polis_edit.setText(u"")
        self.diagnose_edit.setText(u"")
        self.patient_edit.setText(u"")

        for check_lst in self.checks_lst:
            for check in check_lst:
                try:
                    check.setChecked(False)
                except:
                    pass

    def parse_check_name(self, check_or_group):
        try:
            check_name = check_or_group.text()
        except AttributeError:
            try:
                check_name = check_or_group.title()
            except AttributeError:
                check_name = check_or_group

        parted = unicode(check_name).partition(" (")
        retName = parted[0]
        retCab = ""
        if parted[2]:
            retCab = parted[2].rstrip(")")
        return retName, retCab

    def create_blanks(self):
        sel_blanks = self.get_selected_blanks_info()
        if not sel_blanks:
            QMessageBox.warning(self, "MedBlanks", u"�� ������� ������")
        else:
            tmpdir = unicode(self.save_path)
            guid = unicode(pywintypes.CreateGuid())
            tempFileName = tmpdir + "/" + unicode(self.fio_edit.text()) + "_" + guid + ".docx"
            tempFileName = unicode(tempFileName).replace("/", "\\").encode("cp1251")
            win32file.CopyFile("blank.docx", tempFileName, 0)

            wordObject = win32com.client.Dispatch("Word.Application")
            wordObject.Visible = 1

            openedDoc = wordObject.Documents.Open(tempFileName)
            table = openedDoc.Tables(1)
            table.Rows.HeightRule = 1

            # ���������� ��� ��������������...
            sel_blanks_list = []
            rentgstr = u"��������������"
            for blank_item in sel_blanks.items():
                if unicode(blank_item[0]).startswith(rentgstr):
                    sel_blanks_list.insert(0, ("empty", 0))
                    sel_blanks_list.insert(0, blank_item)
                else:
                    sel_blanks_list.append(blank_item)

            rows_count = (len(sel_blanks_list) + 1) / 2
            for i in xrange(rows_count - 1):
                table.Rows.Add()
                table.Rows.Add()

            for i, blank in enumerate(sel_blanks_list):
                row = (i/2)*2 + 1
                table.Rows(row + 1).Borders(1).LineStyle = 0
                table.Rows(row).Height = 28.35 * 11
                table.Rows(row + 1).Height = 28.35
                if blank[0] != "empty":
                    self.fill_cell(table.Cell(row, i % 2 + 1), openedDoc, blank)
                    table.Cell(row + 1, i % 2 + 1).Range.Text = unicode(self.cab)
                else:
                    table.Cell(row, i % 2 + 1).Borders(2).LineStyle = 0
                    table.Cell(row + 1, i % 2 + 1).Borders(2).LineStyle = 0

            self.clear_ui()

            openedDoc.Save()


app = QApplication(sys.argv)
app.setOrganizationName("OvsyannikovVV")
app.setApplicationName("MedBlanks")
form = MedBlanksUI()
form.show()
app.exec_()
