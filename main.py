from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLineEdit, QMessageBox, QMenuBar, QMenu, QAction, \
    QTableWidget, QHeaderView, QTableWidgetItem, QLabel
from PyQt5 import uic, QtCore
import sys
import pandas as pd


class PersonalFile(QMainWindow):
    def __init__(self, incoming_id):
        super(PersonalFile, self).__init__()
        self.id = incoming_id
        uic.loadUi('personalFile.ui', self)

        self.flmc_id = self.findChild(QLabel, 'flmc_id')
        self.flmc_id.setText('ID: ' + self.id)

        self.act_create = self.findChild(QAction, 'create_personal_file')
        self.act_create.triggered.connect(self.view_info)

        self.full_name = self.findChild(QLabel, 'full_name')
        self.birthday_date = self.findChild(QLabel, 'birthday_date')
        self.health_category = self.findChild(QLabel, 'health_category')
        self.military_specialty = self.findChild(QLabel, 'military_specialty')
        self.combat_experience = self.findChild(QLabel, 'combat_experience')

        self.passport_series = self.findChild(QLabel, 'passport_series')
        self.passport_id = self.findChild(QLabel, 'passport_id')
        self.birthday_place = self.findChild(QLabel, 'birthday_place')
        self.living_place = self.findChild(QLabel, 'living_place')
        self.family_status = self.findChild(QLabel, 'family_status')
        self.education = self.findChild(QLabel, 'education')
        self.work_place = self.findChild(QLabel, 'work_place')
        self.reference_from_work = self.findChild(QLabel, 'reference_from_work')
        self.phone_number = self.findChild(QLabel, 'phone_number')
        self.commitee_decision = self.findChild(QLabel, 'commitee_decision')

        self.father_full_name = self.findChild(QLabel, 'father_full_name')
        self.father_birthday_date = self.findChild(QLabel, 'father_birthday_date')
        self.father_birthday_place = self.findChild(QLabel, 'father_birthday_place')
        self.father_work_place = self.findChild(QLabel, 'father_work_place')
        self.mother_full_name = self.findChild(QLabel, 'mother_full_name')
        self.mother_birthday_date = self.findChild(QLabel, 'mother_birthday_date')
        self.mother_birthday_place = self.findChild(QLabel, 'mother_birthday_place')
        self.mother_work_place = self.findChild(QLabel, 'mother_work_place')

        self.visit_date = self.findChild(QLabel, 'visit_date')
        self.comissariat_address = self.findChild(QLabel, 'comissariat_address')
        self.visit_reason = self.findChild(QLabel, 'visit_reason')

        self.examination_date = self.findChild(QLabel, 'examination_date')
        self.examination_place = self.findChild(QLabel, 'examination_place')
        self.doctor_full_name = self.findChild(QLabel, 'doctor_full_name')
        self.doctor_specialty = self.findChild(QLabel, 'doctor_specialty')
        self.complaints = self.findChild(QLabel, 'complaints')
        self.anamnesis = self.findChild(QLabel, 'anamnesis')
        self.objective_research_data = self.findChild(QLabel, 'objective_research_data')
        self.examination_results = self.findChild(QLabel, 'examination_results')
        self.diagnosis = self.findChild(QLabel, 'diagnosis')

    def view_info(self):
        lfmc_row = self.get_row('lfmc.xlsx')
        self.full_name.setText(lfmc_row[1] + ' ' + lfmc_row[2] + ' ' + lfmc_row[3])
        self.birthday_date.setText('Дата рождения: ' + str(lfmc_row[4])[0:10])
        self.health_category.setText('Категория годности: ' + lfmc_row[5])
        self.military_specialty.setText('Военная специальность: ' + lfmc_row[6])
        self.combat_experience.setText('Военный опыт: ' + lfmc_row[7])

        personal_file_row = self.get_row('personal_file.xlsx')
        self.passport_series.setText('Серия: ' + str(personal_file_row[1]))
        self.passport_id.setText('Номер: ' + str(personal_file_row[2]))
        self.birthday_place.setText('Место рождения: ' + str(personal_file_row[3]))
        self.living_place.setText('Место проживания: ' + str(personal_file_row[4]))
        self.family_status.setText('Семейное положение: ' + str(personal_file_row[5]))
        self.education.setText('Образование: ' + str(personal_file_row[6]))
        self.work_place.setText('Место работы: ' + str(personal_file_row[7]))
        self.reference_from_work.setText('Характеристика с работы: ' + str(personal_file_row[8]))
        self.phone_number.setText('Номер телефона: ' + str(personal_file_row[9]))
        self.commitee_decision.setText('Решение комиссии: ' + str(personal_file_row[10]))

        parents_info_row = self.get_row('parents_info.xlsx')
        self.father_full_name.setText(str(parents_info_row[1]))
        self.father_birthday_date.setText('Дата рождения: ' + str(parents_info_row[2])[0:10])
        self.father_birthday_place.setText('Место рождения: ' + str(parents_info_row[3]))
        self.father_work_place.setText('Место работы: ' + str(parents_info_row[4]))
        self.mother_full_name.setText(str(parents_info_row[5]))
        self.mother_birthday_date.setText('Дата рождения: ' + str(parents_info_row[6])[0:10])
        self.mother_birthday_place.setText('Место рождения: ' + str(parents_info_row[7]))
        self.mother_work_place.setText('Место работы: ' + str(parents_info_row[8]))

        army_order_row = self.get_row('army_order.xlsx')
        self.visit_date.setText('Дата посещения: ' + str(army_order_row[1]))
        self.comissariat_address.setText('Адрес военкомата: ' + str(army_order_row[2]))
        self.visit_reason.setText('Причина посещения: ' + str(army_order_row[3]))

        medical_documents_row = self.get_row('medical_documents.xlsx')
        self.examination_date.setText('Дата проведения осмотра: ' + str(medical_documents_row[1])[0:10])
        self.examination_place.setText('Место проведения осмотра: ' + str(medical_documents_row[2]))
        self.doctor_full_name.setText('ФИО врача: ' + str(medical_documents_row[8]))
        self.doctor_specialty.setText('Специальность врача: ' + str(medical_documents_row[9]))
        self.complaints.setText('Жалобы: ' + str(medical_documents_row[3]))
        self.anamnesis.setText('Анамнез: ' + str(medical_documents_row[4]))
        self.objective_research_data.setText('Данные объективного исследования: ' + str(medical_documents_row[5]))
        self.examination_results.setText('Результаты осмотра: ' + str(medical_documents_row[6]))
        self.diagnosis.setText('Диагноз: ' + str(medical_documents_row[7]))

    def get_row(self, file_name):
        data_frame = pd.read_excel('bases/' + file_name)
        data_row = ''
        for row in range(0, data_frame.shape[0]):
            if str(data_frame.iloc[row][0]) == str(self.id):
                data_row = data_frame.iloc[row]
        return data_row


class AboutProgram(QMainWindow):
    def __init__(self):
        super(AboutProgram, self).__init__()
        uic.loadUi('aboutProgram.ui', self)


class Help(QMainWindow):
    def __init__(self):
        super(Help, self).__init__()
        uic.loadUi('helpUI.ui', self)


class MainTable(QMainWindow):
    def __init__(self):
        super(MainTable, self).__init__()
        uic.loadUi('mainTable.ui', self)

        self.mainTableWidget = self.findChild(QTableWidget, 'mainTableWidget')
        self.action_help = self.findChild(QAction, 'help')
        self.action_about_program = self.findChild(QAction, 'about_program')
        self.action_load = self.findChild(QAction, 'load_base')
        self.action_update = self.findChild(QAction, 'update_base')
        self.action_report = self.findChild(QAction, 'view_report')
        self.lfmc_id = self.findChild(QLabel, 'lfmc_id')

        self.action_help.triggered.connect(self.act_help)
        self.action_about_program.triggered.connect(self.act_about_program)
        self.action_load.triggered.connect(self.act_load)
        self.action_update.triggered.connect(self.act_update)
        self.action_report.triggered.connect(self.act_report)
        self.mainTableWidget.itemSelectionChanged.connect(self.cell_clicked)

        self.selected_id = '-'
        self.aboutProgram = AboutProgram()
        self.help = Help()
        self.report = PersonalFile(self.selected_id)
        self.is_loaded = False

    def act_report(self):
        if self.mainTableWidget.rowCount() != 0:
            if self.selected_id != '-':
                self.report.show()
            else:
                self.show_error('Военнообязанный не выбран')
        else:
            self.show_error('База данных не загружена')

    def act_help(self):
        self.help.show()

    def act_about_program(self):
        self.aboutProgram.show()

    def act_load(self):
        lfmc_table = pd.read_excel('bases/lfmc.xlsx')
        self.mainTableWidget.setColumnCount(len(lfmc_table.columns))
        self.mainTableWidget.setRowCount(lfmc_table.shape[0])
        column_header = ("Идентификатор", "Фамилия", "Имя", "Отчество", "Дата рождения", "Категория годности",
                         "Военная специальность", "Боевой опыт")
        self.mainTableWidget.setHorizontalHeaderLabels(column_header)
        self.mainTableWidget.verticalHeader().setVisible(False)
        self.mainTableWidget.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.is_loaded = True
        self.act_update()

    def act_update(self):
        if self.is_loaded:
            lfmc_table = pd.read_excel('bases/lfmc.xlsx')
            self.mainTableWidget.setColumnCount(len(lfmc_table.columns))
            self.mainTableWidget.setRowCount(lfmc_table.shape[0])
            for row in range(0, lfmc_table.shape[0]):
                column = 0
                for col_name, data in lfmc_table.items():
                    item = QTableWidgetItem(str(data[row]))
                    item.setTextAlignment(QtCore.Qt.AlignCenter)
                    self.mainTableWidget.setItem(row, column, item)
                    column += 1
        else:
            self.show_error('База данных не загружена')
    def cell_clicked(self):
        self.selected_id = self.mainTableWidget.model().index(self.mainTableWidget.currentRow(), 0).data()
        self.report = PersonalFile(self.selected_id)
        self.lfmc_id.setText('Выделен военнообязанный с идентификатором: ' + self.selected_id)

    def show_error(self, message):
        QMessageBox().critical(self, 'Ошибка', message, QMessageBox.Ok)


class UI(QMainWindow):
    def __init__(self):
        super(UI, self).__init__()
        uic.loadUi("LoginForm.ui", self)
        self.show()

        self.loginForm_button = self.findChild(QPushButton, 'loginForm_button')
        self.loginForm_login = self.findChild(QLineEdit, 'loginForm_login')
        self.loginForm_password = self.findChild(QLineEdit, 'loginForm_password')

        self.loginForm_button.clicked.connect(self.log_in)

        self.mainTable = MainTable()

    def log_in(self):
        logins_data = pd.read_excel('bases/registration_data.xlsx', usecols=['Login'])
        passwords_data = pd.read_excel('bases/registration_data.xlsx', usecols=['Password'])
        login = self.loginForm_login.text()
        password = self.loginForm_password.text()
        for i in range(0, logins_data.shape[0]):
            if logins_data.iloc[i][0] == login:
                if passwords_data.iloc[i][0] == password:
                    self.mainTable.show()
                    self.close()
                    break
                else:
                    self.show_error("Неверный логин или пароль")
                    break
            if i == logins_data.shape[0]-1:
                self.show_error("Неверный логин или пароль")

    def show_error(self, message):
        QMessageBox().critical(self, 'Ошибка', message, QMessageBox.Ok)


if __name__ == '__main__':
    app = QApplication(sys.argv)
    UIWindow = UI()
    app.exec_()

