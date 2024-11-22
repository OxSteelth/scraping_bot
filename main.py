import sys
import time
import threading
import schedule
# ui
from ui.ui_main import Ui_MainWindow
from ui.ui_settings import Ui_Dialog as Ui_Settings_Dialog
# bot
from bot import Bot
# qt
from PyQt5.QtWidgets import *
from PyQt5.QtCore import *
from PyQt5.QtGui import *
from PyQt5.QtNetwork import *
from PyQt5.QtPrintSupport import *
# xlsx
import openpyxl


APP_NAME = "Data Scraping Bot"


def run_continuously(interval=1):
    cease_continuous_run = threading.Event()

    class ScheduleThread(threading.Thread):
        @classmethod
        def run(cls):
            while not cease_continuous_run.is_set():
                schedule.run_pending()
                time.sleep(interval)

    continuous_thread = ScheduleThread()
    continuous_thread.start()
    return cease_continuous_run


class SettingsDialog(QDialog, Ui_Settings_Dialog):

    def __init__(self, parent):
        # setup ui
        super(SettingsDialog, self).__init__(parent)
        self.setupUi(self)

        # load settings
        self.load_settings()
        # connect
        self.btn_save.clicked.connect(self.slot_save)
        self.btn_close.clicked.connect(self.close)

    def load_settings(self):
        settings = QSettings(APP_NAME, "Bot")

        # login url
        login_url = settings.value("login_url")
        if login_url:
            self.le_login_url.setText(login_url)

        # email
        email = settings.value("email")
        if email:
            self.le_email.setText(email)

        # password
        password = settings.value("password")
        if password:
            self.le_password.setText(password)

        # outlet
        outlet = settings.value("outlet")
        if outlet:
            self.le_outlet.setText(outlet)

        # target url
        target_url = settings.value("target_url")
        if target_url:
            self.le_target_url.setText(target_url)

        # date from
        date_from = settings.value("date_from")
        if date_from:
            self.de_from.setDate(date_from)

        # date to
        date_to = settings.value("date_to")
        if date_to:
            self.de_to.setDate(date_to)

        # folder id
        folder_id = settings.value("folder_id")
        if folder_id:
            self.le_folder_id.setText(folder_id)

        # execution time
        exec_time = settings.value("exec_time")
        if exec_time:
            self.te_exec_time.setTime(exec_time)

    def slot_save(self):
        # check login url
        login_url = self.le_login_url.text().strip()
        if not login_url:
            QMessageBox.warning(self, APP_NAME, "Please input login url.")
            return
        # check email
        email = self.le_email.text().strip()
        if not email:
            QMessageBox.warning(self, APP_NAME, "Please input email.")
            return
        # check password
        password = self.le_password.text().strip()
        if not password:
            QMessageBox.warning(self, APP_NAME, "Please input password.")
            return
        # check outlet
        outlet = self.le_outlet.text().strip()
        if not outlet:
            QMessageBox.warning(self, APP_NAME, "Please input outlet.")
            return
        # check target url
        target_url = self.le_target_url.text().strip()
        if not target_url:
            QMessageBox.warning(self, APP_NAME, "Please input target url.")
            return
        # check date from
        date_from = self.de_from.date()
        if not date_from:
            QMessageBox.warning(self, APP_NAME, "Please input date from.")
            return
        # check date to
        date_to = self.de_to.date()
        if not date_to:
            QMessageBox.warning(self, APP_NAME, "Please input date to.")
            return
        # check folder id
        folder_id = self.le_folder_id.text().strip()
        if not folder_id:
            QMessageBox.warning(self, APP_NAME, "Please input folder id.")
            return
        # check exec time
        exec_time = self.te_exec_time.time()
        if not exec_time:
            QMessageBox.warning(self, APP_NAME, "Please input exec time.")
            return
        # save to registry
        settings = QSettings(APP_NAME, "Bot")
        settings.setValue("login_url", login_url)
        settings.setValue("email", email)
        settings.setValue("password", password)
        settings.setValue("outlet", outlet)
        settings.setValue("target_url", target_url)
        settings.setValue("date_from", date_from)
        settings.setValue("date_to", date_to)
        settings.setValue("folder_id", folder_id)
        settings.setValue("exec_time", exec_time)

        self.accept()


class MainUi(QMainWindow, Ui_MainWindow):
    log_signal = pyqtSignal(str)
    # member variables
    settings_dialog = None
    bot = None
    m_workbook = None

    def __init__(self, parent):
        # setup ui
        super(MainUi, self).__init__(parent)
        self.setupUi(self)

        # set image
        self.lb_logo.setPixmap(QPixmap("./res/logo.png").scaled(70, 70, Qt.KeepAspectRatio, Qt.SmoothTransformation))

        # create objects
        self.settings_dialog = SettingsDialog(self)
        self.bot = Bot('', '', '', '', '', '', '', '')
        self.bot.log_signal = self.log_signal

        # init ui
        self.btn_stop.setEnabled(False)

        # connect
        self.act_settings.triggered.connect(self.slot_settings)
        self.btn_start.clicked.connect(self.slot_start)
        self.btn_stop.clicked.connect(self.slot_stop)
        self.btn_close.clicked.connect(self.slot_close)
        self.bot.log_signal.connect(self.slot_log)

    def slot_start(self):
        # check if bot settings is right
        self.update_settings()
        if not self.bot.is_settings_valid():
            QMessageBox.warning(self, APP_NAME, "Bot settings are not valid. Please check.")
            return

        # Create a workbook and add a worksheet.
        self.m_workbook = openpyxl.Workbook()

        self.btn_start.setEnabled(False)
        self.btn_stop.setEnabled(True)
        self.btn_close.setEnabled(False)

        # schedule tasks
        exec_time = self.settings_dialog.te_exec_time.time().toString('HH:mm')
        schedule.every().day.at(exec_time).do(self.bot_thread_action)
        # Start the background thread
        self.stop_run_continuously = run_continuously()
        self.bot.log("Scheduled a task at {0}".format(exec_time))

    def bot_thread_action(self):
        self.bot.extract(self.m_workbook)

    def slot_log(self, msg):
        self.te_log.append(msg)

    def slot_stop(self):
        self.btn_start.setEnabled(True)
        self.btn_stop.setEnabled(False)
        self.setFixedHeight(160)
        self.btn_close.setEnabled(True)

        self.stop_run_continuously.set()

    def slot_close(self):
        self.close()
        pass

    def update_settings(self):
        self.bot.login_url = self.settings_dialog.le_login_url.text()
        self.bot.email = self.settings_dialog.le_email.text()
        self.bot.password = self.settings_dialog.le_password.text()
        self.bot.outlet = self.settings_dialog.le_outlet.text()
        self.bot.target_url = self.settings_dialog.le_target_url.text()
        self.bot.date_from = self.settings_dialog.de_from.date().toString('dd/MMM/yyyy')
        self.bot.date_to = self.settings_dialog.de_to.date().toString('dd/MMM/yyyy')
        self.bot.folder_id = self.settings_dialog.le_folder_id.text()

    def slot_settings(self):
        if self.settings_dialog.exec_() == QDialog.Accepted:
            self.update_settings()


app = QApplication(sys.argv)
app.setWindowIcon(QIcon("./res/logo.png"))
main_ui = MainUi(None)
main_ui.show()
app.exec_()
