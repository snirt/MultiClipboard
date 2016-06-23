import sys
from datetime import datetime

from PySide import QtCore, QtGui
from PySide.QtCore import QRegExp
from PySide.QtGui import QIcon, QMessageBox, QStyle
from Ui_MultiClipboard import Ui_MainWindow
import pyperclip
import sqlite3 as lite

# CONSTANCE
APP_NAME = 'MultiClipboard'
APP_VERSION = '0.7b'
APP_ICON = 'MultiClipboard.png'


class MainWindow(QtGui.QMainWindow):

    def __init__(self, parent=None):
        QtGui.QMainWindow.__init__(self, parent)
        self.ui = Ui_MainWindow()
        self.ui.setupUi(self)

        # General Settings
        self.setWindowTitle(APP_NAME + ' ' + APP_VERSION)
        self.systemTrayIcon = QtGui.QSystemTrayIcon(self)
        self.setWindowIcon(QIcon(APP_ICON))
        self.systemTrayIcon.setIcon(QIcon(APP_ICON))
        self.systemTrayIcon.setVisible(True)
        self.systemTrayIcon.activated.connect(self.on_system_tray_icon_activated)
        self.setWindowFlags(self.windowFlags() & QtCore.Qt.CustomizeWindowHint)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowMinMaxButtonsHint)

        # Standard item model
        self.item_model = QtGui.QStandardItemModel(0, 3)
        self.item_model.setHorizontalHeaderLabels(['ID', 'DATE', 'VALUE'])

        # Filter proxy model
        self.filter_proxy_model = QtGui.QSortFilterProxyModel()
        self.filter_proxy_model.setSourceModel(self.item_model)
        self.filter_proxy_model.setFilterKeyColumn(2)
        self.filter_proxy_model.setFilterRegExp(QRegExp("", QtCore.Qt.CaseInsensitive, QRegExp.FixedString))

        # line edit filter
        line_edit = self.ui.findLineEdit
        line_edit.textChanged.connect(self.filter_proxy_model.setFilterRegExp)

        # connect table with model and change settings
        self.ui.cb_table.horizontalHeader().setStretchLastSection(True)
        self.ui.cb_table.setModel(self.filter_proxy_model)
        self.ui.cb_table.setColumnWidth(1, 115)
        self.ui.cb_table.doubleClicked.connect(self.mousePressEvent)

        self.clipboard = str(pyperclip.paste()).encode("utf-8")
        self.lastContent = self.clipboard
        self.always_on_top = ""

        # Conncet to db

        try:
            self.conn = lite.connect(r"mc_db1.db")
            self.cur = self.conn.cursor()
            self.check_db()
            self.initialize()
        except lite.Error as e:
            print("Error:" + e.args[0])
            sys.exit(1)
        self.db_to_table()

        # Hide col ID
        self.ui.cb_table.setColumnHidden(0, True)

        # Set UI buttons
        btn_clear = self.ui.clearButton
        btn_clear.setIcon(self.style().standardIcon(QStyle.SP_DialogResetButton))
        btn_clear.clicked.connect(self.clear_table)

        btn_delete_selected = self.ui.btn_delete_selected
        btn_delete_selected.setIcon(self.style().standardIcon(QStyle.SP_DialogCloseButton))
        btn_delete_selected.clicked.connect(self.delete_selected_rows)

        btn_merge_selected = self.ui.btn_merge_selected
        btn_merge_selected.setIcon(self.style().standardIcon(QStyle.SP_DialogSaveButton))
        btn_merge_selected.clicked.connect(self.copy_selected_rows)

        # Menu > File
        menu_reloadProps = self.ui.actionReload_properties
        menu_reloadProps.setIcon(self.style().standardIcon(QStyle.SP_BrowserReload))
        menu_reloadProps.triggered.connect(self.initialize)

        buttonExit = self.ui.actionExit
        buttonExit.setIcon(self.style().standardIcon(QStyle.SP_DialogCancelButton))
        buttonExit.triggered.connect(self.exit_program)

        # Menu > Edit
        menu_clearcont = self.ui.actionClear_content
        menu_clearcont.setIcon(self.style().standardIcon(QStyle.SP_DialogResetButton))
        menu_clearcont.triggered.connect(self.clear_table)

        # Menu > Help
        menu_about = self.ui.actionAbout
        menu_about.setIcon(self.style().standardIcon(QStyle.SP_MessageBoxInformation))
        menu_about.triggered.connect(self.event_about)

        # Clipboard Listener
        self.timer = QtCore.QTimer(self)
        self.timer.setInterval(500)
        self.timer.timeout.connect(self.interval)
        self.timer.start()

        # alwaysOnTop Checkbox
        self.alwaysOnTop = self.ui.alwaysOnTop_check
        self.alwaysOnTop.setChecked(self.get_always_on_top_status())
        self.alwaysOnTop.stateChanged.connect(self.always_on_top_toggle)

    def exit_program(self):
        QtGui.QApplication.quit()


    def mousePressEvent(self):
        rownum = self.ui.cb_table.selectionModel().currentIndex().row()
        index = self.item_model.index(rownum, 2)
        print("COPY TO CLIPBOARD" + str(index.data()))
        pyperclip.copy(str(index.data()))

    def interval(self):
        self.clipboard = str(pyperclip.paste()).encode("utf-8")
        if self.lastContent != self.clipboard:
            self.insert_from_clipboard()
            self.lastContent = self.clipboard


    def initialize(self):
        # Load properties from DB

        # Always on top
        if self.get_always_on_top_status():
            self.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        else:
            self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
        self.show()

    def on_system_tray_icon_activated(self, reason):
        if reason == QtGui.QSystemTrayIcon.DoubleClick:
            if self.isHidden():
                self.show()
            else:
                self.hide()

    @QtCore.Slot()
    def closeEvent(self, event):
        # if self.clipboard_listener.isAlive:
        #     self.clipboard_listener.stop()
        #     time.sleep(0.3)
        # sys.exit()
        event.ignore()
        self.hide()
        self.systemTrayIcon.showMessage(APP_NAME, APP_NAME + ' is running in background.')

    @QtCore.Slot()
    def insert_from_clipboard(self):

        clipboard = str(pyperclip.paste()).encode("utf-8")
        date = datetime.now().strftime('%Y-%m-%d, %H:%M:%S')

        self.item_model.insertRow(0)
        self.item_model.setItem(0, 1, QtGui.QStandardItem(date))
        self.item_model.setItem(0, 2, QtGui.QStandardItem(clipboard))

        print('Item inserted to table')

        # Update database
        try:
            self.cur.execute('INSERT INTO CLIPBOARD(DATE, CONTENT) VALUES(?, ?)', (date, clipboard))
            self.conn.commit()
            self.cur.execute('SELECT MAX(ID) as ID FROM CLIPBOARD')
        except lite.IntegrityError:
            print('SQLite error while adding row')
        id_str = str(self.cur.fetchone()[0])
        print("Item id: " + id_str)
        self.item_model.setItem(0, 0, QtGui.QStandardItem(id_str))

    def event_about(self):
        msgBox = QMessageBox()
        msgBox.setWindowTitle(APP_NAME)
        msgBox.setWindowIcon(QIcon(APP_ICON))
        msgBox.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
        msgBox.setText(
                '''
                \n%s %s programmed by Snir Turgeman\nIf you found some bugs, please report to snir.tur@gmail.com
                \nEnjoy!\n
                ''' % (APP_NAME, APP_VERSION)
        )
        msgBox.exec_()

    @QtCore.Slot()
    def clear_table(self):

        while self.item_model.rowCount() > 0:
            self.item_model.removeRow(0)

        # DB Update
        self.cur.execute('DELETE FROM CLIPBOARD')
        self.cur.execute('delete from sqlite_sequence where name="CLIPBOARD"')
        self.conn.commit()
        print('Table CLIPBOARD deleted')

    def db_to_table(self):
        self.cur.execute('''SELECT * FROM CLIPBOARD''')
        self.item_model.setRowCount(self.cur.rowcount)
        for row, form in enumerate(self.cur.fetchall()):

            # Match data
            id = form[0]
            date = form[1]
            content = form[2]

            # Insert data to table
            self.item_model.insertRow(0)
            self.item_model.setItem(0, 0, QtGui.QStandardItem(str(id)))
            self.item_model.setItem(0, 1, QtGui.QStandardItem(str(date)))
            self.item_model.setItem(0, 2, QtGui.QStandardItem(content))

    def check_db(self):
        try:
            self.cur.execute('''SELECT name FROM sqlite_master WHERE type='table' AND name="CLIPBOARD"''')
            table = self.cur.fetchall()
            if table.__len__() == 0:
                self.cur.execute('CREATE TABLE CLIPBOARD (ID INTEGER PRIMARY KEY ASC AUTOINCREMENT, DATE DATETIME, CONTENT VARCHAR)')
                self.cur.execute('CREATE TABLE PROPERTIES (NAME VARCHAR PRIMARY KEY, VALUE VARCHAR)')
                self.cur.execute('INSERT INTO PROPERTIES(NAME, VALUE) VALUES ("ALWAYS_ON_TOP", "Y")')
                self.cur.execute('INSERT INTO PROPERTIES(NAME, VALUE) VALUES ("BUFFER_SIZE", "200")')
                self.conn.commit()
        except lite.IntegrityError:
            print("Error occurd while trying to create CLIPBOARD table")
            msgBox = QMessageBox().about()
            msgBox.setWindowTitle(APP_NAME)
            msgBox.setWindowIcon(QIcon(APP_ICON))
            msgBox.setWindowFlags(QtCore.Qt.WindowStaysOnTopHint)
            msgBox.setText("Error occurd while trying to create CLIPBOARD table, Please report to developer")
            msgBox.exec_()
    def always_on_top_toggle(self, state):

        if state == QtCore.Qt.Checked:
            self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
            self.cur.execute('UPDATE PROPERTIES SET VALUE = "Y" WHERE NAME="ALWAYS_ON_TOP"')
            self.conn.commit()
        else:
            self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
            self.cur.execute('UPDATE PROPERTIES SET VALUE = "N" WHERE NAME="ALWAYS_ON_TOP"')
            self.conn.commit()

    def get_always_on_top_status(self):
        self.cur.execute('SELECT * FROM PROPERTIES WHERE NAME = ?', ["ALWAYS_ON_TOP"])
        alwaysOnTop = self.cur.fetchone()
        if str(alwaysOnTop[1]) == 'Y':
            self.setWindowFlags(self.windowFlags() | QtCore.Qt.WindowStaysOnTopHint)
            return True
        else:
            self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowStaysOnTopHint)
            return False
    def delete_selected_rows(self):
        # Remove data from table and save their id for Data update
        item_selection = QtGui.QItemSelection(self.ui.cb_table.selectionModel().selection())
        print(str(item_selection.size()) + " Selected")
        for idx in reversed(item_selection.indexes()):
            row = idx.row()
            print("Row " + str(row) + " Deleted")
            row_id = self.filter_proxy_model.index(row, 0).data()
            self.filter_proxy_model.removeRow(row)
            try:
                self.cur.execute('DELETE FROM CLIPBOARD WHERE ID = (?)', [row_id])

            except lite.IntegrityError:
                print("Error! Could not delete items from db")
        self.conn.commit()

    def copy_selected_rows(self):
        copy=''
        item_selection = QtGui.QItemSelection(self.ui.cb_table.selectionModel().selection())
        print(str(item_selection.size()) + " Selected")
        for idx in item_selection.indexes():
            row = idx.row()
            print("Row " + str(row) + " copied")
            row_value = self.filter_proxy_model.index(row, 2).data()
            copy = copy + row_value + '\n'
        pyperclip.copy(copy)

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    frame = MainWindow()
    frame.setWindowFlags(frame.windowFlags() & QtCore.Qt.CustomizeWindowHint)
    frame.setWindowFlags(frame.windowFlags() & ~QtCore.Qt.WindowMinimizeButtonHint)
    frame.initialize()
    frame.show()

    app.exec_()
