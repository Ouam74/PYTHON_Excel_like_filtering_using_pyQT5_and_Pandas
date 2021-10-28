from ExcelFiltering_gui import *
from ExcelFiltering_functions import *
from PyQt5 import QtWidgets, QtCore, QtGui
import sys
import openpyxl
import pandas as pd
sys.setrecursionlimit(100000)
print(chr(27) + "[2J")


def sort_func(list_array):
    global df
    col_count = df.shape[1]
    df_filtered = ~df[df.columns[0]].isin(list_array[0])
    for i in range(col_count):
        df_filtered_new = ~df[df.columns[i]].isin(list_array[i])
        df_filtered = df_filtered & df_filtered_new
    return df.loc[df_filtered]



def getTreeChoices(index):
    global df
    choices = df.iloc[:][df.columns[index]]
    return sorted(list(set(choices)))



class StandardItemModel(QtGui.QStandardItemModel):
    myModelSignal = QtCore.pyqtSignal(object, object)

    def setData(self, index, value, role= QtCore.Qt.EditRole):
        oldvalue = index.data(role)
        result = super(StandardItemModel, self).setData(index, value, role)
        if result and value != oldvalue: # if an item is changed then the signal is emitted
            self.myModelSignal.emit(self.itemFromIndex(index), role)
        if index.data() == "Select all":
            if index.data(role) == 2: # --> Unchecked
                for index in range(self.rowCount()):
                    self.item(index).setCheckState(QtCore.Qt.Checked)
            elif index.data(role) == 0: # --> Checked
                for index in range(self.rowCount()):
                    self.item(index).setCheckState(QtCore.Qt.Unchecked)
        return result



class FormWidget(QtWidgets.QWidget):
    myFormSignal = QtCore.pyqtSignal(str)

    def __init__(self, parent = None):
        super(FormWidget, self).__init__(parent)
        self.setWindowFlag(QtCore.Qt.FramelessWindowHint)
        self.gridLayout = QtWidgets.QGridLayout(self)
        self.gridLayout.setObjectName("gridLayout")
        self.treeView = QtWidgets.QTreeView(self)
        self.treeView.setObjectName("treeView")
        self.gridLayout.addWidget(self.treeView, 0, 0, 1, 1)
        self.horizontalLayout = QtWidgets.QHBoxLayout()
        self.horizontalLayout.setObjectName("horizontalLayout")
        self.enterButton = QtWidgets.QPushButton(self)
        self.enterButton.setText("Enter")
        self.enterButton.setObjectName("enterButton")
        self.horizontalLayout.addWidget(self.enterButton)
        self.cancelButton = QtWidgets.QPushButton(self)
        self.cancelButton.setText("Cancel")
        self.cancelButton.setObjectName("cancelButton")
        self.horizontalLayout.addWidget(self.cancelButton)
        self.gridLayout.addLayout(self.horizontalLayout, 1, 0, 1, 1)
        self.enterButton.clicked.connect(self.onEnterClicked)
        self.cancelButton.clicked.connect(self.onCancelClicked)

    def setTreeItemModel(self, model):
        self.model = model
        self.treeView.setModel(self.model)

    def onEnterClicked(self):
        self.myFormSignal.emit(self.enterButton.objectName())
        self.close()

    def onCancelClicked(self):
        self.close()

    def on_close(self):
        w.close()



class PushButton(QtWidgets.QPushButton):
    def __init__(self, parent=None):
        super(PushButton, self).__init__(parent)
        self.clicked.connect(self.setForm)

    def setForm(self):
        self.form.setTreeItemModel(self.model)
        self.form.enterButton.setObjectName("%i" %self.index)
        self.form.setWindowModality(QtCore.Qt.ApplicationModal)
        self.form.show()

    def setPushItemModel(self, model):
        self.model = model
        self.model.setObjectName("%i" %self.index)

    def setPushIndex(self, index):
        self.index = index

    def setPushForm(self, form):
        self.form = form



class HorizontalHeader(QtWidgets.QHeaderView):
    def __init__(self, parent=None):
        super(HorizontalHeader, self).__init__(QtCore.Qt.Horizontal, parent)
        self.model_array = []
        self.form_array = []
        self.col_names = []
        self.push_array = []
        # self.setSectionsMovable(True)
        self.sectionResized.connect(self.handleSectionResized)
        self.sectionMoved.connect(self.handleSectionMoved)

    # def showEvent(self, event):
    #     # self.show()
    #     self.update()

    def showTreeItems(self, form_array, col_names):
        self.form_array = form_array
        self.col_names = col_names
        self.col_count = df.shape[1]
        for i in range(self.col_count):  # self.count() --> number of headers
            self.push = PushButton(self)
            self.push.setPushIndex(i)
            self.push.setPushForm(self.form_array[i])
            self.push.setText(self.col_names[i])
            self.push.setPushItemModel(self.model_array[i])
            self.push_array.append(self.push)
            self.push.setGeometry(self.sectionViewportPosition(i), 0, self.sectionSize(i)-5, self.height())
            self.push.show()

    def addTreeItems(self, model_array):
        global df
        self.model_array = model_array
        self.col_count = df.shape[1]
        for i in range(self.col_count):
            values = getTreeChoices(i)
            self.model = self.model_array[i]
            self.model.setHorizontalHeaderLabels([''])
            item = QtGui.QStandardItem()
            item.setText("Select all")
            item.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
            item.setData(QtCore.Qt.Checked, QtCore.Qt.CheckStateRole)
            self.model.appendRow(item)
            for j in range(len(values)):
                item = QtGui.QStandardItem()
                item.setText(str(values[j]))
                item.setFlags(QtCore.Qt.ItemIsUserCheckable | QtCore.Qt.ItemIsEnabled)
                item.setData(QtCore.Qt.Checked, QtCore.Qt.CheckStateRole)
                self.model.appendRow(item)

    def handleSectionResized(self, i):
        for i in range(self.count()):
            j = self.visualIndex(i)
            logical = self.logicalIndex(j)
            self.push_array[i].setGeometry(self.sectionViewportPosition(logical), 0, self.sectionSize(logical)-4, self.height())

    def handleSectionMoved(self, i, oldVisualIndex, newVisualIndex):
        for i in range(min(oldVisualIndex, newVisualIndex), self.count()):
            logical = self.logicalIndex(i)
            self.push_array[i].setGeometry(self.sectionViewportPosition(logical), 0, self.sectionSize(logical)-5, self.height())

    def fixButtonPositions(self):
        for i in range(self.count()):
            self.push_array[i].setGeometry(self.sectionViewportPosition(i), 0, self.sectionSize(i)-5, self.height())



class TableWidget(QtWidgets.QTableWidget):
    def __init__(self, header, parent = None):
        super(TableWidget, self).__init__(parent)
        self.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.setDragDropMode(QtWidgets.QAbstractItemView.NoDragDrop)
        self.setAlternatingRowColors(True)
        self.header = header
        self.setHorizontalHeader(self.header)
        self.setRowCount(50)
        self.setColumnCount(50)
        # self.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        # self.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        # self.setSizeAdjustPolicy(QtWidgets.QAbstractScrollArea.AdjustToContents)

    def scrollContentsBy(self, dx, dy):
        super(TableWidget, self).scrollContentsBy(dx, dy)
        if dx != 0:
            self.horizontalHeader().fixButtonPositions()



class DesignerMainWindow(QtWidgets.QMainWindow, Ui_MainWindow):
# class DesignerMainWindow(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        # Initialization of the superclass
        super(DesignerMainWindow, self).__init__(parent)
        # uic.loadUi('DCProbing.ui', self)
        self.setupUi(self)

        # Define Variable
        self.col_count = 0
        self.row_count = 0
        self.model_array = []
        self.form_array = []
        self.col_names = []

        # Creating Table widget
        self.layout = QtWidgets.QGridLayout(self.frame_1)
        self.Header = HorizontalHeader(self)
        self.Table = TableWidget(self.Header)
        self.layout.addWidget(self.Table)

        # Signals
        self.RunPushButton.clicked.connect(self.startThreads)
        self.ActionLoad.triggered.connect(self.readTable)
        self.AbortPushButton.clicked.connect(self.abortClicked)

    def isQuitClicked(self):
        sys.exit(app.exec_())

    def startThreads(self):
        # pdb.set_trace()
        # Create Threads
        self.clearGUIlog()
        self.AbortPushButton.setEnabled(True)
        self.Table.setEnabled(False)
        # FILENAME = self.FilenameLineEdit.text()
        # DMM_ADDR = self.DmmAddrLineEdit.text()
        # MEAS_ARR = []

        self.measureWorker = measureWorker()
        self.measureThread = QtCore.QThread()
        self.measureThread.started.connect(self.measureWorker.run)  # Init worker run() at startup (optional)
        self.measureWorker.update_log.connect(self.updateGUIlog)
        self.measureWorker.update_progressbar.connect(self.updateGUIprogress)
        self.measureWorker.clear_log.connect(self.clearGUIlog)
        self.measureWorker.end_thread.connect(self.endThreads)
        self.measureWorker.moveToThread(self.measureThread)  # Move the Worker object to the Thread object
        self.measureThread.start()

        self.blinkWorker = blinkWorker()
        self.blinkThread = QtCore.QThread()
        self.blinkThread.started.connect(self.blinkWorker.run)  # Init worker run() at startup
        self.blinkWorker.update_blinkcolor.connect(self.updateGUIblinkColor)
        self.blinkWorker.moveToThread(self.blinkThread)  # Move the Worker object to the Thread object
        self.blinkThread.start()

    def abortClicked(self):
        try:
            self.LogPlainTextEdit.appendHtml("<font color = \"red\"> MEASUREMENT ABORTED!")
            self.measureWorker.isAborted = True
            self.blinkWorker.isaborted = True
        except:
            pass

    def endThreads(self):
        # Terminate Qthreads
        try:
            self.LogPlainTextEdit.appendHtml("<font color = \"green\"> END THREAD")
            self.blinkWorker.isrunning = False
            self.measureThread.quit()
            self.measureThread.wait()
            self.blinkThread.quit()
            self.blinkThread.wait()
            # Print pop-up message when measurement finished or aborted
            if self.measureWorker.isAborted == True:  # Measurement aborted
                QtWidgets.QMessageBox.about(self, "STATUS", "Measurement aborted")
            else:
                QtWidgets.QMessageBox.about(self, "STATUS", "Measurement finished")
            # Reset GUI
            self.LogPlainTextEdit.appendHtml("<font color = \"green\"> COMPLETED")
            self.RunPushButton.setStyleSheet("background-color : Green")
            self.AbortPushButton.setEnabled(False)
            self.Table.setEnabled(True)
        except:
            pass

    def closeEvent(self, event):
        # Use if the main window is closed by the user
        close = QtWidgets.QMessageBox.question(self, "QUIT", "Confirm quit?", QtWidgets.QMessageBox.Yes | QtWidgets.QMessageBox.No)
        if close == QtWidgets.QMessageBox.Yes:
            event.accept()
            app.quit()
            # sys.exit(0)
        else:
            event.ignore()

    def updateGUIlog(self, flag, txt):
        if flag == 0:  # Text in black
            self.LogPlainTextEdit.appendHtml("<font color = \"black\"> %s" % txt)
        if flag == 1:  # Text in red
            self.LogPlainTextEdit.appendHtml("<font color = \"red\"> %s" % txt)
        if flag == 2:  # Text in green
            self.LogPlainTextEdit.appendHtml("<font color = \"green\"> %s" % txt)

    def clearGUIlog(self):
        self.ProgressBar.setValue(0)
        self.LogPlainTextEdit.clear()

    def updateGUIprogress(self, i):
        self.ProgressBar.setValue(i)

    def updateGUIblinkColor(self, color):
        if color == 0:
            self.RunPushButton.setStyleSheet(color)
        else:
            self.RunPushButton.setStyleSheet(color)

    def readTable(self):
        global df
        filename, _ = QtWidgets.QFileDialog.getOpenFileName(self, "Load File", "", "Config file (*.xlsx)")
        try:
            df = pd.read_excel(filename, sheet_name=0, header=0, index_col=False, keep_default_na=True, skiprows=0)
            df = df.applymap(str) # Convert all elements to string
            self.col_names = df.columns # Get column names
            self.row_count = df.shape[0]
            self.col_count = df.shape[1]
            for i in range(self.col_count):
                self.model_array.append(StandardItemModel(self))
                self.form_array.append(FormWidget())

            # Create and fill headers with Pushbuttons (fill in models with items and create forms)
            self.Header.addTreeItems(self.model_array)
            self.Header.showTreeItems(self.form_array, self.col_names)
            self.fillTable()
            self.RunPushButton.setEnabled(True)

            # Signals / Slots
            for i in range(self.col_count):
                self.form_array[i].myFormSignal.connect(self.sortData)
        except:
            pass

    def fillTable(self):
        self.Table.setRowCount(self.row_count)
        self.Table.setColumnCount(self.col_count)
        for i in range(self.row_count):
            for j in range(self.col_count):
                item = QtWidgets.QTableWidgetItem(str(df.iloc[i][df.columns[j]]))
                self.Table.setItem(i, j, item)

    # def sortData(self, item, role):
    #     if role == QtCore.Qt.CheckStateRole:
    #         sender = self.sender()
    #         print(sender.objectName(), item.text(), item.checkState())

    def sortData(self):
        global df_filtered
        col_count = df.shape[1]
        list_array = []

        for i in range(col_count):
            liste = []
            row_count = self.model_array[i].rowCount()
            for j in range(row_count):
                ischecked = self.model_array[i].item(j).checkState()
                txt = self.model_array[i].item(j).text()
                if (ischecked == False) and (txt != "Select all"):
                    liste.append(txt)
            list_array.append(liste)

        # Filter the dataframe
        df_filtered = sort_func(list_array)

        # Fill Table
        self.Table.clear()
        self.row_count = df_filtered.shape[0]
        self.col_count = df_filtered.shape[1]
        for i in range(self.row_count):
            for j in range(self.col_count):
                item = QtWidgets.QTableWidgetItem(str(df_filtered.iloc[i][df_filtered.columns[j]]))
                self.Table.setItem(i, j, item)


# Measurement Worker
class measureWorker(QtCore.QObject):
    update_log = QtCore.pyqtSignal(int, str)
    clear_log = QtCore.pyqtSignal()
    update_progressbar = QtCore.pyqtSignal(int)
    end_thread = QtCore.pyqtSignal()
    def __init__(self, parent = None):
        QtCore.QThread.__init__(self, parent = None)
        self.isAborted = False
    @QtCore.pyqtSlot()
    def run(self):
        measure1(self)

# Blinking Worker
class blinkWorker(QtCore.QObject):
    update_blinkcolor = QtCore.pyqtSignal(str)
    def __init__(self, parent = None):
        QtCore.QThread.__init__(self, parent = None)
        self.isaborted = False
        self.isrunning = True
    @QtCore.pyqtSlot()
    def run(self):
        blink(self)

########################################################################################################################

if __name__ == '__main__':
    df = None
    df_filtered = None
    if not QtWidgets.QApplication.instance():
        app = QtWidgets.QApplication(sys.argv)
    else:
        app = QtWidgets.QApplication.instance()
    app.setStyle('Fusion')  # 'Breeze', 'Oxygen', 'QtCurve', 'Windows', 'Fusion'
    w = DesignerMainWindow()
    w.show()
    sys.exit(app.exec_())