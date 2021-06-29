
"""
    Project Searcher using python and QtGui wrapper.
    Author: Sarath Kumar
    Version: 2021, June
"""

import sys
from PyQt5 import QtCore, QtWidgets, QtGui, uic
import os
import re
from pathlib import Path
import pandas as pd
from time import sleep, strftime, gmtime
from UI import Ui_MainWindow


class Project_Search_App(QtWidgets.QMainWindow, Ui_MainWindow):

    def __init__(self, *args, **kwargs):

        super().__init__(*args, **kwargs)

        self.setupUi(self)
        self.threadSearch = None
        self.actionExit.triggered.connect(self.qexit)
        self.actionSave.triggered.connect(self.saveOutput)
        self.actionInfo.triggered.connect(self.aboutApp)
        self.actionCopy.triggered.connect(self.copyOutput)
        self.actionClear.triggered.connect(self.clearOutput)

        self.Button_Quit.clicked.connect(self.qexit)
        self.Button_StartSearch.clicked.connect(self.startSearch)
        self.Button_FileBrowser1.clicked.connect(self.searchPathBrowser)
        self.Button_CancelSearch.clicked.connect(self.stopSearch)


        self.threadOutputTerminal = OutputTerminalTreadClass()
        self.threadOutputTerminal.start()
        self.Clipboard = QtWidgets.QApplication.clipboard()

        self.statusBar().showMessage('Ready to Start')

    def exceptionLog(self, title, text, logOutput = True):
        if logOutput:
            self.threadOutputTerminal.append(title + ":" + text)
        QtWidgets.QMessageBox.critical(None, title , text)

    def informationLog(self, title, text, logOutput = True):
        if logOutput:
            self.threadOutputTerminal.append(title + ":" + text)
        QtWidgets.QMessageBox.information(None, title , text)

    def aboutApp(self):
        message = """
        <h3>Project</b>: Project Searcher</h3>
        <br>
        <b>Build</b>: 2021.June<br>
        <b>Author</b>: Sarath Kuamr K<br>
        <b>Github</b>: <a href="https://github.com/Sarath060/Project-Searcher.git">https://github.com/Sarath060/Project-Searcher.git</a>
        """
        about_dialog = QtWidgets.QMessageBox()
        about_dialog.about(self, "About", message)

    def copyOutput(self):
        try:
            res = str(self.Text_OutputTerminal.toPlainText())
            self.Clipboard.setText(res)
        except Exception as e:
            self.exceptionLog("Error", f"Oops!{e.__class__} Error occurred at Copying.")

    def clearOutput(self):
        try:
            self.Text_OutputTerminal.setPlainText("")
        except Exception as e:
            self.exceptionLog("Error", f"Oops!{e.__class__} Error occurred at Clearing Output.")

    @staticmethod
    def qexit():
        sys.exit(0)

    def saveOutput(self):
        try:
            saveFile = QtWidgets.QFileDialog.getSaveFileName(None, "Save File","/home/jana/Search Result.txt", "Textfile (*.txt)")
            filename = saveFile[0]
            file1 = open(filename, "w")
            res = str(self.Text_OutputTerminal.toPlainText())
            file1.writelines(res)
            file1.close()
            self.informationLog( "Information", f"File is Saved at {filename}")
        except Exception as e:
            self.exceptionLog("Error", f"Oops!{e.__class__} Error occurred at Saving Output.")

    def checkBoxSetting(self):
        extensionList = []
        # formate(Extension, return folder or files, excel sheet name)
        if self.checkBox_Step7.isChecked():
            pass
        extensionList.append(["S7S", 0, "Step7 Classic"])
        if self.checkBox_WinCC.isChecked(): extensionList.append(["mcp", 0, "WinCC"])
        if self.checkBox_TIAv11.isChecked(): extensionList.extend([['ap11', 0, "TIA 11"], ['zap11', -1, "TIA 11"]])
        if self.checkBox_TIAv12.isChecked(): extensionList.extend([['ap12', 0, "TIA 12"], ['zap12', -1, "TIA 12"]])
        if self.checkBox_TIAv13.isChecked(): extensionList.extend([['ap13', 0, "TIA 13"], ['zap13', -1, "TIA 13"]])
        if self.checkBox_TIAv14.isChecked(): extensionList.extend([['ap14', 0, "TIA 14"], ['zap14', -1, "TIA 14"]])
        if self.checkBox_TIAv15.isChecked(): extensionList.extend([['ap15', 0, "TIA 15"], ['zap15', -1, "TIA 15"]])
        if self.checkBox_TIAv16.isChecked(): extensionList.extend([['ap16', 0, "TIA 16"], ['zap16', -1, "TIA 16"]])
        if self.checkBox_TIAv17.isChecked(): extensionList.extend([['ap17', 0, "TIA 17"], ['zap17', -1, "TIA 17"]])

        otherExtensions = self.textList(self.Input_ExtensionList.text())

        for extension in otherExtensions:
            extensionList.append([extension, "0", extension])

        return extensionList

    def startSearch(self):

        if not self.checkParameters():
            self.threadOutputTerminal.append(f"Search Started not Started")
            return

        self.threadOutputTerminal.append(f"Search Started in path Directory")
        self.Button_StartSearch.setEnabled(False)
        path = self.Input_SearchPath.text()
        outputFile = self.Input_OutputPath.text()

        extensionList = self.checkBoxSetting()
        containsTxt = self.textList(self.Input_IncludeText.text())
        excludeTxt = self.textList(self.Input_ExcludeText.text())
        subFolders = self.checkBox_SearchSubFolder.isChecked()
        includeArchive = self.checkBox_SearchArchive.isChecked()

        self.statusBar().showMessage('Running Please Wait...')

        self.threadSearch = SearchThreadClass(path, extensionList, containsTxt, excludeTxt, outputFile, subFolders,
                                              includeArchive)
        self.threadSearch.start()

    @staticmethod
    def textList(txt):
        if txt == "":
            return []
        else:
            return list(txt.split(","))

    def stopSearch(self):
        self.threadSearch.stop()
        self.threadOutputTerminal.append(f"Search forced to stop")

    def checkParameters(self):
        if not self.checkPath(self.Input_SearchPath.text()):
            self.informationLog("Invalid Input","Select a vaild path for Searching")
            return False
        if not self.checkPath(os.path.dirname(self.Input_OutputPath.text())):
            self.informationLog("Invalid Input","Select a vaild path for Saving Excel file")
            return False
        return True

    @staticmethod
    def fileBrowser(defaultPath='D:\\', extension="*.*"):
        try:
            dlg = QtWidgets.QFileDialog()
            dlg.setFileMode(QtWidgets.QFileDialog.Directory)
            dlg.setDirectory(defaultPath)
            dlg.show()
            windowsfilenames = []
            if dlg.exec_():
                filenames = dlg.selectedFiles()
                for files in filenames:
                    windowsfilenames.append(files.replace("/", "\\"))
            else:
                return [defaultPath]
            return windowsfilenames
        except Exception as e:
            mainWindow.threadOutputTerminal.append(f"Error Oops!{e.__class__} Error occurred file browser.")
            return [defaultPath]

    @staticmethod
    def checkPath(path):
        if os.path.isdir(path):
            return True
        return False

    def searchPathBrowser(self):
        path = self.fileBrowser()[0]
        self.Input_SearchPath.setText(path)
        self.Input_OutputPath.setText(path + "\Project Search Result.xlsx")
        self.threadOutputTerminal.append(f"Output Path Set to {self.Input_OutputPath.text()}")


class SearchThreadClass(QtCore.QThread):
    any_signal = QtCore.pyqtSignal(int)

    def __init__(self, path=None, extensionList=None, containsTxt=None, excludeTxt=None,
                 outputFile=None, subFolders=None, includeArchive=None, parent=None):
        super(SearchThreadClass, self).__init__(parent)
        self.path = path
        self.extensionList = extensionList
        self.containsTxt = containsTxt
        self.excludeTxt = excludeTxt
        self.outputFile = outputFile
        self.subFolders = subFolders
        self.includeArchive = includeArchive
        self.extensions, self.folderIndents, self.sheetNames = zip(*extensionList)
        self.regExp = self.listToRegEXP()
        self.is_running = True


    def run(self):
        mainWindow.threadOutputTerminal.append(f'Search Started {strftime("%Y-%m-%d %H:%M:%S", gmtime())}...')
        try:
            try:
                df = self.getDF()
            except Exception as e:
                mainWindow.threadOutputTerminal.append(f"Error Oops!{e.__class__} Error occurred in Search.")

            if df.shape[0] > 0:
                df["FLS_File_PathName"] = '=HYPERLINK("' + df["FLS_File_PathName"] + '")'
                tempList = list(self.sheetNames)
                tempList.sort()
                excelSheets = set(tempList)
                try:
                    with pd.ExcelWriter(self.outputFile) as writer:
                        mainWindow.threadOutputTerminal.append(f"Started to write in excel sheet")
                        for sheet in excelSheets:
                            sleep(0.1)
                            dfFilter = [extension for extension, sheetName in zip(self.extensions, self.sheetNames) if sheet == sheetName]
                            try:
                                dfWriter = df.query('Extension in @dfFilter')
                                if dfWriter.shape[0] > 0:
                                    dfWriter.to_excel(writer, sheet_name=sheet)
                                    mainWindow.threadOutputTerminal.append(f"Sheet Name: {sheet}: Project found {dfWriter.shape[0]}")
                                else:
                                    dfWriter = pd.DataFrame({
                                        'FLS_File_Size': ["No Project Found"],
                                        'FLS_File_Access_Date': ["No Project Found"],
                                        'FLS_File_Modification_Date': ["No Project Found"],
                                        'FLS_File_Creation_Date': ["No Project Found"],
                                        'FLS_File_PathName': ["No Project Found"],
                                        'Extension': ["No Project Found"],
                                    })
                                    dfWriter.to_excel(writer, sheet_name=sheet)
                                    mainWindow.threadOutputTerminal.append( f"Sheet Name: {sheet}: No project found")
                            except Exception as e:
                                mainWindow.threadOutputTerminal.append(f"Oops!{e.__class__} Error occurred Write "
                                                                       f"Filtering and Writing to Excel sheet{sheet}.")
                except Exception as e:
                    mainWindow.threadOutputTerminal.append(f"Oops!{e.__class__} excel writer. ")

            elif df.shape[0] <= 0:
                mainWindow.threadOutputTerminal.append(f"No project found")

        except Exception as e:
            mainWindow.threadOutputTerminal.append(f"Oops!{e.__class__} Search run. ")
        mainWindow.threadOutputTerminal.append(f'Finished Search {strftime("%Y-%m-%d %H:%M:%S", gmtime())}...')
        mainWindow.Button_StartSearch.setEnabled(True)
        mainWindow.statusBar().showMessage('Ready to Start')


    def stop(self):
        self.is_running = False
        mainWindow.threadOutputTerminal.append('Stopping thread... File Search Thread Class')
        mainWindow.Button_StartSearch.setEnabled(True)
        self.terminate()

    def findFilesInFolderYield(self, path):

        myregexobj1 = re.compile(self.regExp)  # Makes sure the file extension is at the end and is preceded by a .

        try:  # Trapping a OSError or FileNotFoundError:  File permissions problem I believe
            for entry in os.scandir(path):

                if entry.is_file() and myregexobj1.search(entry.path):  #

                    bFlagContains = [(True for txt in self.containsTxt if txt in entry.path)]
                    bFlagExclude = [(True for txt in self.excludeTxt if txt not in entry.path)]

                    if (bFlagContains.count(True) == len(self.containsTxt)) and (bFlagExclude.count(True) == len(self.excludeTxt)):

                        sleep(0.01)
                        mainWindow.threadOutputTerminal.append(f'...{str(entry.path)[-60:]}')

                        extension = str(os.path.splitext(entry.path)[1]).replace(".", "")
                        folderIndent = self.folderIndents[self.extensions.index(extension)]
                        if folderIndent == -1:
                            yield entry.stat().st_size, entry.stat().st_atime_ns, entry.stat().st_mtime_ns, entry.stat().st_ctime_ns, entry.path, extension
                        else:
                            projectPath = str(Path(entry.path).parents[folderIndent])
                            projectEntry = os.stat(projectPath)
                            projectSize = sum(file.stat().st_size for file in Path(projectPath).rglob('*'))
                            yield projectSize, projectEntry.st_atime_ns, projectEntry.st_mtime_ns, projectEntry.st_ctime_ns, projectPath, extension

                elif entry.is_dir() and self.subFolders:  # if its a directory, then repeat process as a nested function
                    yield from self.findFilesInFolderYield(entry.path)

        except FileNotFoundError as fnf:
            mainWindow.threadOutputTerminal.append(path + ' not found ', fnf)
        except OSError as ose:
            mainWindow.threadOutputTerminal.append('Cannot access ' + path + '. Probably a permissions error ', ose)
        except Exception as e:
            mainWindow.threadOutputTerminal.append(f"Error Oops!{e.__class__} Error occurred Searching.")

    def listToRegEXP(self):
        regExp = ""
        for extension in self.extensions:
            regExp += '\.' + extension + '$|'
        return regExp[0:-1]

    def getDF(self):
        try:
            fileSizes, accessTimes, modificationTimes, creationTimes, paths, Extensions = zip(*(self.findFilesInFolderYield(self.path)))

            df = pd.DataFrame({
                'FLS_File_Size': fileSizes,
                'FLS_File_Access_Date': accessTimes,
                'FLS_File_Modification_Date': modificationTimes,
                'FLS_File_Creation_Date': creationTimes,
                'FLS_File_PathName': paths,
                'Extension': Extensions,
            })

            df['FLS_File_Modification_Date'] = pd.to_datetime(df['FLS_File_Modification_Date'], infer_datetime_format=True)
            df['FLS_File_Creation_Date'] = pd.to_datetime(df['FLS_File_Creation_Date'], infer_datetime_format=True)
            df['FLS_File_Access_Date'] = pd.to_datetime(df['FLS_File_Access_Date'], infer_datetime_format=True)
        except Exception as e:
            mainWindow.threadOutputTerminal.append(f"Error Oops!{e.__class__} Error occurred Database.")
            df = pd.DataFrame({
                'FLS_File_Size': [],
                'FLS_File_Access_Date': [],
                'FLS_File_Modification_Date': [],
                'FLS_File_Creation_Date': [],
                'FLS_File_PathName': [],
                'Extension': [],
            })

        return df


class OutputTerminalTreadClass(QtCore.QThread):
    any_signal = QtCore.pyqtSignal(int)

    def __init__(self, parent=None):
        super(OutputTerminalTreadClass, self).__init__(parent)
        self.is_running = True
        self.entryCount = 0
        self.maxEntryCount = 600

    def run(self):
        pass

    def append(self, text):
        try:
            mainWindow.Text_OutputTerminal.appendPlainText(text)
            self.entryCount += 1
            mainWindow.Text_OutputTerminal.moveCursor(QtGui.QTextCursor.End)
        except Exception as e:
            print("Error", f"Oops!{e.__class__} Error occurred at output V1 Terminal {self.entryCount}.")
        if self.entryCount >= self.maxEntryCount:
            self.removeEntry()

    def removeEntry(self):
        self.entryCount = 0
        mainWindow.Text_OutputTerminal.setPlainText("")

    def stop(self):

        mainWindow.threadOutputTerminal.append('Stopping thread... Output Terminal Tread')
        self.is_running = False
        self.terminate()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    mainWindow = Project_Search_App()
    mainWindow.show()
    sys.exit(app.exec_())
