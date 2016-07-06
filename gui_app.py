#!/usr/bin/env python
#-- coding: utf-8 -*-
'''
Created on 2016年6月13日

@author: Wan
'''

from PyQt5.QtCore import (QFile, QFileInfo, QPoint, QSettings, QSize,
        Qt, QTextStream)
from PyQt5.QtGui import QIcon, QKeySequence
from PyQt5.QtWidgets import (QAction, QApplication, QFileDialog, QMainWindow,
        QMessageBox, QTextBrowser)
from common import module_path


class MainWindow(QMainWindow):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setWindowTitle("Rongda ERP version 0.1")
        self.setWindowIcon(QIcon("resources/rd.ico"))
                           
        self.logfile = 'app.log'
        self.loadLog(self.logfile)

        self.textBrowser = QTextBrowser()
        self.setCentralWidget(self.textBrowser)

        self.createActions()
        self.createMenus()
        self.createToolBars()
        self.createStatusBar()

        self.readSettings()


    def closeEvent(self, event):
        if self.maybeSave():
            self.writeSettings()
            event.accept()
        else:
            event.ignore()

    def open(self):
        if self.maybeSave():
            fileName, _ = QFileDialog.getOpenFileName(self)
            if fileName:
                self.loadFile(fileName)

    def about(self):
        QMessageBox.about(self, "About Ronda ERP",
                "The <b>Application</b> is very <b>NB</b>! <br />"
                "it using Qt, with a menu bar, "
                "toolbars, and a status bar.<br />"
                "Author: Wan <br />"
                "Developer: wnh3yang@gmail.com")

    def documentWasModified(self):
        self.setWindowModified(self.textBrowser.document().isModified())

    def createActions(self):
        root = module_path()

        self.newAct = QAction(QIcon(root + '/resources/new.png'), "&New", self,
                shortcut=QKeySequence.New, statusTip="Create a new formula",
                triggered=self.newFormula)

        self.openAct = QAction(QIcon(root + '/resources/open.png'), "&Open...",
                self, shortcut=QKeySequence.Open,
                statusTip="Open an existing formula", triggered=self.open)

        self.saveAct = QAction(QIcon(root + '/resources/save.png'), "&Save", self,
                shortcut=QKeySequence.Save,
                statusTip="Save the document to database", triggered=self.saveFormula)
        
        self.filterAct = QAction(QIcon(root + '/resources/filter.png'), "&Filter", 
                self, shortcut="Ctrl+F",
                statusTip="Filter formulas",
                triggered=self.filterFormula)
        
        self.listFormulasAct = QAction(QIcon(root + '/resources/list.png'), "Formula List", 
                self, statusTip="List all formulas",
                triggered=self.listFormulas)
        
        self.listMaterialsAct = QAction("Show All", self,
                statusTip="List all materials",
                triggered=self.listMaterials)
        
        self.searchMaterialsAct = QAction("Material Search", self,
                statusTip="Search all materials",
                triggered=self.searchMaterials)
        
        self.addMaterialsAct = QAction("Material Add", self,
                statusTip="Add a material",
                triggered=self.addMaterials)

        self.exitAct = QAction("E&xit", self, shortcut="Ctrl+Q",
                statusTip="Exit the application", triggered=self.close)

        self.createAct = QAction(QIcon(root + '/resources/connect.png'), "Crea&te", self,
                shortcut="Ctrl+T",
                statusTip="Create a work order",
                triggered=self.createWork)

        self.executeAct = QAction(QIcon(root + '/resources/cmd.png'), "&Execute", self,
                shortcut="Ctrl+E",
                statusTip="Bulk generate work orders",
                triggered=self.executeWork)

        self.searchAct = QAction(QIcon(root + '/resources/document_search.png'), "Sea&rch",
                self, shortcut="Ctrl+R",
                statusTip="Search work order",
                triggered=self.searchWork)

        self.aboutAct = QAction("&About", self,
                statusTip="Show the application's About box",
                triggered=self.about)

        self.aboutQtAct = QAction("About &Qt", self,
                statusTip="Show the Qt library's About box",
                triggered=QApplication.instance().aboutQt)

    def createMenus(self):
        self.formulaMenu = self.menuBar().addMenu("&Formula")
        self.formulaMenu.addAction(self.newAct)
        self.formulaMenu.addAction(self.openAct)
        self.formulaMenu.addAction(self.saveAct)
        self.formulaMenu.addSeparator();
        self.formulaMenu.addAction(self.filterAct)
        self.formulaMenu.addAction(self.listFormulasAct)
        
        self.materialMenu = self.formulaMenu.addMenu("&Material")
        self.materialMenu.addAction(self.listMaterialsAct)
        self.materialMenu.addAction(self.searchMaterialsAct)
        self.materialMenu.addAction(self.addMaterialsAct)
        
        self.formulaMenu.addAction(self.exitAct)

        self.documentMenu = self.menuBar().addMenu("&Work")
        self.documentMenu.addAction(self.createAct)
        self.documentMenu.addAction(self.executeAct)
        self.documentMenu.addAction(self.searchAct)

        self.menuBar().addSeparator()

        self.helpMenu = self.menuBar().addMenu("&Help")
        self.helpMenu.addAction(self.aboutAct)
        self.helpMenu.addAction(self.aboutQtAct)

    def createToolBars(self):
        self.fileToolBar = self.addToolBar("File")
        self.fileToolBar.addAction(self.newAct)
        self.fileToolBar.addAction(self.openAct)
        self.fileToolBar.addAction(self.saveAct)
        self.fileToolBar.addAction(self.filterAct)

        self.editToolBar = self.addToolBar("Work")
        self.editToolBar.addAction(self.createAct)
        self.editToolBar.addAction(self.executeAct)
        self.editToolBar.addAction(self.searchAct)

    def createStatusBar(self):
        self.statusBar().showMessage("Ready")

    def readSettings(self):
        settings = QSettings("RdSoft", "Rongda ERP")
        pos = settings.value("pos", QPoint(200, 200))
        size = settings.value("size", QSize(400, 400))
        self.resize(size)
        self.move(pos)

    def writeSettings(self):
        settings = QSettings("RdSoft", "Rongda ERP")
        settings.setValue("pos", self.pos())
        settings.setValue("size", self.size())

    def maybeSave(self):
        if self.textBrowser.document().isModified():
            ret = QMessageBox.warning(self, "Application",
                    "The document has been modified.\nDo you want to save "
                    "your changes?",
                    QMessageBox.Save | QMessageBox.Discard | QMessageBox.Cancel)

            if ret == QMessageBox.Save:
                return self.save()

            if ret == QMessageBox.Cancel:
                return False

        return True

    def loadFile(self, fileName):
        file = QFile(fileName)
        if not file.open(QFile.ReadOnly | QFile.Text):
            QMessageBox.warning(self, "Application",
                    "Cannot read file %s:\n%s." % (fileName, file.errorString()))
            return

        inf = QTextStream(file)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        self.textBrowser.setPlainText(inf.readAll())
        QApplication.restoreOverrideCursor()

        self.setCurrentFile(fileName)
        self.statusBar().showMessage("File loaded", 2000)

    def saveFile(self, fileName):
        file = QFile(fileName)
        if not file.open(QFile.WriteOnly | QFile.Text):
            QMessageBox.warning(self, "Application",
                    "Cannot write file %s:\n%s." % (fileName, file.errorString()))
            return False

        outf = QTextStream(file)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        outf << self.textBrowser.toPlainText()
        QApplication.restoreOverrideCursor()

        self.setCurrentFile(fileName);
        self.statusBar().showMessage("File saved", 2000)
        return True

    def strippedName(self, fullFileName):
        return QFileInfo(fullFileName).fileName()
    
    def loadLog(self, fileName):
        pass
    
    def filterFormula(self):
        pass
    
    def listFormulas(self):
        pass
    
    def listMaterials(self):
        pass
    
    def createWork(self):
        pass
    
    def executeWork(self):
        pass
    
    def searchWork(self):
        pass
    
    def newFormula(self):
        pass
    
    def saveFormula(self):
        pass
    
    def searchMaterials(self):
        pass
    
    def addMaterials(self):
        pass


if __name__ == '__main__':

    import sys

    app = QApplication(sys.argv)
    mainWin = MainWindow()
    mainWin.show()
    sys.exit(app.exec_())