#-- coding: utf-8 -*-
'''
Created on 2015年7月10日

@author: Wan
'''
import win32com.client
from win32com.client import constants

class Word(object):
    def __init__(self, filename=None):
        self.Word = win32com.client.Dispatch('Word.Application')
        # 后台运行，不显示，不警告
        self.Word.Visible = False
        self.Word.DisplayAlerts = False
        
        if filename:
            self.doc = self.Word.Documents.Open(filename)
        else:
            # 创建新的文档
            self.doc = self.Word.Documents.Add()
            
    def replace(self, datadict):
        for oldstr, newstr in datadict.items():
            oldstr = "{%s}" % oldstr
            self.Word.Selection.Find.ClearFormatting()
            self.Word.Selection.Find.Replacement.ClearFormatting()
            self.Word.Selection.Find.Execute(oldstr, False, False, False, False, 
                                               False, True, 1, True, newstr, 2)
            
    def SaveAs(self, filename):
        self.doc.SaveAs(filename)
        
 
    def ExportAsPDF(self, filename):
        self.doc.ExportAsFixedFormat(filename, constants.wdExportFormatPDF,
                                     Item = constants.wdExportDocumentWithMarkup, 
                                     CreateBookmarks = constants.wdExportCreateHeadingBookmarks)

    def close(self, SaveChanges=0):
        self.doc.Close(SaveChanges)
        self.Word.Quit()
        
        
class Excel:
    '''
    Some convenience methods for Excel documents accessed
    through COM.
    '''
    def __init__(self, filename=None):
        '''
        Create a new application
        if filename is None, create a new file
        else, open an exsiting one
        '''
        self.xlApp = win32com.client.Dispatch('Excel.Application')
        self.xlApp.Visible = True
        if filename:
            self.filename = filename
            self.xlBook = self.xlApp.Workbooks.Open(filename)
        else:  
            self.xlBook = self.xlApp.Workbooks.Add()
            self.filename = ''
            
    def visible(self,visible = True):
        '''
        if Visible is true, the applicaion is visible
        '''
        self.Visible = visible
        
    def save(self, newfilename=None):
        '''
        if filename is None, save the openning file
        else save as another file used the given name
        '''
        if newfilename:
            self.filename = newfilename  
            self.xlBook.SaveAs(newfilename)  
        else:  
            self.xlBook.Save()
             
    def close(self, SaveChanges=1):
        '''
        Close the application
        '''
        self.xlBook.Close(SaveChanges)
        
    def quit(self):
        self.xlApp.Quit()
        
    def getCell(self, sheet, row, col):  
        '''
        Get value of one cell
        '''
        sht = self.xlBook.Worksheets(sheet)
        return sht.Cells(row, col).Value  
    
    def setCell(self, sheet, row, col, value):  
        '''
        Set value of one cell
        '''
        sht = self.xlBook.Worksheets(sheet)  
        sht.Cells(row, col).Value = value 
        
    def getRange(self,sheet,row1,col1,row2,col2):
        '''
        Return a 2d array (i.e. tuple of tuples)
        '''
        sht = self.xlBook.Worksheets(sheet)
        return sht.Range(sht.Cells(row1,col1),sht.Cells(row2,col2)).Value
    
    def setRange(self,sheet,leftCol,topRow,data):
        '''
        Insert a 2d array starting at given location.
        i.e. [['a','b','c'],['a','b','c'],['a','b','c']]
        Works out the size needed for itself
        '''
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        sht = self.xlBook.Worksheets(sheet)
        sht.Range(sht.Cells(topRow, leftCol), sht.Cells(bottomRow, rightCol)).Value = data
        
    def get_numrows(self, sheet):
        sht = self.xlBook.Worksheets(sheet)
        return sht.UsedRange.Rows.Count