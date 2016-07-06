#-- coding: utf-8 -*-
'''
Created on 2015年7月10日

@author: Wan
'''
import win32com.client
from win32com.client import constants

class Word(object):
    def __init__(self, filename=None):
        self.engine = win32com.client.Dispatch('Word.Application')
        self.engine.Visible = False
        self.engine.DisplayAlerts = False
        if filename:
            self.filename = filename
            self.doc = self.engine.Documents.Open(self.filename)
        else:
            self.filename = 'newfile.docx'
            self.doc = self.engine.Documents.Add()
            
    def replace(self, datadict):
        for oldstr, newstr in datadict.items():
            oldstr = "{%s}" % oldstr
            self.engine.Selection.Find.ClearFormatting()
            self.engine.Selection.Find.Replacement.ClearFormatting()
            self.engine.Selection.Find.Execute(oldstr, False, False, False, False, 
                                               False, True, 1, True, newstr, 2)
            
    def save(self, filename):
        if not filename:
            filename = self.filename
        self.doc.SaveAs(filename)
        self.close()
        
 
    def exportAsPDF(self, filename):
        self.doc.ExportAsFixedFormat(filename, constants.wdExportFormatPDF,
                                     Item = constants.wdExportDocumentWithMarkup, 
                                     CreateBookmarks = constants.wdExportCreateHeadingBookmarks)

    def close(self, SaveChanges=0):
        self.doc.Close(SaveChanges)
        self.engine.Quit()
        
        
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
        self.engine = win32com.client.Dispatch('Excel.Application')
        self.engine.Visible = True
        if filename:
            self.filename = filename
            self.wb = self.engine.Workbooks.Open(filename)
        else:
            self.filename = 'newfile.xlsx'
            self.wb = self.engine.Workbooks.Add()
            
    def select(self, sheet_name=''):
        ''' select and return a sheet '''
        if sheet_name:
            self.ws = self.wb.Worksheets(sheet_name)
        else:
            self.ws = self.wb.Worksheets(1)
        return self.ws
            
    def visible(self, visible=True):
        '''
        if Visible is true, the applicaion is visible
        '''
        self.engine.Visible = visible
        
    def save(self, filename=None):
        '''
        if filename is None, save the openning file
        else save as another file used the given name
        '''
        if filename:
            self.wb.SaveAs(filename)
        else:
            self.wb.Save()
            
    def close(self, SaveChanges=1):
        '''
        Close the application
        '''
        self.wb.Close(SaveChanges)
        
    def quit(self):
        self.engine.Quit()
        
    def get_cell_value(self, row, col):
        '''
        Get value of one cell
        '''
        return self.ws.Cells(row, col).Value
    
    def set_cell_value(self, row, col, value):  
        '''
        Set value of one cell
        '''
        self.ws.Cells(row, col).Value = value 
        
    def get_range_value(self, row1, col1, row2, col2):
        '''
        Return a 2d array (i.e. tuple of tuples)
        '''
        return self.ws.Range(self.ws.Cells(row1,col1), self.ws.Cells(row2,col2)).Value
    
    def set_range_value(self, leftCol, topRow, data):
        '''
        Insert a 2d array starting at given location.
        i.e. [['a','b','c'],['a','b','c'],['a','b','c']]
        Works out the size needed for itself
        '''
        bottomRow = topRow + len(data) - 1
        rightCol = leftCol + len(data[0]) - 1
        self.ws.Range(self.ws.Cells(topRow, leftCol), self.ws.Cells(bottomRow, rightCol)).Value = data
        
    def max_row(self, sheet=''):
        ws = self.wb.Worksheets(sheet) if sheet else self.ws
        return ws.UsedRange.Rows.Count