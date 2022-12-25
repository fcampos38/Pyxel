"""Creates a python module to interact with Microsoft Excel Workbooks on Windows Machines"""

'''
Required Imports
'''
import win32com.client as win32
import numpy as np
from win32com.universal import com_error
import logging
import re
import string
import itertools

'''
Creates a customized exception to use during the program
'''
class PyxelException(Exception):
    pass

'''
Creates a python class to represent a Microsoft Excel workbook
A workbook object will result from the instantiation of this class
'''
class Excel():
    def __init__(self, PathToWorkbook: str, BackgroundExecution=True, OverWriteIfExists=False):
        """
        :param: PathToWorkbook: string representing the full path of the excel workbook
        :param: BackgroundExecution: boolean stating whether or not the Dispatch should run on the background
        :param: OverWriteIfExists: boolean stating whether or not the workbook shold be overwritten if it already exists
        
        :return: An excel workbook object
        """
        
        self.PathToWorkbook = PathToWorkbook # Workbook's path attribute
        self.XlConstants = win32.constants  # Excel constants attribute

        # Tries to dispatch Microsoft Excel in order to interact with it
        try:
            self.XlDispatch = win32.gencache.EnsureDispatch('Excel.Application')
        
        except Exception as error1:
            raise PyxelException(f"Microsoft Excel would not respond. An excel object could be be instantiated: {error1}")
        
        # Chooses whether to run the dispatch commands in the background or in visible mode
        self.xldispatch.Visible = not(BackgroundExecution)

        # Initiates a workbook object
        # Opens it if exists and OverWriteIfExists is False, otherwise creates a new workbook
        try:
            if not OverWriteIfExists:
                # Tries to open the workbook
                self.WorkbookObj = self.XlDispatch.Workbooks.Open(self.PathToWorkbook)
                
            else:
                raise PyxelException()
                
        except com_error as error1:
            # If non-existing code is raised on exception, tries to create a new one
            if error1.args[0] == -2147352567 or OverWriteIfExists:
                try:
                    new_workbook = self.XlDispatch.Workbooks.Add()
                    new_workbook.SaveAs(self.PathToWorkbook)
                    self.WorkbookObj = self.XlDispatch.Workbooks.Open(self.PathToWorkbook)

                except Exception as error2:
                    raise PyxelException(
                        f"The workbook {'does not exist and ' if error1.args[0] == -2147352567 else ''}could not be created or opened: {error2}"
                                    )
                    
            else:
                raise PyxelException(f"The workbook could not be opened: {error1}")
        
        # Gets all workbook's sheets into an array like attribute
        self.worksheets = tuple((sheet for sheet in self.WorkbookObj.Sheets))
        
                    
    def __del__(self):
        # Quits Microsoft Excel (terminates the dispatch command)
        self.XlDispatch.Quit()
        
    def __repr__(self):
        return chr(39) + f'ExcelWorkbookObject for "{self.PathToWorkbook}"' + chr(39)
        
