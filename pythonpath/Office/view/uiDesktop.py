# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno, unohelper
from com.sun.star.task import XJobExecutor

from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import \
    BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, \
    BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY, \
    DEFAULT_BUTTON_OK , DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, \
    DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL, RETRY, IGNORE

from com.sun.star.ui.dialogs.ExecutableDialogResults import OK as _DLG_OK

#from ..devTools import xray

from .. import Session
_s = Session()
_MY_OFFICE = _s.OfficeName  # aka 'LibreOffice, OpenOffice, ...'

from com.sun.star.ui.dialogs.TemplateDescription import \
    FILEOPEN_SIMPLE, FILESAVE_SIMPLE, FILESAVE_AUTOEXTENSION

class UiDesktop(unohelper.Base, XJobExecutor, object):
    """ UiDesktop : Ensures a LibreOffice valid environment is present
    to serve GUI useful requests.
    Credits:
    https://wiki.openoffice.org/wiki/Python/Transfer_from_Basic_to_Python
    https://wiki.openoffice.org/wiki/FR/Documentation/Python/Transfer_from_Basic_to_Python
    https://forum.openoffice.org/fr/forum/viewtopic.php?f=8&t=53097 for FilePicker bugs
    """
    #ctx = uno.getComponentContext()
    #dsk = UIDesktop.createUnoService("com.sun.star.frame.Desktop")
    def __init__(self):
        self.ctx = uno.getComponentContext()
        self.desktop = self._createUnoService("com.sun.star.frame.Desktop")
        self.fp = self._createUnoService("com.sun.star.ui.dialogs.OfficeFilePicker")
        #self.fp.initialize((FILEOPEN_SIMPLE,))
        """_model = self.desktop.getCurrentComponent()  # May return 'NoneType' object
        # if isinstance(model, 'NoneType'): raise Exception(RuntimeError)
        ui = _model.CurrentController.Frame.ContainerWindow  # Throws 'AttributeError' on 'CurrentController' if 'NoneType'
        self.ui = ui"""
    """                    PUBLIC METHODS                            """
    def set_CurrentFilter(self, value: str): self.fp.CurrentFilter(value)
    @property  # str[0..*]
    def Files(self): return self.fp.Files
    #def get_DisplayDirectory(self): return self.fp.DisplayDirectory
    #def set_DisplayDirectory(self, value): self.fp.DisplayDirectory = value
    @property
    def DisplayDirectory(self):return self.fp.DisplayDirectory
    @DisplayDirectory.setter  # URL
    def DisplayDirectory(self, value): self.fp.DisplayDirectory = value
    @property  # bool
    def MultiSelectionMode(self): return self.fp.MultiSelectionMode
    @MultiSelectionMode.setter  # bool (write-only)
    def MultiSelectionMode(self, value): self.fp.MultiSelectionMode = value
    @property  # str
    def Title(self): return self.fp.Title
    @Title.setter  # str
    def Title(self, value): self.fp.Title = value
    """                    PUBLIC METHODS                            """
    """                   PRIVATE METHODS                            """
    @staticmethod
    def convertFromURL(fileURL):
        return uno.fileUrlToSystemPath(fileURL)
    @staticmethod
    def convertToURL(systemPath):
        return uno.systemPathToFileUrl(systemPath)
    """def dialogBox(self):  # Office dialog libraries
        pass"""
    def appendFilter(self, name: str, filters: str):
        self.fp.appendFilter( name, filters)
    def editDocument(self):  # Db. item e.g. Book, Movie etc..
        pass
    def MessageBox(self, message, title=_MY_OFFICE, box_type=MESSAGEBOX, buttons=BUTTONS_OK ):
        """ Credits: Hubert Lambert
        Arguments are ordered to ease coding.
        They default 'close to that' of MsgBox Basic statement.
        Note: arg2, arg3 are mixed as buttons+box_type in Basic MsgBox.
        message:  str - 
        title:    str - 
        box-type: int - cf.com.sun.star.awt.MessageBoxType
        buttons:  int - cf. com.sun.star.awt.MessageBoxButtons
        return:   int - cf. com.sun.star.awt.MessageBoxResults
        """
        tk = self._createUnoService("com.sun.star.awt.Toolkit")
        win = self.desktop.getCurrentFrame().ContainerWindow
        box = tk.createMessageBox(win, box_type, buttons, title, message)
        return box.execute()
    """@staticmethod
    def createUnoService(serviceName, arguments=None):
        ctx = uno.getComponentContext()
        sm = ctx.ServiceManager
        return sm.createInstanceWithContext(serviceName, ctx)"""
    def pickDirectory(self, title):
        pass
    @staticmethod
    def MsgBox( msg, title=_MY_OFFICE, box_type=MESSAGEBOX, buttons=BUTTONS_OK):
        ctx = uno.getComponentContext()
        sm = ctx.ServiceManager
        dsk = sm.createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        tk = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
        win = dsk.getCurrentFrame().ContainerWindow
        box = tk.createMessageBox(win, box_type, buttons, title, msg)
        return box.execute()
    def openFileDialog(self):  # i.e. com.sun.star.ui.OfficeFilePicker
        #xray(self.fp)
        #self.fp.initialize((FILEOPEN_SIMPLE,))
        if self.fp.execute() == _DLG_OK:
            pass  # xray(self.fp.Files)
        self.fp.dispose
        return
    def saveFileDialog(self):  # i.e. com.sun.star.ui.OfficeFilePicker
        #self.fp.initialize((FILESAVE_SIMPLE,))
        return
    """                    PUBLIC METHODS                            """
    """                   PRIVATE METHODS                            """
    def _createUnoService(self, serviceName, arguments=None):
        sm = self.ctx.ServiceManager
        return sm.createInstanceWithContext(serviceName, self.ctx)



class UiFilms(UiDesktop):
    @property
    def currentDataSource(self):
        _desktop = XSCRIPTCONTEXT.getDesktop()
        _model = _desktop.getCurrentComponent()
        if hasattr(_model, 'DataSource'):
            print(_model.Title, _model.DataSource.URL)
            return _model.DataSource
        #else:
            #_model = _desktop.loadComponentFromURL(
                #"private:factory/base", "_blank", 0, ())
    @property
    def currentRecord(self):  # Db. item e.g. Book, Movie etc..
        pass
    """                    PUBLIC METHODS                            """
    def composeDocument(self):  # Db. item e.g. Book, Movie etc..
        pass
    def editDocument(self):  # Db. item e.g. Book, Movie etc..
        pass



#region Fonctions Basic dédiées à l'API

def GetProcessServiceManager():
    return XSCRIPTCONTEXT.getComponentContext()
def CreateUnoService(service):
    pass
def StarDesktop():
    return XSCRIPTCONTEXT.getDesktop()
def ThisComponent():
    return XSCRIPTCONTEXT.getDocument()
def ThisDatabase():
    _model = StarDesktop().getCurrentComponent()
    if hasattr(_model, 'DataSource'):
        return _model.DataSource
def GetDefaultContext():
    return XSCRIPTCONTEXT.getComponentContext()
def CreateUnoStruct(typename, args, kwargs):
    return uno.createUnoStruct(typename)
def CreateObject(OLE_COM_object_name):
    pass
def CreateUnoDialog():
    pass
def CreateUnoValue(UNO_type, Basic_value):
    pass
def CreateUnoListener():
    pass
def IsUnoStruct():
    pass
def EqualUnoObjects():
    pass
def HasUnoInterfaces(obj, i1, i2):
    pass

# Credits: « Programmation OpenOffice & LibreOffice.. » B.Marcelly & L.Godard Eyrolles

#endregion

import unohelper
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation( \
    UiDesktop, "Libre.Office.UiDesktop", \
    ("Libre.Office.UiDesktop",),)

# Quick n Dirty Testing
if __name__ == '__main__':

    ui = UiDesktop()
    result = ui.MsgBox(MESSAGEBOX, BUTTONS_OK, "Here the title", "Here the content of the message")
    result = ui.MsgBox(QUERYBOX, BUTTONS_OK, 'titre', 'texte')
    if result == OK:
        print("OK")

    ds = ui.currentDataSource
    if hasattr(ds, 'URL'):
        print( ds.URL)  #, ds.QueryDefinitions, ds.Tables)
    print(uno.isInterface(ds))
