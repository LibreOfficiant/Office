# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno

import Office
from Office import UiDesktop as GUI
from Office.UiDesktop import MsgBox

from com.sun.star.awt.MessageBoxType import MESSAGEBOX, INFOBOX, WARNINGBOX, ERRORBOX, QUERYBOX
from com.sun.star.awt.MessageBoxButtons import \
    BUTTONS_OK, BUTTONS_OK_CANCEL, BUTTONS_YES_NO, BUTTONS_YES_NO_CANCEL, \
    BUTTONS_RETRY_CANCEL, BUTTONS_ABORT_IGNORE_RETRY, \
    DEFAULT_BUTTON_OK , DEFAULT_BUTTON_CANCEL, DEFAULT_BUTTON_RETRY, \
    DEFAULT_BUTTON_YES, DEFAULT_BUTTON_NO, DEFAULT_BUTTON_IGNORE
from com.sun.star.awt.MessageBoxResults import OK, YES, NO, CANCEL, RETRY, IGNORE


import uno
from com.sun.star.uno import RuntimeException as _rtex
def xray(myObject):
  try:
    sm = uno.getComponentContext().ServiceManager
    mspf = sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", uno.getComponentContext())
    scriptPro = mspf.createScriptProvider("")
    xScript = scriptPro.getScript("vnd.sun.star.script:XrayTool._Main.Xray?language=Basic&location=application")
    xScript.invoke((myObject,), (), ())
    return
  except:
     raise _rtex("\nBasic library Xray is not installed", uno.getComponentContext())

def test():
    s = Office.Session()
    MsgBox(s.OfficeName)
    ctx = uno.getComponentContext()
    office = ctx.ServiceManager.createInstanceWithContext(
        "Office",ctx)
    xray(office)
    #office.UiDesktop.MsgBox("dghqh")
    

def callUiDesktopMethods():
    s = Office.Session()
    MsgBox( "UiDesktop.MsgBox static method call" )
    msg = __file__
    msg += "\n%s" % UI.convertFromURL(__file__)
    msg += "\n\n%s" % s.SysExecutable
    msg += "\n%s" % UI.convertToURL(s.SysExecutable)
    MsgBox( msg, "This program name & application executable are")

def _CreateMessageBoxTypes():
    MsgBox('MESSAGEBOX', box_type=MESSAGEBOX)
    MsgBox('INFOBOX', box_type=INFOBOX)
    MsgBox('WARNINGBOX', box_type=WARNINGBOX)
    MsgBox('ERRORBOX', box_type=ERRORBOX)
    MsgBox('QUERYBOX', box_type=QUERYBOX)

    s = Office.Session()
    MsgBox("API UNO MessageBox", box_type=WARNINGBOX, buttons=BUTTONS_YES_NO_CANCEL )
    MsgBox("Information text", s.OfficeName, ERRORBOX )

#from com.sun.star.ui.dialogs.TemplateDescription import \
    #FILEOPEN_SIMPLE, FILESAVE_SIMPLE, FILESAVE_AUTOEXTENSION
#from Office.devTools import mri,xray
#def FileDialogs():
    #ui = GUI()
    #ui.openFileDialog("default FilePicker service")
    #MsgBox(ui.Files[0])
    #MsgBox(ui.DisplayDirectory)
    #xray(ui.fp)
    #MsgBox("")
    #ui.MultiSelectionMode = True
    #ui.openFileDialog('openFileDialog() - Opening zero-to-many files')




