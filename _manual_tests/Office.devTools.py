# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno

from Office.devTools import mri, xray, inspectObject, \
	InputBox, MsgBox

""" MsgBox Types, Buttons, Default & Result enumerations
Basic constant names are not used on purpose..
"""
class BoxType(object):
	""" MsgBox Types enumeration """
	MSGBOX = 0       # No Sign        MessageBoxTypes.MESSAGEBOX = 0
	ERRORBOX =16     # Stop sign        ERRORBOX = 3 
	QUERYBOX = 32    # Question mark    QUERYBOX = 4
	WARNINGBOX = 48  # Warning sign     WARNINGBOX = 2
	INFOBOX = 64     # Bulb sign        INFOBOX = 1
class BoxButtons(object):
	""" MsgBox Buttons enumeration """
	OK = 0                  #
	OK_CANCEL = 1           # MessageBoxButtons.BUTTONS_OK
	ABORT_RETRY_IGNORE = 2  # ..BUTTONS_OK_CANCEL
	YES_NO_CANCEL = 3       # ..BUTTONS_YES_NO
	YES_NO = 4              # ..BUTTONS_YES_NO_CANCEL
	RETRY_CANCEL = 5        # ..BUTTONS_RETRY_CANCEL
	# = 6                   # ..BUTTONS_ABORT_IGNORE_RETRY
class BoxDefault(object):
	""" MsgBox Default Button enumeration """
	BUTTON1 = 0   # 1st button is selected by DEFAULT
	BUTTON2 = 256 # 2nd button is selected
	BUTTON3 = 512 # 3rd button is selected
class BoxResult(object):
	""" MsgBox Result code enumeration """
	# = 0        # MessageBoxResults.CANCEL
	OK = 1       # ..Results.OK 
	CANCEL = 2   # ..Results.YES
	ABORT = 3    # ..Results.NO
	RETRY = 4    # ..Results.RETRY
	IGNORE = 5   # ..Results.IGNORE
	YES = 6      #
	NO = 7       #

def Mri_test():
	MsgBox(Box.ERRORBOX)
	obj = uno.getComponentContext()
	mri(obj)
def demoXray(): 
	""" example : using xray in a Writer document """
	# RAISES Attribute Error When no Writer document
	xTextDoc = XSCRIPTCONTEXT.getDocument()
	xText = xTextDoc.Text 
	xray(xText)
	xray("Demo is finished") 
def exObjInspector():
	doc = XSCRIPTCONTEXT.getDocument()
	dsk = XSCRIPTCONTEXT.getDesktop()
	inspectObject(doc, "Document Inspector")
	inspectObject(dsk, title="StarDesktop")

def _Py2Basic():
	rc = MsgBox('MsgBox', title='Title', type_buttons_default=BoxButtons.RETRY_CANCEL)
	MsgBox(str(rc), title='rc')
def _MsgBox():
	MsgBox('BoxType.MSGBOX')  # aka MsgBox('BoxType.MSGBOX', type_buttons_default=0)
	MsgBox('BoxType.ERRORBOX', type_buttons_default=BoxType.ERRORBOX)
	MsgBox('BoxType.QUERYBOX', type_buttons_default=BoxType.QUERYBOX)
	MsgBox('BoxType.WARNINGBOX', type_buttons_default=BoxType.WARNINGBOX)
	MsgBox('BoxType.INFOBOX', type_buttons_default=BoxType.INFOBOX)

	MsgBox('BoxButtons.OK', type_buttons_default=BoxButtons.OK)
	MsgBox('BoxButtons.OK_CANCEL', type_buttons_default=BoxButtons.OK_CANCEL)
	MsgBox('BoxButtons.ABORT_RETRY_IGNORE', type_buttons_default=BoxButtons.ABORT_RETRY_IGNORE)
	MsgBox('BoxButtons.YES_NO_CANCEL', type_buttons_default=BoxButtons.YES_NO_CANCEL)
	MsgBox('BoxButtons.YES_NO', type_buttons_default=BoxButtons.YES_NO)
	MsgBox('BoxButtons.RETRY_CANCEL', type_buttons_default=BoxButtons.RETRY_CANCEL)

	MsgBox('BoxButtons.YES_NO_CANCEL+BoxDefault.BUTTON3', type_buttons_default=BoxButtons.YES_NO_CANCEL+BoxDefault.BUTTON3)
	MsgBox('BoxButtons.RETRY_CANCEL+BoxDefault.BUTTON2', type_buttons_default=BoxButtons.RETRY_CANCEL+BoxDefault.BUTTON2)

	MsgBox('Message text.',title='Python.Title')
def _InputBox():
	txt = InputBox("Please press 'Esc'" )
	MsgBox("%s was entered." % txt, "InputBox")
	txt = InputBox('Please enter text:', 'Title')
	MsgBox("%s was entered." % txt, "InputBox")
	txt = InputBox('Please enter a number:', 'Title', '5935')
	MsgBox("%s was entered." % txt, "InputBox")






