# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno
import unohelper

""" NOTE: Methods are not signed to be Python 2.7 compatible         """

from . import Session
_s = Session()
_MY_OFFICE = _s.OfficeName  # aka 'LibreOffice, OpenOffice, ...'

""" MRI                        BEGIN                             MRI """
import uno
def Mri_test():
	obj = uno.getComponentContext()
	mri(obj)
def mri(target):
	ctx = uno.getComponentContext()
	mri = ctx.ServiceManager.createInstanceWithContext("mytools.Mri", ctx)
	if mri is None:  # i.e. MRI is not installed or deactivated
		return  # exit silently
	mri.inspect(target)
""" MRI                         END                              MRI """

""" XRAY                       BEGIN                            XRAY """
import uno
from com.sun.star.uno import RuntimeException as _rtex
def xray(myObject):
	try:
		xScript = getScript("Xray", module="_Main", library="XrayTool")
		xScript.invoke((myObject,), (), ())
		return
	except:
		raise _rtex("\nBasic library Xray is not installed", uno.getComponentContext())
def demoXray(): 
	""" example : using xray in a Writer document """
	# RAISES Attribute Error When no Writer document
	xTextDoc = XSCRIPTCONTEXT.getDocument()
	xText = xTextDoc.Text 
	xray(xText)
	xray("Demo is finished") 
""" XRAY                        END                             XRAY """

""" OI                         BEGIN                              OI """
import uno
def inspectObject( obj, title=_MY_OFFICE ):
	ctx = uno.getComponentContext()
	oi = ctx.ServiceManager.createInstanceWithContext("org.openoffice.InstanceInspector", ctx)
	if oi is None:  # i.e. InstanceInspector is not installed or deactivated
		return  # exit silently
	oi.inspect( obj, title )
def exObjInspector():
	doc = XSCRIPTCONTEXT.getDocument()
	dsk = XSCRIPTCONTEXT.getDesktop()
	inspectObject(doc, "Document Inspector")
	inspectObject(dsk, title="StarDesktop")
""" OI                         BEGIN                              OI """

""" PYTHON2BASIC               BEGIN                    PYTHON2BASIC """

""" MsgBox Types, Buttons, Default & Result enumerations
Basic constant names are not used on purpose..
"""
class BoxType(object):
	""" MsgBox Types enumeration
	MB_ICON: STOP, QUESTION, EXCLAMATION, INFORMATION in Basic
	"""
	MSGBOX = 0       # No Sign        MessageBoxTypes.MESSAGEBOX = 0
	ERRORBOX =16     # Stop sign        ERRORBOX = 3 
	QUERYBOX = 32    # Question mark    QUERYBOX = 4
	WARNINGBOX = 48  # Warning sign     WARNINGBOX = 2
	INFOBOX = 64     # Bulb sign        INFOBOX = 1
class BoxButtons(object):
	""" MsgBox Buttons enumeration
	MB_: OK, OKCANCEL, ABORTRETRYIGNORE,
	YESNOCANCEL, YESNO, RETRYCANCEL in Basic
	"""
	OK = 0                  #
	OK_CANCEL = 1           # MessageBoxButtons.BUTTONS_OK
	ABORT_RETRY_IGNORE = 2  # ..BUTTONS_OK_CANCEL
	YES_NO_CANCEL = 3       # ..BUTTONS_YES_NO
	YES_NO = 4              # ..BUTTONS_YES_NO_CANCEL
	RETRY_CANCEL = 5        # ..BUTTONS_RETRY_CANCEL
	# = 6                   # ..BUTTONS_ABORT_IGNORE_RETRY
class BoxDefault(object):
	""" MsgBox Default Button enumeration
	MB_DEF..BUTTON1, BUTTON2, BUTTON3 in Basic
	"""
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

from com.sun.star.lang import XMultiServiceFactory
from com.sun.star.script.provider import \
    ScriptFrameworkErrorException, \
    XScriptProviderFactory, \
    XScriptProvider, \
    XScript
def getScript(script, library='_Basic', module='devTools'):
    """ Locates & loads any Basic macro/script
    Exporting/Importing '_Basic' library and deactivating it is simply
    hiding it from Basic IDE. Functions calls keep operating fine.
    script:  str - Script/macro name
    library: str - Script library name
    module:  str - Script module name
    Credit: B.Marcelly in 'Calling XRay from Python'
    """
    try:
        sm = uno.getComponentContext().ServiceManager
        mspf = sm.createInstanceWithContext("com.sun.star.script.provider.MasterScriptProviderFactory", uno.getComponentContext())
        scriptPro = mspf.createScriptProvider("")
        scriptName = "vnd.sun.star.script:"+library+"."+module+"."+script+"?language=Basic&location=application"
        xScript = scriptPro.getScript(scriptName)  # raises ScriptFrameworkErrorException
        return xScript
    except ScriptFrameworkErrorException:
        raise _rtex("\n Basic script '%s.%s.%s' is not installed" % (library, module, script), uno.getComponentContext())
    except RuntimeException:
        raise _rtex("\n An unexpected error occurred: ", uno.getComponentContext())

def MsgBox(message, type_buttons_default=BoxButtons.OK , title=_MY_OFFICE):
    """ LibreOffice Basic 'MsgBox' statement call.
    message: str
    type_buttons_default: int - cf. BoxType.xx + Box.BUTTONS_xx + Box.DEFAULT_xx
    title:                str
    return:               int - cf. Box.RESULT_xx
    Credits:
    cf. https://wiki.openoffice.org/wiki/FR/Documentation/BASIC_Guide/Message_and_Input_Boxes_(Runtime_Library)
    cf. " Programmation OpenOffice.." Bernard Marcelly, Eyrolles
    """
    xScript = getScript("_MsgBox")  # cf. XScript interface API documentation
    try:
        res = xScript.invoke((message,type_buttons_default,title), (), ())
        return res[0]  # 3-uple containing (aOutParam, aOutParamIndex, (aParams aka *args))
        """ Potential memory leak due to NEW output :params RETURNED BACK """
    except:
        raise _rtex("_MsgBox illegal argument call", uno.getComponentContext())
def InputBox(prompt, title=_MY_OFFICE, defaultValue=''):
    """ LibreOffice Basic 'InputBox' statement call.
    prompt: str - 
    title: str  -
    return: str - User input or '' empty string.
    """
    xScript = getScript("_InputBox")
    try:
        res = xScript.invoke((prompt,title,defaultValue), (), ())
        """ Potential memory leak due to EXTRA output :params RETURNED BACK """
        return res[0]
    except:
        raise _rtex("_InputBox illegal argument call", uno.getComponentContext())
""" PYTHON2BASIC                END                     PYTHON2BASIC """

g_exportedScripts = demoXray, Mri_test, exObjInspector

""" Change Log aka Summary of changes

Date		Full name			Description																		Ref. #
2017-Apr-16 Alain H. Romedenne  BoxXx.XX constant enumerations
2017-Apr-04 Alain H. Romedenne  Python v2 refactoring (no signatures in methods)
2017-Mar-29 Alain H. Romedenne  Replaced XSCRIPTCONTEXT.getComponentContext() by uno.getComponentContext()
2017-Mar-29 Alain H. Romedenne  Imported Session.OfficeName i.e. 'LibreOffice, OpenOffice, ..'
2017-Feb-14 Alain H. Romedenne  Module creation																	......

"""

if __name__ == '__main__':
	""" Python IDEs Tests """
	from Office.bridge import XSCRIPTCONTEXT
	demoXray()  # Geany: KO
	Mri_test()
	exObjInspector()

""" Basic scripts from 'devTools' module in '_Basic' library are:

Private Function _MsgBox( msg As String, Optional options As Integer, Optional title As String ) As Integer
	_MsgBox = MsgBox( msg, options, title )
End Function

Private Function _InputBox( prompt As String, Optional title As String, Optional defaultInput As String) As String
	_InputBox = InputBox( prompt, title, defaultInput )
End Function

"""
