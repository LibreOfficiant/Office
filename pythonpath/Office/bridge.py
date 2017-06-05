""" Module header

Module name : _bridge
Purpose     : Substitute to XSCRIPTCONTEXT when running Python scripts outside of (Libre|Open)Office.
Author      : Alain H. Romedenne
Description : The XSCRIPTCONTEXT object facilitates connections to running instances of (Libre|Open)Office.
    It connects to a piped instance of (Libre|Open)Office named 'LibreOffice'
    It connects to a localhost instance of (Libre|Open)Office using port=yyyy (yyyy=current year)
    It can connect to any variation of pipe name or port# if started.
    .ComponentContext, .CurrentDocument, .Desktop properties can be used.
    .connect(..): obj, .createInstance(name: str), .createUNOService(service: str) methods are available.
Intended    : Enable SSH connections to remote instances of (Libre|Open)Office.

Usage       :
    from Office._bridge import XSCRIPTCONTEXT  #  Please comment this line when running within (Libre|Open)Office

Credits     :
    christopher5106.github.io/office/2015/12/06/openoffice-libreoffice-automate-your-office-tasks-with-python-macros.html
    http://www.linuxjournal.com/content/starting-stopping-and-connecting-openoffice-python
    pyoo (c) 2014 Seznam.cz, a.s.

"""

#region Imports, constants

import datetime
import sys
# import socket  # only needed on win32-OOo3.0.0

#region UNO imports
import uno
# UNO interfaces
from com.sun.star.script.provider import XScriptContext                     # Script Context
from com.sun.star.frame import XModel                                       # - Current document object
from com.sun.star.document import XScriptInvocationContext                  # - Invocation dependent object
from com.sun.star.frame import XDesktop                                     # - Desktop object
from com.sun.star.uno import XComponentContext                              # - Component object
# UNO Exceptions
from com.sun.star.uno import RuntimeException
from com.sun.star.lang import IllegalArgumentException
from com.sun.star.connection import NoConnectException
#from com.sun.star.container import NoSuchElementException

_NoConnectException = uno.getClass('com.sun.star.connection.NoConnectException')
_IllegalArgumentException = uno.getClass('com.sun.star.lang.IllegalArgumentException')
_RuntimeException = uno.getClass('com.sun.star.uno.RuntimeException')
# We try to catch them and to re-throw Python standard exceptions, if running outside of (Libre|Open)Office

#endregion

_VERBOSE = False
_MY_OFFICE = 'LibreOffice'
_MY_HOST = 'localhost'
_MY_PORT = datetime.date.today().year
_MY_PIPE = _MY_OFFICE
NOCONNECT_EXCEPTION = 'Failed to connect to %s on host=%s,port=%d,pipe=%s'
ILLEGALARGUMENTCALL_EXCEPTION = ''
RUNTIME_EXCEPTION = ''
NODESKTOP_EXCEPTION = 'Failed to create % desktop on host=%s,port=%d,pipe=%s'
_UNSUPPORTED_PLATFORM_MSG = "'%s' platform is not supported yet."
_NOSTARTUP_EXCEPTION_MSG = 'Failed to start %s on host=%s,port=%d,pipe=%s'

#endregion

if _VERBOSE: print(datetime.date.today())

#region Object-Oriented Code
class Singleton(type):
    """
    A Singleton design pattern
    Credits: « Python in a Nutshell » by Alex Martelli, O'Reilly
    """
    _instances = {}
    def __call__(cls, *args, **kwargs):
        if cls not in cls._instances:
            cls._instances[cls] = super(Singleton, cls).__call__(*args, **kwargs)
            return cls._instances[cls]
        #else:  # Remove comments if you expect explicit coding rejection.
        #   raise Exception("An instance of that object already exists.")

class _OfficeSession(XScriptContext, object, metaclass=Singleton,):
    """
    This class substitutes XSCRIPTCONTEXT when running Python scripts outside (Libre|Open)Office
    Usage:
    from LibreOffice import XSCRIPTCONTEXT  #  Please comment this line when running within (Libre|Open)Office
    """
    def __init__(self):  # Class constructor
        _OfficeSession.connect(hostname=_MY_HOST, port=_MY_PORT)
        _OfficeSession.connect(pipe=_MY_PIPE)
    #region Properties
    @property
    def ComponentContext(self) -> XComponentContext :
        """
        Connects to a running listening 'LibreOffice' IPC instance.
        :return: XSCRIPTCONTEXT substitute
        """
        return _OfficeSession._sessions[_MY_PIPE]
    @property
    def CurrentDocument(self) -> XModel :
        return self.getDocument() #  Office current base/IDE/calc/draw/impress/math/writer document Or 'None'
    @property
    def Desktop(self) -> XDesktop :
        return self.createInstance("com.sun.star.frame.Desktop")
    def InvocationContext(self) -> XScriptInvocationContext :
        return None
    #endregion

    #region Methods
    _sessions = {}  # (Libre|Open)Office listening IPC instances identified by pipe name or port #
    @staticmethod
    def connect(hostname=_MY_HOST, port=_MY_PORT, pipe=None) -> XComponentContext :
        """
        Connects to a running listening instance of (Libre|Open)Office
        e.g. ctx = obj.connect(port=2016) connects to 'started LibreOffice on port#2016' throws exception otherwise
        When specified pipe argument takes precedence over hostname:port arguments pair
        :param hostname: defaults to 'localhost'
        :param port: defaults to this year e.g. 2016
        :param pipe: .. 'LibreOffice' ..
        :return: last established context: defaults to 'LibreOffice' piped visible instance
        """
        local_context = uno.getComponentContext()
        resolver = local_context.ServiceManager.createInstanceWithContext(
            'com.sun.star.bridge.UnoUrlResolver', local_context)
        try:  # conn = 'pipe,name=%s' % pipe if pipe: else 'socket,host=%s,port=%d' % (hostname, port)
            if pipe:
                conn = 'pipe,name=%s' % pipe
            else:
                conn = 'socket,host=%s,port=%d' % (hostname, port)
            connection_url = 'uno:%s;urp;StarOffice.ComponentContext' % conn
            if _VERBOSE: print(connection_url)
            _established_context = resolver.resolve(connection_url)
        #except ConnectionSetupException:
        except NoConnectException:  # thrown when LibreOffice specified instance isn't started
            tb = sys.exc_info()[2]
            raise Exception(NOCONNECT_EXCEPTION % (_MY_OFFICE, hostname, port, pipe)).with_traceback(tb) \
                from _NoConnectException
        except IllegalArgumentException:
            tb = sys.exc_info()[2]
            raise Exception(ILLEGALARGUMENTCALL_EXCEPTION).with_traceback(tb) \
                from _NoConnectException
        except RuntimeException:
            tb = sys.exc_info()[2]
            raise Exception(RUNTIME_EXCEPTION).with_traceback(tb) \
                from _NoConnectException
        if pipe:
            _key = pipe
        elif port:
            _key = port
        else:
            _key = _MY_PORT
        _OfficeSession._sessions[_key] = _established_context
        if _VERBOSE: print(__name__ + ' (yyyy-mmm-dd hh:mm:ss) Connection to %s established with %s' % (_MY_OFFICE, conn))
        return _established_context
    def createInstance(self, name):
        _remote_context = self.ComponentContext
        obj = _remote_context.ServiceManager.createInstanceWithContext(name, _remote_context)
        return obj
    def createUNOService(self, name):  # as offered in Office Basic
        return self.createInstance(name)
    def getComponentContext(self):
        return self.ComponentContext
    def getDesktop(self):  # = Desktop()
        return self.Desktop
    def getDocument(self):
        return self.Desktop.CurrentComponent
    #endregion
    #region Events
    #endregion
"""
class _ExtendedSession(_OfficeSession):
    #region Properties
    #endregion
    #region Methods
    #endregion
"""
#endregion

XSCRIPTCONTEXT = _OfficeSession()

#region Functions
def createObject(class_name: str):
    """
    LibreOffice class objects factory.
    Substitute to LibreOffice Basic CreateUnoService(), equivalent to OLE/COM createObject()...
    Throws uno.RuntimeException
    :param class_name: e.g. 'com.sun.star.sdb.DatabaseContext'
    :return: a LibreOffice instance of 'class_name'.
    """
    ctx = XSCRIPTCONTEXT.getComponentContext()
    return ctx.getServiceManager().createInstanceWithContext(class_name, ctx)

def CreateObject(class_name: str):
    """
    LibreOffice class objects factory.
    Substitute to LibreOffice Basic CreateUnoService(), equivalent to OLE/COM createObject()...
    Throws uno.RuntimeException
    :param class_name: e.g. 'com.sun.star.sdb.DatabaseContext'
    :return: a LibreOffice instance of 'class_name'.
    """
    ctx = uno.getComponentContext()
    return ctx.getServiceManager().createInstanceWithContext(class_name, ctx)

#endregion


""" Change Log aka Summary of changes

Date        Full name           Description                                                                     Ref. #
2016-Dec-12 Alain H. Romedenne  Added createObject() public function in place of <session>.createXXX() methods
2016-Oct-15 Alain H. Romedenne  Module creation                                                                 ... ...

"""

# Quick n Dirty Testing
if __name__ == '__main__':

    ctx = XSCRIPTCONTEXT.ComponentContext
    dsk = XSCRIPTCONTEXT.getDesktop()
    doc = XSCRIPTCONTEXT.getDocument()  # May return None
    obj = createObject("com.sun.star.frame.Desktop")
    print(ctx)
    print(dsk)  # com.sun.star.uno.XInterface
    print(doc)  # com.sun.star.lang.XComponent - base, calc, draw, basicIDE, impress, math, writer, ..


    #  print(XSCRIPTCONTEXT.ComponentContext.ImplementationName)
    print(dsk.ImplementationName)  #  com.sun.star.comp.framework.Desktop
    if doc is not None:
        print(doc.ImplementationName)  # com.sun.star.comp.*

    u1 = uno.getComponentContext()
    u1
    u2 = uno.getCurrentContext()
    u2
