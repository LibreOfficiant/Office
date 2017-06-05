# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno, platform, os, sys
from com.sun.star.uno import XComponentContext


_NODE_OFFICE = "/org.openoffice.Setup/Office"
_NODE_PRODUCT = "/org.openoffice.Setup/Product"
_NODE_L10N = "/org.openoffice.Setup/L10N"  # L(ocalizatio)n

_PATH_SETS ="com.sun.star.util.PathSettings"
_PATH_SUBS ="com.sun.star.util.PathSubstitution"

""" _Configurations & _Paths are not inherited by Session class thus 
enforcing code maintenance. They are encapsulated instead, which permits
to exhibit the most interesting properties.
Encapsulation also facilitates x-product compatibility between
LibreOffice & OpenOffice.
"""  

class _Configurations(object):  # com.sun.star.configuration.* 
    """ Configuration: Represents (Libre/Open)Office.Session
    properties & methods
    """
    """ CONSTRUCTOR """
    def __init__(self):
        self.ctx = uno.getComponentContext()
        self.l10n = self._getConfigurationAccess(_NODE_L10N)
        self.office = self._getConfigurationAccess(_NODE_OFFICE)
        self.product = self._getConfigurationAccess(_NODE_PRODUCT)
    """                   PUBLIC PROPERTIES                          """
    """                    PUBLIC  METHODS                           """

    """           PRIVATE PROPERTIES - INTERNAL USE ONLY             """
    """ L10N // L(ocalisatio)n information as of Office UI menu:
    Tools - Options - Linguistics - Language
                                                        cf. Issue. 93203
    """
    @property  # bool
    def _L10NDecimalSeparatorAsLocale(self): return self.l10n.DecimalSeparatorAsLocale
    @property  # e.g. 'fr'
    def _L10NLocale(self): return self.l10n.ooLocale
    @property  #                                                   EMPTY
    def _L10NSetupCurrency(self): return self.l10n.ooSetupCurrency
    @property  # e.g. 'fr-FR'                            BUGGED with 5.3
    def _L10NSetupSystemLocale(self): return self.l10n.ooSetupSystemLocale
        #return self.intl.getByName('ooSetupSystemLocale')
    @property  # bool                                   LIBREOFFICE ONLY
    def _L10NIgnoreLanguageChange(self): return self.l10n.IgnoreLanguageChange
    @property  # ';' separated tuple                    LIBREOFFICE ONLY
    def _L10NDateAcceptancePatterns(self): return self.l10n.DateAcceptancePatterns
    
    """ OFFICE node information """
    @property  # service object
    def _OfficeFactories(self): return self.office.Factories
    @property  # service object
    def _OfficeInstalledLocales(self): return self.office.InstalledLocales
    @property  # str
    def _OfficeLastCompatibilityCheckID(self): \
        return self.office.LastCompatibilityCheckID
    @property  # bool
    def _OfficeMigrationCompleted(self): return self.office.MigrationCompleted
    @property  #                                                   EMPTY
    def _OfficeSetupConnectionURL(self): return self.office.ooSetupConnectionURL
    @property  # bool
    def _OfficeSetupInstCompleted(self): return self.office.ooSetupInstCompleted
    """ PRODUCT node STATIC information 
    Note: Setters aren't required here.
    """
    @property  # e.g. 'LibreOffice, OpenOffice'
    def _ProductName(self): return self.product.ooName  # aka product.getByName('ooName')
    @property
    def _ProductOpenSourceContext(self): return self.product.ooOpenSourceContext
    @property  # aka '.5.1, '
    def _ProductSetupExtension(self): return self.product.ooSetupExtension
    @property  # e.g. '5.2.5.1, 4.3.1'
    def _ProductSetupVersionAboutBox(self): 
        return self.product.getByName('ooSetupVersionAboutBox')  # aka product.ooSetupVersionAboutBox
    @property  # e.g. '.5.1'                            LIBREOFFICE ONLY
    def _ProductSetupVersionAboutBoxSuffix(self):
        return self.product.getByName('ooSetupVersionAboutBoxSuffix')  #.ooSetupVersionAboutBoxSuffix
    @property  # e.g. 'Document Foundation, Apache Software Foundation'
    def _ProductVendor(self): return self.product.ooVendor  # aka reader.getByName('ooVendor')
    @property  # e.g. '5.2, 4.3.1'
    def _ProductVersion(self): return self.product.getByName('ooSetupVersion')  # aka product.ooSetupVersionreader
    @property  # e.g. 'OpenOffice.org'
    def XMLFileFormatName(self): return self.product.ooXMLFileFormatName
    @property  # e.g. '1.0'
    def XMLFileFormatVersion(self): return self.product.ooXMLFileFormatVersion

    """            PRIVATE METHODS- INTERNAL USE ONLY                """
    def _getConfigurationAccess(self, nodevalue, updatable=False):
        """ Credit: 'apso' by Hubert Lambert """
        cp = self.ctx.getServiceManager().createInstanceWithContext( 
            "com.sun.star.configuration.ConfigurationProvider", self.ctx)
        node = uno.createUnoStruct("com.sun.star.beans.PropertyValue")
        node.Name = "nodepath"
        node.Value = nodevalue
        if updatable:
            return cp.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationUpdateAccess", (node,))
        else:
            return cp.createInstanceWithArguments(
                "com.sun.star.configuration.ConfigurationAccess", (node,))


class _Paths(object):  # com.sun.star.util.*
    """ Paths: Provides (Libre/Open)Office.Session common paths-related
    properties & methods
    """
    """ CONSTRUCTOR """
    def __init__(self):
        ctx = uno.getComponentContext()
        sm = ctx.ServiceManager
        self.urls = sm.createInstanceWithContext(_PATH_SETS, ctx)
        self.subs = sm.createInstanceWithContext(_PATH_SUBS, ctx)
    """                   PUBLIC PROPERTIES                          """
    """                    PUBLIC  METHODS                           """

    """           PRIVATE PROPERTIES - INTERNAL USE ONLY             """
    """ PATHSETTINGS service STATIC information, mostly dirs URLs
    Credits: B.Marcelly 'Programmation... ' Part 4 """
    @property
    def _Addin(self): return self.urls.Addin  # aka dirs.getByName('ooName')
    @property
    def _AutoCorrect(self): return self.urls.AutoCorrect
    @property
    def _AutoText(self): return self.urls.AutoText
    @property
    def _Backup(self): return self.urls.Backup
    @property  # Current user Basic macros & dialogs
    def _Basic(self): return self.urls.Basic
    @property  # Toolbars' icons
    def _Bitmap(self): return self.urls.Bitmap
    @property  # (Libre/Open)Office configuration files
    def _Config(self): return self.urls.Config
    @property
    def _Dictionary(self): return self.urls.Dictionary
    @property
    def _Favorite(self): return self.urls.Favorite
    @property
    def _Filter(self): return self.urls.Filter
    @property
    def _Gallery(self): return self.urls.Gallery
    @property  # Image files
    def _Graphic(self): return self.urls.Graphic
    @property
    def _Help(self): return self.urls.Help
    @property  # Dictionaries
    def _Linguistic(self): return self.urls.Linguistic
    @property  # Program modules
    def _Module(self): return self.urls.Module
    @property
    def _Palette(self): return self.urls.Palette
    @property
    def _Plugin(self): return self.urls.Plugin
    @property  # User information about e-mail, ftp, ..
    def _Storage(self): return self.urls.Storage
    @property
    def _Temp(self): return self.urls.Temp
    @property  # Document templates
    def _Template(self): return self.urls.Template
    @property  # GUI configuration files
    def _UIConfig(self): return self.urls.UIConfig
    @property  # Current user configuration
    def _UserConfig(self): return self.urls.UserConfig
    @property  # Current user *Office documents
    def _Work(self): return self.urls.Work
    """ PATHSUBSTITUTION service information, mostly dirs URLs
    Credits: B.Marcelly 'Programmation... ' Part 4 """
    @property
    def _BrandBaseURL(self): return self._getVarValue('$(brandbaseurl)')
    @property  # User data directory
    def _HomeURL(self): return self._getVarValue('$(home)')
    @property
    def _InstURL(self): return self._getVarValue('$(inst)')  # aka '$(insturl)'
    @property  # Environment PATH variable, ';' separated
    def _PathURL(self): return self._getVarValue('$(path)')
    @property
    def _ProgURL(self): return self._getVarValue('$(prog)')  # aka '$(progurl)'
    @property  # Temporary files directory
    def _TempURL(self): return self._getVarValue('$(temp)')
    @property  # User directory
    def _UserURL(self): return self._getVarValue('$(user)')  #aka'$(userdataurl)
    @property  # Work directory
    def _WorkURL(self): return self._getVarValue('$(work)')
    @property  # UI language number e.g. France=33, USA=01
    def _UILangNbr(self): return self._getVarValue('$(lang)')  # ONLY OpenOffice.org 
    @property  # UI Language CODE e.g. France=1036
    def _UILangID(self): return self._getVarValue('$(langid)')
    @property  # UI Locale e.g. fr, en-US
    def _UILocale(self): return self._getVarValue('$(vlang)')
    @property  # User account
    def _UserName(self): return self._getVarValue('$(username)')  # ONLY LibreOffice
    """            PRIVATE METHODS- INTERNAL USE ONLY                """
    def _getVarValue(self, variableName ):
        return self.subs.getSubstituteVariableValue(variableName)

import unohelper
from com.sun.star.task import XJobExecutor
class Session( unohelper.Base, XJobExecutor, object):
    """ Session: Represents the environment of the current script, 
    providing access to configuration information, information about the
    current user, and information about the 'OpenOffice' platform and
    release number.
    """
    def __init__(self):
        """ ENCAPSULATION INSTEAD OF INHERITANCE """
        self.ctx = uno.getComponentContext()
        self.config = _Configurations()
        self.path = _Paths()
    """                   PUBLIC PROPERTIES                          """
    @property
    def CurrentDatabase(self): pass
    """@property
    def OfficeBuildVersion(self): pass"""
    @property  # e.g. 'fr' - The product (regional) language settings.
    def OfficeLocale(self): return self.config._L10NLocale
    @property  # e.g. 'LibreOffice, OpenOffice'
    def OfficeName(self): return self.config._ProductName
    @property  # e.g. 'The Document Foundation, Apache Software Foundation'
    def OfficeVendor(self): return self.config._ProductVendor
    @property  # e.g. '5.2, 4.3.1'
    def OfficeVersion(self): return self.config._ProductVersion
    """ os, platform, sys information """
    @property  # e.g. 'fr-FR'
    def OSLocale(self): return self.config._L10NSetupSystemLocale
    @property
    def OSPathBasename(self): return os.path.basename(__file__)
    @property
    def OSPathDirname(self): return os.path.dirname(__file__)
    @property
    def OSCurrentWorkDir(self): return os.getcwd()
    @property  # e.g. 'Darwin, Linux, Windows'
    def Platform(self): return platform.system()
    @property
    def PythonBuild(self): return platform.python_build()[1]
    @property
    def PythonCompiler(self): return platform.python_compiler()
    @property
    def PythonVersion(self): return "%s.%s.%s" % sys.version_info[:3]
    @property
    def SysExecutable(self): return sys.executable

    """                    PUBLIC  METHODS                           """
    def createUnoService(self, serviceName, arguments=None):
        """ 
        serviceName: str    - UNO service name
        arguments: sequence - Service dependant tuple.
        return: object      - UNO service
        """
        sm = self.ctx.ServiceManager
        return sm.createInstanceWithContext(serviceName, self.ctx)
    """def getInvocationContext(self):
        return self.InvocationContext"""

"""                      BEGIN EXTENSION                             """
import unohelper
from com.sun.star.task import XJobExecutor
class _UnoComponentOOoSession( Session, unohelper.Base, XJobExecutor):
    """ internal wrapper intended to distribute Session features to
    other languages e.g. Basic, Beanshell, JavaScript, etc..
    """
    def __init__(self, ctx):
        super().__init__()
        self.ctx = ctx
    def getOfficeName(self): return super().OfficeName
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation( \
    _UnoComponentOOoSession, \
    "Open.Office.Session", \
    ("com.sun.star.task.Job",),)
"""                       END EXTENSION                              """

""" Change Log aka Summary of changes

Date        Full name           Description                                                                     Ref. #
2017-Apr-11                     Created _Configurations, _Paths classes to facilitate maintenance
2017-Apr-09                     Completed 'L10n, Office, Product' configuration nodes information 
2017-Apr-06                     Added createUnoService() method. Credit: 'PythonScriptOrganizer' by Hanya
2017-Apr-06 Alain H. Romedenne  Shifted 'LibreOffice' properties & methods to Libre(Office)Session
2017-Apr-05 Alain H. Romedenne  Module creation      

"""

""" 
def createUnoService(serviceName):
    ctx = uno.getComponentContext()
    sm = ctx.ServiceManager
    return sm.createInstanceWithContext(serviceName, ctx)
from com.sun.star.awt.MessageBoxType import MESSAGEBOX
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
def MsgBox( message, title="Session.L10N", box_type=MESSAGEBOX, buttons=BUTTONS_OK ):
    tk = createUnoService("com.sun.star.awt.Toolkit")
    win = createUnoService("com.sun.star.frame.Desktop").getCurrentFrame().ContainerWindow
    box = tk.createMessageBox(win, box_type, buttons, title, message)
    return box.execute()
"""

if __name__ == '__main__':
    """ Quick test cf. UnitTests for exhaustive tests.
    """

    import sys
    sys.exit(main(sys.argv))

