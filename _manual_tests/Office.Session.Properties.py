
_LIBRE_OFFICE = 'LibreOffice'
_OPEN_OFFICE = 'OpenOffice'

def getSessionProperties():

    """ 
    SESSION PUBLIC PROPERTIES
    """
    s = Office.Session()
    msg = 'CurrentDatabase: %s' % s.CurrentDatabase  # ThisDatabaseDoc..
    """ (Open/Libre)Office information """
    #msg += '\nOfficeBuildVersion: %s' % s.OfficeBuildVersion
    msg += '\nOfficeLocale: %s  aka l10n.Locale' % s.OfficeLocale  # OS Language settings
    msg += '\nOfficeName: %s  aka product.Name' % s.OfficeName
    msg += '\nOfficeVendor: %s  aka product.Vendor' % s.OfficeVendor
    msg += '\nOfficeVersion: %s  aka product.Version' % s.OfficeVersion
    if s.OfficeName == _LIBRE_OFFICE:
        msg += '\nOfficeRevision: %s  aka product.SetupVersionAboutBox' % s.OfficeRevision
    """ os, platform information """
    msg += '\nOSLocale: %s  aka l10n.SetupSystemLocale' % s.OSLocale
    msg += '\nOSPathBasename: %s' % s.OSPathBasename
    msg += '\nOSPathDirname: %s' % s.OSPathDirname
    msg += '\nOSCurrentWorkDir: %s' % s.OSCurrentWorkDir
    msg += '\nPlatform: %s' % s.Platform
    """ Python information """
    msg += '\nPythonBuild: %s' % s.PythonBuild
    msg += '\nPythonCompiler: %s' % s.PythonCompiler
    msg += '\nPythonVersion: %s' % s.PythonVersion
    if s.OfficeName == _LIBRE_OFFICE:
        msg += '\nUserName: %s' % s.UserName    
    msg += '\nSysExecutable: %s' % s.SysExecutable

    MsgBox(msg, title="%s.Session" %s.OfficeName)

    """ 
    SESSION HIDDEN POTENTIAL PROPERTIES with conditional output
    """
    from Office.sessions.ooSession import _Configurations as Configs
    c = Configs()
    msg = ''
    if _L10N:  # L10n // L(ocalisatio)n GUI information // Tools - Options- ... menu
        msg += '\nl10n.DecimalSeparatorAsLocale: %s' % c._L10NDecimalSeparatorAsLocale  # bool
        msg += '\nl10n.Locale: %s  aka OfficeLocale' % c._L10NLocale  # str
        msg += '\nl10n.SetupCurrency: %s' % c._L10NSetupCurrency  # 
        msg += '\nl10n.SetupSystemLocale: %s  aka International, OSLocale' % c._L10NSetupSystemLocale  # OS locale settings
        if s.OfficeName == _LIBRE_OFFICE:
            msg += '\nl10n.IgnoreLanguageChange: %s' % c._L10NIgnoreLanguageChange  # bool
            msg += '\nl10n.DateAcceptancePatterns: %s' % c._L10NDateAcceptancePatterns  # str
            #mri(c.l10n)

    if _OFFICE:  # Office information
        #msg += '\noffice.Factories: %s' % c._OfficeFactories  # service object
        #msg += '\noffice.InstalledLocales: %s' % c._OfficeInstalledLocales  # service object
        #mri(c._OfficeFactories)  # API service object
        #mri(c._OfficeInstalledLocales)  # API service object
        msg += '\noffice.LastCompatibilityCheckID: %s' % c._OfficeLastCompatibilityCheckID
        msg += '\noffice.MigrationCompleted: %s' % c._OfficeMigrationCompleted
        msg += '\noffice.SetupConnectionURL: %s' % c._OfficeSetupConnectionURL
        msg += '\noffice.SetupInstCompleted: %s' % c._OfficeSetupInstCompleted

    if _PRODUCT:  # Product static information
        msg += '\nproduct.Name: %s' % c._ProductName
        msg += '\nproduct.ProductOpenSourceContext: %s' % c._ProductOpenSourceContext
        msg += '\nproduct.ProductSetupExtension: %s' % c._ProductSetupExtension
        msg += '\nproduct.SetupVersionAboutBox: %s' % c._ProductSetupVersionAboutBox
        if s.OfficeName == _LIBRE_OFFICE:
            msg += '\nproduct.SetupVersionAboutBoxSuffix: %s' % c._ProductSetupVersionAboutBoxSuffix
        msg += '\nproduct.Vendor: %s' % c._ProductVendor
        msg += '\nproduct.Version: %s' % c._ProductVersion
        msg += '\nproduct.XMLFileFormatName: %s' % c.XMLFileFormatName
        msg += '\nproduct.XMLFileFormatVersion: %s' % c.XMLFileFormatVersion

    if _L10N or _OFFICE or _PRODUCT:
        MsgBox(msg, title="%s._Configurations" %s.OfficeName)

    from Office.sessions.ooSession import _Paths as Util
    u = Util()
    msg = ''
    if _DIRS:  # *.util.PathSettings service information
        url = '\nPathSettings.Addin: %s' % u._Addin
        url += '\nPathSettings.AutoCorrect: %s' % u._AutoCorrect
        url += '\nPathSettings.AutoText: %s' % u._AutoText
        url += '\nPathSettings.Backup: %s' % u._Backup
        url += '\nPathSettings.Basic: %s' % u._Basic
        url += '\nPathSettings.Bitmap: %s' % u._Bitmap
        url += '\nPathSettings.Config: %s' % u._Config
        url += '\nPathSettings.Dictionary: %s' % u._Dictionary
        url += '\nPathSettings.Favorite: %s' % u._Favorite
        url += '\nPathSettings.Filter: %s' % u._Filter
        url += '\nPathSettings.Gallery: %s' % u._Gallery
        url += '\nPathSettings.Graphic: %s' % u._Graphic
        url += '\nPathSettings.Help: %s' % u._Help
        url += '\nPathSettings.Linguistic: %s' % u._Linguistic
        url += '\nPathSettings.Module: %s' % u._Module
        url += '\nPathSettings.Palette: %s' % u._Palette
        url += '\nPathSettings.Plugin: %s' % u._Plugin
        url += '\nPathSettings.Storage:g %s' % u._Storage
        url += '\nPathSettings.Temp: %s' % u._Temp
        url += '\nPathSettings.Template: %s' % u._Template
        url += '\nPathSettings.UIConfig: %s' % u._UIConfig
        url += '\nPathSettings.UserConfig: %s' % u._UserConfig
        url += '\nPathSettings.Work: %s' % u._Work
        MsgBox ( url, title="%s.Session.PathSettings" %s.OfficeName )
    if _VARS:  # *.util.PathSubstitution service information
        txt = 'PathSubst..BrandBaseURL: %s' % u._BrandBaseURL
        txt += '\n\nPathSubst..HomeURL: %s' % u._HomeURL
        txt += '\n\nPathSubst..InstURL: %s' % u._InstURL
        txt += '\n\nPathSubst..PathURL: %s' % u._PathURL
        txt += '\n\nPathSubst..ProgURL: %s' % u._ProgURL
        txt += '\n\nPathSubst..TempURL: %s' % u._TempURL
        txt += '\n\nPathSubst..UserURL: %s' % u._UserURL
        txt += '\n\nPathSubst..WorkURL: %s' % u._WorkURL
        MsgBox ( txt, title="%s.Session.PathSubstitution" %s.OfficeName )
        if s.OfficeName == _OPEN_OFFICE:
            msg += '\nPathSubst..UILangNbr %s' % u._UILangNbr
        msg += '\nPathSubst..UILangID %s' % u._UILangID
        msg += '\nPathSubst..UILocale %s' % u._UILocale
        if s.OfficeName == _LIBRE_OFFICE:
            msg += '\nPathSubst..UserName: %s' % u._UserName

    if _DIRS or _VARS:
        MsgBox(msg, title="%s._Paths" %s.OfficeName)

