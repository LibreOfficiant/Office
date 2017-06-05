# -*- coding: utf-8 -*-
from __future__ import unicode_literals
import uno, unohelper

from .ooSession import Session as ooSession  # OpenOffice base class

class Session(ooSession, object):
    """ Session: Represents the environment of the current script, 
    providing access to configuration information, information about the
    current user, and information about the 'LibreOffice' platform and
    release number.
    """
    def __init__(self):
        super().__init__()
    """@property
    def CurrentDatabase(self):
        pass"""
    @property
    def OfficeBuildVersion(self):
        pass
    @property  # e.g. '5.2.5.1'
    def OfficeRevision(self) -> str :  # aka ProductSetupVersionAboutBox
        return self.config._ProductSetupVersionAboutBox
    @property  # e.g. '.5.1'                            LIBREOFFICE ONLY
    def OfficeRevisionSuffix(self) -> str : 
        return self.config._ProductSetupExtension
    @property  # Current account name
    def UserName(self) -> str :  # as of LibreOffice v5.2 ONLY
        if self.OfficeVersion >= '5.2': return self._UserName
        else: raise AttributeError
    """
    From here, internal properties - DO NOT USE
    Ordered by L10N, Office, Product, PathXXX
    """
    @property  # UI language number e.g. France=33, USA=01
    def _UILangNbr(self): raise AttributeError  #        OPENOFFICE ONLY
    @property  # User account
    def _UserName(self) -> str : return self.path._UserName
    def getInvocationContext(self):
        return self.InvocationContext

"""                      BEGIN EXTENSION                             """
import unohelper
from com.sun.star.task import XJobExecutor
class _UnoComponentLibSession( Session, unohelper.Base, XJobExecutor):
    """ internal wrapper intended to distribute Session features to
    other languages e.g. Basic, Beanshell, JavaScript, etc..
    """
    def __init__(self, ctx):
        super().__init__()
        self.ctx = ctx
import unohelper
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation( \
    _UnoComponentLibSession, \
    "Libre.Office.Session", \
    ("arg.LibreOffice.Session",),)
"""                       END EXTENSION                              """

""" Change Log aka Summary of changes

Date        Full name           Description                                                                     Ref. #
2017-Apr-06 Alain H. Romedenne  Created inheritance from Open(Office)Session base class
2017-Mar-25 Alain H. Romedenne  Module creation      

"""


if __name__ == '__main__':
    """ Quick test cf. UnitTests for exhaustive tests.
    """

    import sys
    sys.exit(main(sys.argv))

