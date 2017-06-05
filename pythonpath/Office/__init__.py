# -*- coding: utf-8 -*-
from __future__ import unicode_literals

__all__ = ["model", "view", "controller"]

"""

<user>python/pythonpath/Office Python package.

This package contains 2 modules + 3 (Open/Libre)Office sub-packages:

bridge -- DO NOT import Bridge as it´s meant to be used 
          outside (Open/LibreOffice)
          Note: excluded in " from Office import * " statement.

devTools -- Utilities to assist Python development e.g. mri, xray,.. + 
            Open/LibreOffice Basic clones such as MsgBox, InputBox.
            Note: excluded in " from Office import * " statement.

<< MODEL                                                        MODEL >>
session -- Represents the environment of the current script, providing
           access to configuration information, information about the
           current user, and information about the *Office platform and
           release number.

<< VIEW                                                          VIEW >>
UiDesktop -- Represents the current *Office workplace windows. Holds
               various GUI facilities e.g. xxFilePicker.

<< CONTROLLER                                              CONTROLLER >>
ev<Classname> -- Represents graphical objects controllers, holding event
                 listeners and their respective actions.

eReferences: 
    https://docs.python.org/3/tutorial/modules.html#packages
    http://stackoverflow.com/questions/448271/what-is-init-py-for

"""

from sys import version_info
if version_info < (3,):
    from .model.ooSession import Session
else:
    from .model.libSession import Session
from .view.uiDesktop import UiDesktop 

# bridge is intended for Integrated Development Environments (IDEs)
# devTools is intended to assist testing ONLY

"""
     bridge   For IDEs only, not in production environments
   devTools   Not recommended in production environment 
    session       Model - BackOffice - LibreOffice back-end connectivity
uxWorkplace        View - FrontOffice front-end aka GUI
    ev<...>  Controller - Engines - Events, listeners, ... 

Caution:
-  'bridge' IS NOT meant to be used inside (Open/LibreOffice)
-  'devTools' IS NOT intended for production environments

"""
