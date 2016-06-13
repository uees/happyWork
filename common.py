#-*- coding: utf-8 -*-

import os
import sys
import winreg

def get_user_home_path():
    return os.path.expanduser("~")


def get_desktop_path():
    key = winreg.OpenKey(winreg.HKEY_CURRENT_USER,\
                          r'Software\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders',)
    return winreg.QueryValueEx(key, "Desktop")[0]


def we_are_frozen():
    """Returns whether we are frozen via py2exe.
    This will affect how we find out where we are located."""
    return hasattr(sys, "frozen")


def module_path():
    """ This will get us the program's directory,
    even if we are frozen using py2exe"""
    if we_are_frozen():
        return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)


def is_number(value):
    try:
        x = int(value)
    except:
        return False
    else:
        return True
