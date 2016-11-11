# -*- coding: utf-8 -*-

import os
import sys

if 'posix' in sys.builtin_module_names:
    import readline

    def rlinput(prompt, prefill=''):
        readline.set_startup_hook(lambda: readline.insert_text(prefill))
        try:
            return input(prompt)
        finally:
            readline.set_startup_hook()
elif 'nt' in sys.builtin_module_names:
    import win32console
    _stdout = win32console.GetStdHandle(win32console.STD_OUTPUT_HANDLE)
    _stdin = win32console.GetStdHandle(win32console.STD_INPUT_HANDLE)

    def rlinput(prompt, prefill=''):
        keys = []
        for c in prefill:
            evt = win32console.PyINPUT_RECORDType(win32console.KEY_EVENT)
            evt.Char = c
            evt.RepeatCount = 1
            evt.KeyDown = True
            keys.append(evt)
        _stdin.WriteConsoleInput(keys)
        return input(prompt)


def get_home_path():
    return os.path.expanduser("~")


def get_desktop_path():
    return os.path.join(os.path.expanduser("~"), 'Desktop')


def module_path():
    """ This will get us the program's directory,
    even if we are frozen using py2exe"""
    if hasattr(sys, "frozen"):
        return os.path.dirname(sys.executable)
    return os.path.dirname(__file__)


def is_number(value):
    try:
        value + 1
    except TypeError:
        return False
    else:
        return True


def null2str(value):
    if value is None:
        value = ''
    return value
