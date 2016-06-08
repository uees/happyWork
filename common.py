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


def exe_path():
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


class Objdict(dict):
    def __getattr__(self, name):
        if name in self:
            return self[name]
        else:
            raise AttributeError("No such attribute: %s" % name)

    def __setattr__(self, name, value):
        self[name] = value

    def __delattr__(self, name):
        if name in self:
            del self[name]
        else:
            raise AttributeError("No such attribute: %s" % name)


class ObjIterator(object):
    ''' 把一个字典迭代器的简单封装为对象字典迭代器
    '''
    def __init__(self, fetchObject):
        self.fetchObject = fetchObject
    
    def __iter__(self):
        for obj in self.fetchObject:
            yield Objdict(obj)
    
    def next(self):
        obj = next(self.fetchObject)
        if obj:
            obj = Objdict(obj)
        return obj
    
    @property
    def count(self):
        return len(self.fetchObject)
    
    def index(self, i):
        return Objdict(self.fetchObject[i])

   
class Category(object):
    def __init__(self, name, slug):
        self.name = name
        self.slug = slug
        
    def __str__(self):
        return self.name
    
    def __unicode__(self):
        return self.name