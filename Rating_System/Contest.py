#coding=utf-8
__author__ = 'kido'
from fileIO import *

class contestant:

    name = ''
    oldrating = 1500
    newrating = 1500
    id = 1

    def __init__(self,name,rating):
        self.name = name
        self.oldrating = rating

    def __init__(self):
        self.name = 'Robot'
        self.oldrating = 1500

    def strout(self):
        return self.name+'\t'+self.oldrating+'\t'+self.newrating

    def delta(self):
        return self.newrating-self.oldrating

    def input(self):
        dict = xlrfile.totalinfo()

