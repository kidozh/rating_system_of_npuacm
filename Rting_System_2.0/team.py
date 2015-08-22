__author__ = 'Administrator'
from calc import *
from fileIO import teamdata,teamoutput

if __name__ =='__main__':
    data = teamdata()
    print data.timestring
    data.nicktorating()
    data.nicktoreal()
    data.nicktorank()
    output = teamoutput()
    output.summary()
    output.student()
