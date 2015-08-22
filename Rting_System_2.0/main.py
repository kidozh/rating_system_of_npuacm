__author__ = 'Administrator'

from fileIO import singledata,singleoutput

if __name__ =='__main__':
    data = singledata()
    print data.timestring
    data.nicktoreal()
    data.nicktorank()
    data.nicktorating()
    output = singleoutput()
    print '!> Init Output OK '
    output.student()
    output.summary()