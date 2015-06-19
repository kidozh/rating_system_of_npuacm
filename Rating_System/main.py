__author__ = 'kido'
from Contest import *
from fileIO import *
from Caculating import *

a = contest()
a.rating()
a.allview()
#print a.score
a.res.sort(key=lambda xlwtfile: xlwtfile.rank)
a.tot.sort(key=lambda xlwtfile:xlwtfile.newrating,reverse=1)

for (i,j) in a.map.items():
    print i,j
    #print i
b = xlwtfile(1,2,3,4,5)
b.wrtxls(a.res)
b.wrtall(a.tot)


for i in a.tot:
    print i.nickname,i.name,i.oldrating,i.newrating
