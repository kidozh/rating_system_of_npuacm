__author__ = 'kido'
import xlrd
from xlwt import *
from Contest import *
import time

class xlrfile():
    xlsxfile =  xlrd.open_workbook('Rating.xlsx')
    timestring = ''
    valid = 0
    def __init__(self):
        self.xlsxfile = xlrd.open_workbook('Rating.xlsx')
        self.timestring = time.asctime( time.localtime(time.time()) )

    def curcontest(self):
        table = self.xlsxfile.sheets()
        curindex =  len(table)-1
        table = self.xlsxfile.sheets()[curindex]
        #print 'Your data now comes from '+table.by_name
        nrow = table.nrows
        dict = {}
        for i in range(nrow):
            if i==0:
                continue
            rank = table.cell(i,0).value
            rank = int(rank)
            nickname = table.cell(i,1).value
            #print rank,nickname
            dict[nickname] = rank
        valid = nrow+1
        return dict

    def map(self):
        table = self.xlsxfile.sheet_by_name(u'student')
        nrow = table.nrows
        dict = {}
        for i in range(nrow):
            nickname = table.cell(i,0).value
            realname = table.cell(i,1).value
            dict[nickname] = realname
        return dict

    def totalview(self):
        table = self.xlsxfile.sheet_by_name(u'total')
        return table


    def totalinfo(self):
        tmptable = self.totalview()
        nrow = tmptable.nrows
        ncol = tmptable.ncols
        dict = {'robot':'1500'}
        for i in range (nrow):
            try :
                oldrating = tmptable.cell(i,ncol-1).value
                name = tmptable.cell(i,0).value
            except BaseException:
                print 'Error : the excel file may be destroyed or total sheet does not exist '
            dict[name]=oldrating

        return dict


class xlwtfile():
    oldrating = 1500
    newrating = 1500
    delta = 0
    name = 'I am robot'
    nickname = 'robot'
    rank = 0

    def __init__(self):
        pass

    def __init__(self,nickname,name,oldrating,delta,rank):
        self.nickname = nickname
        self.name = name
        self.oldrating = oldrating
        self.delta = delta
        self.newrating = int(self.oldrating)+delta
        self.rank = rank

    def __str__(self):
        return str(self.rank)+'\t'+str(self.oldrating)+'\t'+str(self.newrating)+'\t'+str(self.delta)

    def __cmp__(self, other):
        return other.newrating>self.newrating

    def wrtxls(self,list):
        w = Workbook()
        ws = w.add_sheet('Result',cell_overwrite_ok=True)
        cnt = 0
        ws.write(cnt,0,'rank')
        ws.write(cnt,1,'nickname')
        ws.write(cnt,2,'name')
        ws.write(cnt,3,'newrating')
        ws.write(cnt,4,'oldrating')
        ws.write(cnt,5,'delta')
        cnt = 1

        for i in list :
            ws.write(cnt,0,i.rank)
            ws.write(cnt,1,i.nickname)
            ws.write(cnt,2,i.name)
            ws.write(cnt,3,i.newrating)
            ws.write(cnt,4,i.oldrating)
            ws.write(cnt,5,i.delta)
            cnt +=1
            print i.rank,i.nickname,i.name,i.newrating,i.oldrating,i.delta

        w.save('result.xls')

    def wrtall(self,list):
        w = Workbook()
        ws = w.add_sheet('Result',cell_overwrite_ok=True)
        cnt = 0
        ws.write(cnt,0,'rank')
        ws.write(cnt,1,'nickname')
        ws.write(cnt,2,'name')
        ws.write(cnt,3,'newrating')
        ws.write(cnt,4,'oldrating')
        ws.write(cnt,5,'delta')
        cnt = 1
        print '#### ',list

        for i in list :
            ws.write(cnt,0,cnt)
            ws.write(cnt,1,i.nickname)
            ws.write(cnt,2,i.name)
            ws.write(cnt,3,i.newrating)
            ws.write(cnt,4,i.oldrating)
            ws.write(cnt,5,i.delta)
            cnt +=1
            #print i.rank,i.nickname,i.name,i.newrating,i.oldrating,i.delta

        w.save('result1.xls')







