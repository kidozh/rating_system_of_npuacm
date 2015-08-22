#coding=utf-8
__author__ = 'Administrator'
import fileIO
import xlrd
import time
from math import pow
import sys
reload(sys)
sys.setdefaultencoding('utf-8')

class teamdata :
    singleinfo = fileIO.singledata()
    nicknametorealname = {}
    nicknametorating = {}
    nicknametorank = {}
    nicknametomember = {}

    def __init__(self):
        self.workbook = xlrd.open_workbook('Rating.xlsx')
        self.timestring = time.asctime()
        # get individual information
        self.singleinfo.nicktorating()
        self.singleinfo.nicktoreal()
        self.nicktoreal()
        self.nicktorank()

    def nicktoreal(self):
        table = self.workbook.sheet_by_name(u'team')
        nrows = table.nrows
        ncols = table.ncols
        dict = {}
        for i in range(nrows) :
            nickname = table.cell(i,0).value
            realname = table.cell(i,1).value
            dict[nickname] = realname
        self.nicknametorealname = dict

    def nicktorank(self):
        table = self.workbook.sheets()
        curindex =  len(table)-1
        table = self.workbook.sheets()[curindex]
        nrow = table.nrows
        dict = {}
        for i in range(nrow):
            if i==0:
                continue
            rank = table.cell(i,0).value
            rank = int(rank)
            nickname = table.cell(i,1).value
            dict[nickname] = rank
        self.nicknametorank = dict

    def nicktorating(self):
        table = self.workbook.sheet_by_name(u'team')
        nrows = table.nrows
        ncols = table.ncols
        dict = {}
        ratingdict = {}
        for i in range(nrows) :
            list = []
            nickname = table.cell(i,0).value
            rating = 0
            valid = 0
            for j in range(ncols):
                if j==0 or j==1 :
                    continue
                realname = table.cell(i,j).value
                try:
                    singlenickname = self.singleinfo.realtonickname[realname]
                except :
                    singlenickname = 'robot'
                rating += int(self.singleinfo.nicknametorating[singlenickname])
                if realname == '':
                    continue
                valid +=1
                list.append(realname)

            dict[nickname] = list
            ratingdict[nickname] = int(rating/valid)
        self.nicknametomember = dict # append list
        self.nicknametorating = ratingdict

class singledata : # copy from fileIO
    nicknametorealname = {}
    nicknametorating = {}
    nicknametorank = {}
    realtonickname = {}

    def __init__(self):
        self.workbook = xlrd.open_workbook('Rating.xlsx')
        self.timestring = time.asctime()
        #self.nicktoreal()
        #self.nicknametorating()

    def nicktoreal(self):
        table = self.workbook.sheet_by_name(u'student')
        nrows =  table.nrows
        dict ={}
        redict = {}
        dict['robot'] = '机器人'
        redict['机器人'] = 'robot'
        for i in range(nrows):
            nickname = str(table.cell(i,0).value)
            realname = str(table.cell(i,1).value)
            dict[nickname]=realname
            redict[realname]=nickname
        self.nicknametorealname = dict
        self.realtonickname = redict



    def nicktorating(self):
        table = self.workbook.sheet_by_name(u'total')
        ncols = table.ncols
        nrows = table.nrows
        dict = {}
        dict['robot'] = 1500
        for i in range(nrows):
            score = table.cell(i,ncols-1).value
            nickname = str(table.cell(i,0).value)
            dict[nickname]=score
        self.nicknametorating = dict

    def nicktorank(self):
        table = self.workbook.sheets()
        curtable = len(table)-1
        table = self.workbook.sheets()[curtable]
        nrows = table.nrows
        dict = {}
        for i in range(nrows):
            if i ==0:
                continue
            rank = int(table.cell(i,0).value)
            nickname = table.cell(i,1).value
            dict[nickname] = rank
        self.nicknametorank = dict


class single :
    nicknametorealname = {}
    nicknametorating = {}
    nicknametorank = {}
    nicknametoresult = {}


    def __init__(self):
        data = fileIO.singledata()
        data.nicktoreal()
        data.nicktorank()
        data.nicktorating()
        self.nicktorating = data.nicknametorating
        self.nicknametorating = data.nicknametorating
        self.nicknametorealname = data.nicknametorealname
        self.nicknametorank = data.nicknametorank
        #for (i,j) in self.nicknametorating


    def cmpa(self,rati,ratj): # key caculation
        return 1/float((1+pow(10,(float(rati-ratj))/400)))

    def rating(self):

        for (nicknamei,ranki) in self.nicknametorank.items() :
            delta = 0
            robot = False

            if self.nicknametorating.has_key(nicknamei) :
                namei = nicknamei
                ratingi = self.nicktorating[nicknamei]
            else :
                namei = 'Robot?'
                robot = True
                ratingi = 1500
            if ratingi >= 1700:
                power = 1
            else :
                power = 2

            for (nicknamej,rankj) in self.nicknametorank.items() :
                if self.nicknametorating.has_key(nicknamej) :
                    namej = nicknamej
                    ratingj = self.nicknametorating[nicknamej]
                else :
                    namej = 'Robot?'
                    ratingj = 1500

                if ranki > rankj:
                    score = 0
                elif ranki < rankj :
                    score = 1
                else :
                    score = 0.5

                delta +=(score - self.cmpa(ratingi,ratingj))*power
            #print '# Change On : ',nicknamei,' In ',ratingi+delta ,' Delta : ',delta
            self.nicknametoresult[nicknamei] = ratingi + delta


class teamcalc :
    nicknametorealname = {}
    nicknametorating = {}
    nicknametorank = {}
    nicknametoresult = {}
    singletoresult = {}
    nicknametomember ={}

    def __init__(self) :
        data = fileIO.teamdata()
        data.nicktoreal()
        data.nicktorank()
        data.nicktorating()
        self.nicktorating = data.nicknametorating
        self.nicknametorating = data.nicknametorating
        self.nicknametorealname = data.nicknametorealname
        self.nicknametorank = data.nicknametorank
        self.nicknametomember = data.nicknametomember
        #for (i,j) in self.nicknametorating


    def cmpa(self,rati,ratj): # key caculation
        return 1/float((1+pow(10,(float(rati-ratj))/400)))

    def rating(self):

        for (nicknamei,ranki) in self.nicknametorank.items() :
            delta = 0
            valid = 0
            if self.nicknametorating.has_key(nicknamei) :
                namei = nicknamei
                ratingi = self.nicknametorating[nicknamei]
            else :
                namei = 'Robot?'
                robot = True
                ratingi = 1500
            if ratingi >= 1700:
                power = 1
            else :
                power = 2

            for (nicknamej,rankj) in self.nicknametorank.items() :
                if self.nicknametorating.has_key(nicknamej) :
                    namej = nicknamej
                    ratingj = self.nicknametorating[nicknamej]
                else :
                    namej = 'Robot?'
                    ratingj = 1500

                if ranki > rankj:
                    score = 0
                elif ranki < rankj :
                    score = 1
                else :
                    score = 0.5

                delta +=(score - self.cmpa(ratingi,ratingj))*power

            self.nicknametoresult[nicknamei] = ratingi + delta
            for (nicknname,memberlist) in self.teamdata.nicknametomember.items() :
                for j in memberlist :

                    self.singletoresult[j] = delta/len(memberlist)





