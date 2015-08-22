# coding=utf-8
__author__ = 'Administrator'
from xlwt import *
import xlrd
import time
import datetime
import sys

reload(sys)
sys.setdefaultencoding('utf-8')
from BeautifulSoup import BeautifulSoup
from calc import *


class singledata :
    nicknametorealname = {}
    nicknametorating = {}
    nicknametorank = {}
    realtonickname = {}

    def __init__(self):
        self.workbook = xlrd.open_workbook('Rating.xlsx')
        self.timestring = time.asctime()
        self.nicktoreal()
        #self.nicknametorating()
        #self.nicknametorealname()

    def nicktoreal(self):
        table = self.workbook.sheet_by_name(u'student')
        nrows = table.nrows
        dict = {}
        redict = {}
        dict['robot'] = '机器人'
        redict['机器人'] = 'robot'
        for i in range(nrows):
            nickname = str(table.cell(i, 0).value)
            realname = str(table.cell(i, 1).value)
            dict[nickname] = realname
            redict[realname] = nickname
        self.nicknametorealname = dict
        self.realtonickname = redict


    def nicktorating(self):
        table = self.workbook.sheet_by_name(u'total')
        ncols = table.ncols
        nrows = table.nrows
        dict = {}
        dict['robot'] = 1500
        for i in range(nrows):
            score = table.cell(i, ncols - 1).value
            nickname = str(table.cell(i, 0).value)
            dict[nickname] = score
        self.nicknametorating = dict

    def nicktorank(self):
        table = self.workbook.sheets()
        curtable = len(table) - 1
        table = self.workbook.sheets()[curtable]
        nrows = table.nrows
        dict = {}
        for i in range(nrows):
            if i == 0:
                continue
            rank = int(table.cell(i, 0).value)
            nickname = str(table.cell(i, 1).value)
            print '<fileIO LINE 68> Nickname : ', nickname
            dict[nickname] = rank
        self.nicknametorank = dict


class teaminfo:
    nickname = ''
    teamname = ''
    contestrank = 0
    oldrating = 1500
    newrating = 1500
    member = []

    def __init__(self, nick, teamname, memberlist, contestrank, oldrating, newrating):
        self.nickname = nick
        self.teamname = teamname
        self.contestrank = contestrank
        self.oldrating = oldrating
        self.newrating = newrating
        self.member = memberlist


class teamdata:
    singleinfo = singledata()
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
        for i in range(nrows):
            nickname = table.cell(i, 0).value
            realname = table.cell(i, 1).value
            dict[nickname] = realname
        self.nicknametorealname = dict

    def nicktorank(self):
        table = self.workbook.sheets()
        curindex = len(table) - 1
        table = self.workbook.sheets()[curindex]
        nrow = table.nrows
        dict = {}
        for i in range(nrow):
            if i == 0:
                continue
            rank = table.cell(i, 0).value
            rank = int(rank)
            nickname = table.cell(i, 1).value
            dict[nickname] = rank
        self.nicknametorank = dict

    def nicktorating(self):
        table = self.workbook.sheet_by_name(u'team')
        nrows = table.nrows
        ncols = table.ncols
        dict = {}
        ratingdict = {}
        for i in range(nrows):
            list = []
            nickname = table.cell(i, 0).value
            rating = 0
            valid = 0
            for j in range(ncols):
                if j == 0 or j == 1:
                    continue
                realname = table.cell(i, j).value
                try:
                    singlenickname = self.singleinfo.realtonickname[realname]
                except:
                    singlenickname = 'robot'
                rating += int(self.singleinfo.nicknametorating[singlenickname])
                if realname == '':
                    continue
                valid += 1
                list.append(realname)

            dict[nickname] = list
            ratingdict[nickname] = int(rating / valid)
        self.nicknametomember = dict  # append list
        self.nicknametorating = ratingdict


class generalinfo:
    nickname = ''
    realname = ''
    contestrank = 0
    oldrating = 1500
    newrating = 1500

    def __init__(self, nick, real, rank, oldrating, newrating):
        self.nickname = nick
        self.realname = real
        self.contestrank = rank
        self.oldrating = oldrating
        self.newrating = newrating


class singleoutput:
    nicktorealname = {}
    nicktorating = {}
    nicktorank = {}
    nicktoresult = {}
    infolist = []
    studentlist = []

    def __init__(self):
        result = single()
        self.map = result
        data = result
        data.rating()
        self.nicktorating = data.nicknametorating
        self.nicktorealname = data.nicknametorealname
        self.nicknametorank = data.nicknametorank
        self.nicktorank = data.nicknametorank
        self.nicktoresult = data.nicknametoresult
        self.makelist()
        self.totallist()


    def numcolor(self, num):
        if num >= 2600:
            return '<th class="r26p">' + str(num) + '</th>'
        elif num >= 2200:
            return '<th class="r22p">' + str(num) + '</th>'
        elif num >= 2000:
            return '<th class="r20p">' + str(num) + '</th>'
        elif num >= 1900:
            return '<th class="r19p">' + str(num) + '</th>'
        elif num >= 1700:
            return '<th class="r17p">' + str(num) + '</th>'
        elif num >= 1500:
            return '<th class="r15p">' + str(num) + '</th>'
        elif num >= 1300:
            return '<th class="r13p">' + str(num) + '</th>'
        elif num >= 1200:
            return '<th class="r12p">' + str(num) + '</th>'
        else:
            return '<th class="r12d">' + str(num) + '</th>'

    def posneg(self, num):
        if num >= 0:
            return '<th class="cp">' + str(num) + '</th>'
        else:
            return '<th class="cd">' + str(num) + '</th>'

    def rankcolor(self, rank):
        if rank <= 5:
            return '<th class="fst">' + str(rank) + '</th>'
        elif rank <= 21:
            return '<th class="sec">' + str(rank) + '</th>'
        else:
            return '<th class="trd">' + str(rank) + '</th>'


    def makelist(self):
        for nickname in self.nicktorank.keys():
            flag = False
            try:
                realname = self.nicktorealname[str(nickname)]
            except:
                realname = 'Robot?'
                #realname = self.map.nicknametorealname[nickname]
                print '[ FATAL ] Unknown Realname ', nickname
                flag = True
            rank = self.nicktorank[nickname]
            try:
                oldrating = self.nicktorating[nickname]
            except:
                oldrating = 1500
            newrating = self.nicktoresult[nickname]
            self.infolist.append(generalinfo(nickname, realname, rank, oldrating, newrating))
            pass

    def totallist(self):
        for nickname in self.nicktorealname.keys():
            flag = False
            try:
                realname = self.nicktorealname[nickname]
            except:
                realname = '无法识别的字符'
                print '[ FATAL ] Nickname :', nickname, ' Is Not Recognized. '
            rank = 0
            try:
                oldrating = self.nicktorating[nickname]
            except:
                print nickname, realname
                oldrating = 1500
            try:
                newrating = self.nicktoresult[nickname]
            except:
                newrating = oldrating
            self.studentlist.append(generalinfo(nickname, realname, rank, oldrating, newrating))


    def student(self):
        w = Workbook()
        self.studentlist.sort(key=lambda generalinfo: generalinfo.newrating, reverse=1)
        #---------------Excel------------
        ws = w.add_sheet('Result', cell_overwrite_ok=True)
        cnt = 0
        ws.write(cnt, 0, 'rank')
        ws.write(cnt, 1, 'nickname')
        ws.write(cnt, 2, 'name')
        ws.write(cnt, 3, 'newrating')
        ws.write(cnt, 4, 'oldrating')
        ws.write(cnt, 5, 'delta')
        cnt = 1

        for i in self.studentlist:
            ws.write(cnt, 0, cnt)
            ws.write(cnt, 1, str(i.nickname).decode('utf-8'))
            ws.write(cnt, 2, str(i.realname).decode('utf-8'))
            ws.write(cnt, 3, str(i.newrating).decode('utf-8'))
            ws.write(cnt, 4, str(i.oldrating).decode('utf-8'))
            ws.write(cnt, 5, str(i.newrating - i.oldrating).decode('utf-8'))
            cnt += 1
            #print i.rank,i.nickname,i.name,i.newrating,i.oldrating,i.delta

        w.save('result.xls')





        #---------------HTML-------------
        now = datetime.datetime.now()
        otherStyleTime = now.strftime("%Y-%m-%d")
        timestring = time.asctime()
        file = open('Student.html', 'w')
        body = ''
        head = '''
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html" charset = "utf-8">
	<title>Board|Contest Rating System</title>
	<link rel="stylesheet" href="board.css" type="text/css" />
</head>
<body>
    <div class="container">
'''
        tail = '<div class="footer"><p>Generated At : ' + timestring + '</p></div></body></html>'
        title = '<h1> NWPU Student Condition @ ' + otherStyleTime + '</h1>'
        tablehead = '''
        <table id="crsboard">
		<thead>
			<th> Rank </th>
			<th> Nickname </th>
			<th> Name </th>
			<th> New Rating </th>
			<th> Old Rating </th>
			<th> Change </th>
		</thead>'''
        cnt = 1
        for i in self.studentlist:
            cnt += 1
            if cnt % 2 == 0:
                body += '<tr class="row1">' + '\n'
                if 1:
                    body += self.rankcolor(cnt - 1)
                    body += '<th> ' + str(i.nickname) + ' </th>\n'
                    body += '<th> ' + (i.realname).decode('utf-8') + ' </th>\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
            if cnt % 2 == 1:
                body += '<tr class="row2">' + '\n'
                if 1:
                    body += self.rankcolor(cnt - 1)
                    body += '<th> ' + str(i.nickname) + ' </th>' + '\n'
                    body += '<th> ' + (i.realname) + ' </th>' + '\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
        htmlraw = head + title + tablehead + body + tail
        htmlcode = BeautifulSoup(htmlraw)
        file.write(htmlcode.prettify())


    def summary(self):
        self.infolist.sort(key=lambda generalinfo: generalinfo.contestrank)

        now = datetime.datetime.now()
        otherStyleTime = now.strftime("%Y-%m-%d")
        timestring = time.asctime()
        file = open('Summary@' + otherStyleTime + '.html', 'w')
        filecopy = open('Summary.html', 'w')
        body = ''
        head = '''
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html" charset = "utf-8">
	<title>Board|Contest Rating System</title>
	<link rel="stylesheet" href="board.css" type="text/css" />
</head>
<body>
    <div class="container">
'''
        tail = '<div class="footer"><p>Generated At : ' + timestring + '</p></div></body></html>'
        title = '<h1> NWPU ACM Contest @ ' + otherStyleTime + '</h1>'
        tablehead = '''
        <table id="crsboard">
		<thead>
			<th> Rank </th>
			<th> Nickname </th>
			<th> Name </th>
			<th> New Rating </th>
			<th> Old Rating </th>
			<th> Change </th>
		</thead>'''
        cnt = 1
        for i in self.infolist:
            cnt += 1
            if cnt % 2 == 0:
                body += '<tr class="row1">' + '\n'
                if 1:
                    body += self.rankcolor(i.contestrank)
                    body += '<th> ' + str(i.nickname) + ' </th>\n'
                    body += '<th> ' + (i.realname).decode('utf-8') + ' </th>\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
            if cnt % 2 == 1:
                body += '<tr class="row2">' + '\n'
                if 1:
                    body += self.rankcolor(i.contestrank)
                    body += '<th> ' + str(i.nickname) + ' </th>' + '\n'
                    body += '<th> ' + (i.realname) + ' </th>' + '\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
        htmlraw = head + title + tablehead + body + tail
        htmlcode = BeautifulSoup(htmlraw)
        file.write(htmlcode.prettify())
        filecopy.write(htmlcode.prettify())


class teaminfo:
    nickname = ''
    realname = ''
    contestrank = 0
    oldrating = 1500
    newrating = 1500
    memberlist = []

    def __init__(self, nick, real, list, rank, oldrating, newrating):
        self.nickname = nick
        self.realname = real
        self.memberlist = list
        self.contestrank = rank
        self.oldrating = oldrating
        self.newrating = newrating


class teamoutput:
    nicktorealname = {}
    nicktorating = {}
    nicktorank = {}
    nicktoresult = {}
    singleresult = {}
    infolist = []
    studentlist = []
    nicktomember = {}

    def __init__(self):
        result = teamcalc()
        data = result
        data.rating()
        self.nicktorating = data.nicknametorating
        self.nicktorealname = data.nicknametorealname
        self.nicknametorank = data.nicknametorank
        self.nicktorank = data.nicknametorank
        self.nicktoresult = data.nicknametoresult
        self.singleresult = data.singletoresult
        self.nicktomember = data.nicknametomember
        self.makelist()
        self.totallist()

    def numcolor(self, num):
        if num >= 2600:
            return '<th class="r26p">' + str(num) + '</th>'
        elif num >= 2200:
            return '<th class="r22p">' + str(num) + '</th>'
        elif num >= 2000:
            return '<th class="r20p">' + str(num) + '</th>'
        elif num >= 1900:
            return '<th class="r19p">' + str(num) + '</th>'
        elif num >= 1700:
            return '<th class="r17p">' + str(num) + '</th>'
        elif num >= 1500:
            return '<th class="r15p">' + str(num) + '</th>'
        elif num >= 1300:
            return '<th class="r13p">' + str(num) + '</th>'
        elif num >= 1200:
            return '<th class="r12p">' + str(num) + '</th>'
        else:
            return '<th class="r12d">' + str(num) + '</th>'

    def posneg(self, num):
        if num >= 0:
            return '<th class="cp">' + str(num) + '</th>'
        else:
            return '<th class="cd">' + str(num) + '</th>'

    def rankcolor(self, rank):
        if rank <= 5:
            return '<th class="fst">' + str(rank) + '</th>'
        elif rank <= 21:
            return '<th class="sec">' + str(rank) + '</th>'
        else:
            return '<th class="trd">' + str(rank) + '</th>'

    def makelist(self):
        for nickname in self.nicktorank.keys():  # current contest print
            flag = False
            try:
                realname = self.nicktorealname[nickname]
            except:
                realname = 'Robot?'
                flag = True
            rank = self.nicktorank[nickname]

            if flag == 0:
                oldrating = self.nicktorating[nickname]
            else:
                oldrating = 1500
            newrating = self.nicktoresult[nickname]
            member = self.nicktomember[nickname]
            self.infolist.append(teaminfo(nickname, realname, member, rank, oldrating, newrating))
            #self.infolist.append(generalinfo(nickname,realname,rank,oldrating,newrating))

    def totallist(self):
        for nickname in self.nicktorealname.keys():  # total student print
            flag = False
            try:
                realname = self.nicktorealname[nickname]
            except:
                realname = u'无法识别的字符'
            rank = 0
            try:
                oldrating = self.nicktorating[nickname]
            except:
                print nickname, realname
                oldrating = 1500
            try:
                #newrating = self.nicktoresult[nickname]
                newrating = self.singleresult[nickname]
            except:
                newrating = oldrating
            self.studentlist.append(generalinfo(nickname, realname, rank, oldrating, newrating))


    def student(self):
        w = Workbook()

        self.studentlist.sort(key=lambda generalinfo: generalinfo.newrating, reverse=1)
        #---------------Excel------------
        ws = w.add_sheet('Team', cell_overwrite_ok=True)
        cnt = 0
        ws.write(cnt, 0, 'rank')
        ws.write(cnt, 1, 'nickname')
        ws.write(cnt, 2, 'name')
        ws.write(cnt, 3, 'newrating')
        ws.write(cnt, 4, 'oldrating')
        ws.write(cnt, 5, 'delta')
        cnt = 1

        for i in self.studentlist:
            ws.write(cnt, 0, cnt)
            ws.write(cnt, 1, str(i.nickname).decode('utf-8'))
            ws.write(cnt, 2, str(i.realname).decode('utf-8'))
            ws.write(cnt, 3, str(i.newrating).decode('utf-8'))
            ws.write(cnt, 4, str(i.oldrating).decode('utf-8'))
            ws.write(cnt, 5, str(i.newrating - i.oldrating).decode('utf-8'))
            cnt += 1
            #print i.rank,i.nickname,i.name,i.newrating,i.oldrating,i.delta

        w.save('result.xls')





        #---------------HTML-------------
        now = datetime.datetime.now()
        otherStyleTime = now.strftime("%Y-%m-%d")
        timestring = time.asctime()
        file = open('Student.html', 'w')
        body = ''
        head = '''
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html" charset = "utf-8">
	<title>Board|Contest Rating System</title>
	<link rel="stylesheet" href="board.css" type="text/css" />
</head>
<body>
    <div class="container">
'''
        tail = '<div class="footer"><p>Generated At : ' + timestring + '</p></div></body></html>'
        title = '<h1> NWPU Student Condition @ ' + otherStyleTime + '</h1>'
        tablehead = '''
        <table id="crsboard">
		<thead>
			<th> Rank </th>
			<th> Nickname </th>
			<th> Name </th>
			<th> New Rating </th>
			<th> Old Rating </th>
			<th> Change </th>
		</thead>'''
        cnt = 1
        for i in self.studentlist:
            cnt += 1
            if cnt % 2 == 0:
                body += '<tr class="row1">' + '\n'
                if 1:
                    body += self.rankcolor(cnt - 1)
                    body += '<th> ' + str(i.nickname) + ' </th>\n'
                    body += '<th> ' + (i.realname).decode('utf-8') + ' </th>\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
            if cnt % 2 == 1:
                body += '<tr class="row2">' + '\n'
                if 1:
                    body += self.rankcolor(cnt - 1)
                    body += '<th> ' + str(i.nickname) + ' </th>' + '\n'
                    body += '<th> ' + (i.realname) + ' </th>' + '\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
        htmlraw = head + title + tablehead + body + tail
        htmlcode = BeautifulSoup(htmlraw)
        file.write(htmlcode.prettify())


    def summary(self):
        self.infolist.sort(key=lambda generalinfo: generalinfo.contestrank)

        now = datetime.datetime.now()
        otherStyleTime = now.strftime("%Y-%m-%d")
        timestring = time.asctime()
        file = open('Summary@' + otherStyleTime + '.html', 'w')
        filecopy = open('Summary.html', 'w')
        body = ''
        head = '''
<html>
<head>
	<meta http-equiv="Content-Type" content="text/html" charset = "utf-8">
	<title>Board|Contest Rating System</title>
	<link rel="stylesheet" href="board.css" type="text/css" />
</head>
<body>
    <div class="container">
'''
        tail = '<div class="footer"><p>Generated At : ' + timestring + '</p></div></body></html>'
        title = '<h1> NWPU Team Contest @ ' + otherStyleTime + '</h1>'
        tablehead = '''
        <table id="crsboard">
		<thead>
			<th> Rank </th>
			<th> Nickname </th>
			<th> TeamName </th>
			<th> Member </th>
			<th> New Rating </th>
			<th> Old Rating </th>
			<th> Change </th>
		</thead>'''
        cnt = 1
        for i in self.infolist:
            cnt += 1
            if cnt % 2 == 0:
                body += '<tr class="row1">' + '\n'
                if 1:
                    body += self.rankcolor(i.contestrank)
                    body += '<th> ' + str(i.nickname) + ' </th>\n'
                    body += '<th> ' + (i.realname).decode('utf-8') + ' </th>\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
            if cnt % 2 == 1:
                body += '<tr class="row2">' + '\n'
                if 1:
                    body += self.rankcolor(i.contestrank)
                    body += '<th> ' + str(i.nickname) + ' </th>' + '\n'
                    body += '<th> ' + (i.realname) + ' </th>' + '\n'
                    body += self.numcolor(int(i.newrating)) + '\n'
                    body += self.numcolor(int(i.oldrating)) + '\n'
                    body += self.posneg(int(i.newrating - i.oldrating)) + '\n'
                body += '</tr>' + '\n'
        htmlraw = head + title + tablehead + body + tail
        htmlcode = BeautifulSoup(htmlraw)
        file.write(htmlcode.prettify())
        filecopy.write(htmlcode.prettify())



