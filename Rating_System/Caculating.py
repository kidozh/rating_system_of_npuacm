#coding=utf-8
__author__ = 'kido'
from fileIO import *
from math import pow


class contest:
    res = []
    score = {}
    tot = []
    contestinfo = ''

    def __init__(self):
        xlrfile1 = xlrfile()
        self.current =  xlrfile1.curcontest()
        self.lastdata = xlrfile1.totalinfo()
        self.map = xlrfile1.map()

    def cmpa(self,e,rati,ratj):#e means math qi wang(pinyin)
        return 1/float((1+pow(10,(float(rati-ratj))/400)))

    def ntn(self,nickname):# trans nickname to name
        return  self.map[nickname]

    def getrating(self,name):
        return (self.lastdata[name])

    def rating(self):  # https://zh.wikipedia.org/wiki/%E7%AD%89%E7%BA%A7%E5%88%86
        for (nici,rati) in self.current.items():
            people = 0
            delta = 0
            for (nicj,ratj) in self.current.items():# caculate each people rating
                people +=1
                try :
                    nami = self.ntn((nici))
                except:
                    nami = 'anymous'
                try :
                    namj = self.ntn(nicj)
                except:
                    namj = 'anymous'


                try :
                    ratingi = self.getrating(nici)

                except:
                    #print '-----',ratingi,ratingj,nami,namj
                    #print '[Warning] Nickname : ',nici,' has been recognized as a guest!'
                    ratingi = 1500

                    pass
                try :
                    ratingj = self.getrating(nicj)


                except:
                    ratingj = 1500
                    #print '[Warning] Nickname : ',nicj,' has been recognized as a guest!'

                    pass
                rati =int(rati)
                ratj =int(ratj)

                #----caculate ratio
                if rati<ratj:
                    t = 1
                elif rati > ratj:
                    t = 0
                else :
                    t = 0.5

                if rati<ratj:
                    c = 1
                elif rati > ratj:
                    c = 0 # failed  give 0.2 point
                else :
                    c = 0.5

                if ratingi >=1700: # rating cut off
                    power = 1
                else :
                    power = 2
                delta +=(c - self.cmpa(t,int(ratingj),int(ratingi)))
                #print self.cmpa(t,int(ratingj),int(ratingi)),people
                print '^^^^^^^^^^^^^^',delta,nici,nami,ratingi,rati,nicj,namj,ratingj,ratj,c,delta
            #delta /= people-1
            self.score[nami] = delta*power
            print '###########',delta,nici,nami,ratingi,rati,c
            tmp = xlwtfile(nici,nami,ratingi,delta*power+people/10,rati)

            self.res.append(tmp)

    def allview(self):
        #print self.lastdata.items()
        for (nickname,rating) in self.lastdata.items():
            if self.current.has_key(nickname):
                for i in self.res:
                    if i.nickname == nickname:
                        self.tot.append(i)
                        break
            else:
                self.tot.append(xlwtfile(nickname,self.ntn(nickname),rating,0,0))


