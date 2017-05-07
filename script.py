import os
from resource_helper import *
import argparse

parser = argparse.ArgumentParser(description='Syntax Example: python script.py 60 /Users/user/Desktop/matrix.xlsx')
parser.add_argument('yearGrowth', metavar='[yearGrowth]', type=int, help='Example: 60')
parser.add_argument('filename', metavar='[inputFilename]', help='Example: /Users/dacobos/Desktop/matrix.xlsx')
args = parser.parse_args()


yearGrowth = args.yearGrowth/100+1
kpiVal = 0.8


siteList = getSiteList(args.filename)
tenDemands = getTenDemands(args.filename)
hundDemands = getHundDemands(args.filename)
initialBw = getInitialbw(args.filename)



bw18 = []
newTen18 = []
utilTen18 = []
newHund18 = []
utilHund18 = []

bw19 = []
newTen19 = []
utilTen19 = []
newHund19 = []
utilHund19 = []

bw20 = []
newTen20 = []
utilTen20 = []
newHund20 = []
utilHund20 = []

bw21 = []
newTen21 = []
utilTen21 = []
newHund21 = []
utilHund21 = []

bw22 = []
newTen22 = []
utilTen22 = []
newHund22 = []
utilHund22 = []

bw23 = []
newTen23 = []
utilTen23 = []
newHund23 = []
utilHund23 = []

prevTen = []
prevHun = []
#
#
print 'Running script for Year Growth '+str(yearGrowth)+'%'+'and KPI 80%'


for i in range(len(siteList)):
    siteBw = initialBw[i]
    siteBw = round(siteBw*yearGrowth,2)
    bw18.append(siteBw)
    prevTen.append(tenDemands[i])
    prevHun.append(hundDemands[i])

    if siteBw >= (kpiVal*40):
        newTen18.append(0)
        utilTen18.append(0)
        newHund18.append(kpiHund(siteBw, prevHun[i]))
        prevHun[i] = prevHun[i] + newHund18[i]
        utilHund18.append(round(siteBw*100/(prevHun[i]*100),2))

    else:
        newTen18.append(kpiTen(siteBw, prevTen[i]))
        prevTen[i] = prevTen[i]+newTen18[i]
        utilTen18.append(round(siteBw * 100 /(prevTen[i]*10),2))
        newHund18.append(0)
        utilHund18.append(0)

    if bw18[i] == 0:
        print 'Error collecting initialBw for:'+siteList[i]




print 'Calculating for 2018'
print 'Calculating for 2019'
for i in range(len(siteList)):
    siteBw = round(bw18[i]*yearGrowth,2)
    bw19.append(siteBw)

    if siteBw >= (kpiVal*40):
        newTen19.append(0)
        utilTen19.append(0)
        newHund19.append(kpiHund(siteBw, prevHun[i]))
        prevHun[i] = prevHun[i] + newHund19[i]
        utilHund19.append(round(siteBw*100/(prevHun[i]*100),2))

    else:
        newTen19.append(kpiTen(siteBw, prevTen[i]))
        prevTen[i] = prevTen[i]+newTen19[i]
        utilTen19.append(round(siteBw * 100 /(prevTen[i]*10),2))
        newHund19.append(0)
        utilHund19.append(0)



    # print bw19[i]
    # print newTen19[i]
    # print utilTen19[i]
    # print newHund19[i]
    # print utilHund19[i]

print 'Calculating for 2020'
for i in range(len(siteList)):
    siteBw = round(bw19[i]*yearGrowth,2)
    bw20.append(siteBw)

    if siteBw >= (kpiVal*40):
        newTen20.append(0)
        utilTen20.append(0)
        newHund20.append(kpiHund(siteBw, prevHun[i]))
        prevHun[i] = prevHun[i] + newHund20[i]
        utilHund20.append(round(siteBw*100/(prevHun[i]*100),2))

    else:
        newTen20.append(kpiTen(siteBw, prevTen[i]))
        prevTen[i] = prevTen[i]+newTen20[i]
        utilTen20.append(round(siteBw * 100 /(prevTen[i]*10),2))
        newHund20.append(0)
        utilHund20.append(0)

    # print bw20[i]
    # print newTen20[i]
    # print utilTen20[i]
    # print newHund20[i]
    # print utilHund20[i]

print 'Calculating for 2021'
for i in range(len(siteList)):
    siteBw = round(bw20[i]*yearGrowth,2)
    bw21.append(siteBw)

    if siteBw >= (kpiVal*40):
        newTen21.append(0)
        utilTen21.append(0)
        newHund21.append(kpiHund(siteBw, prevHun[i]))
        prevHun[i] = prevHun[i] + newHund21[i]
        utilHund21.append(round(siteBw*100/(prevHun[i]*100),2))

    else:
        newTen21.append(kpiTen(siteBw, prevTen[i]))
        prevTen[i] = prevTen[i]+newTen21[i]
        utilTen21.append(round(siteBw * 100 /(prevTen[i]*10),2))
        newHund21.append(0)
        utilHund21.append(0)
    #
    # print bw21[i]
    # print newTen21[i]
    # print utilTen21[i]
    # print newHund21[i]
    # print utilHund21[i]

print 'Calculating for 2022'
for i in range(len(siteList)):
    siteBw = round(bw21[i]*yearGrowth,2)
    bw22.append(siteBw)

    if siteBw >= (kpiVal*40):
        newTen22.append(0)
        utilTen22.append(0)
        newHund22.append(kpiHund(siteBw, prevHun[i]))
        prevHun[i] = prevHun[i] + newHund22[i]
        utilHund22.append(round(siteBw*100/(prevHun[i]*100),2))

    else:
        newTen22.append(kpiTen(siteBw, prevTen[i]))
        prevTen[i] = prevTen[i]+newTen22[i]
        utilTen22.append(round(siteBw * 100 /(prevTen[i]*10),2))
        newHund22.append(0)
        utilHund22.append(0)
    #
    # print bw22[i]
    # print newTen22[i]
    # print utilTen22[i]
    # print newHund22[i]
    # print utilHund22[i]
    #

print 'Calculating for 2023'
for i in range(len(siteList)):
    siteBw = round(bw22[i]*yearGrowth,2)
    bw23.append(siteBw)

    if siteBw >= (kpiVal*40):
        newTen23.append(0)
        utilTen23.append(0)
        newHund23.append(kpiHund(siteBw, prevHun[i]))
        prevHun[i] = prevHun[i] + newHund23[i]
        utilHund23.append(round(siteBw*100/(prevHun[i]*100),2))

    else:
        newTen23.append(kpiTen(siteBw, prevTen[i]))
        prevTen[i] = prevTen[i]+newTen23[i]
        utilTen23.append(round(siteBw * 100 /(prevTen[i]*10),2))
        newHund23.append(0)
        utilHund23.append(0)

    # print bw23[i]
    # print newTen23[i]
    # print utilTen23[i]
    # print newHund23[i]
    # print utilHund23[i]


print 'Writing results to: '+args.filename
result = [bw18,newTen18,utilTen18,newHund18,utilHund18,bw19,newTen19,utilTen19,newHund19,utilHund19,bw20,newTen20,utilTen20,newHund20,utilHund20,bw21,newTen21,utilTen21,newHund21,utilHund21,bw22,newTen22,utilTen22,newHund22,utilHund22,bw23,newTen23,utilTen23,newHund23,utilHund23]
matrixWriter(result, args.filename)
