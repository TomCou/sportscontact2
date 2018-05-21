
#!/usr/bin/python

import ui
import pip
import opt
o=opt.O
#pip.main(['install', 'schedule'])
#pip.main(['install', 'gitpython'])
import os
import tkinter
from tkinter import *
from inspect import currentframe, getframeinfo
import shutil
import os
import glob
from datetime import datetime
import threading
import schedule
import time
from os import listdir
from os.path import isfile, join
import json
from pprint import pprint
from shutil import *
import subprocess
import csv
import SCLIB
import opt

#excelFolder='//SCONTACTSRV/Public/invCustom/'

#rw_dir='C:/Users/developmentPC/Documents/dev/sportscontact/'
#repo = Repo(rw_dir)

def makeInvFile(data,chORsn):  # DictResAttribute):
    # FILE CONTROL
    if (chORsn is "SN"):
        oldFile = o.scSoulierFileSN
    elif (chORsn is "HB"):
        oldFile = o.scSoulierFileCH
    tempFile = o.excelFolder
    tempFile = tempFile + oldFile
    rwh = ui.RWHANDLE(tempFile,mul=True)
    rwh.mulSheetWrite(data,chORsn,True)
   # indIter = 0
  #  maxIter = 2000

#    for x in data:
 #       indIter=indIter+1

        # if (chORsn is "SN"):
        #     if int(x['qty_sn']) > 0:
        #         if x['sdep']=='SOULIERS':
        #             if x['size'] != '5K' or '5.5K' or '6K' or '6.5K':
        #                 rwh.mulSheetWrite(x,chORsn,True)
        #             if indIter >= maxIter:
        #                 break
        #         else:
        #             pass
        #     else:
        #         pass
        # elif (chORsn is "HB"):
        #     if int(x['qty_hb']) > 0:
        #         if x['sdep'] == 'SOULIERS':
        #             if x['size'] != '5K' or '5.5K' or '6K' or '6.5K':
        #                 rwh.mulSheetWrite(x, chORsn, True)
        #             if indIter >= maxIter:
        #                 break
        #         else:
        #             pass
        #     else:
        #         pass

    rwh.saveFile()




    # fieldnames = ['cdp', 'size', 'id', 'qty_hb', 'qty_ns', 'price']
    # # readHeader = csv.DictReader(csvfile)#'.\\ui_template\\ui_ImportRessources.csv')
    # with open(tempFile, 'w') as csvfile2:
    #     writer = csv.DictWriter(csvfile2, fieldnames=fieldnames)
    #     writer.writeheader()
    #     for item in result:
    #         # writer = csv.DictWriter(csvfile2, fieldnames=rman.getDictKeys())
    #         if isfloat(item['size']):
    #             writer.writerow(item)
    #         elif isint(item['size']):
    #             writer.writerow(item)
    #         # {'first_name': 'Baked', 'last_name': 'Beans'})
    #         # writer.writerow({'first_name': 'Lovely', 'last_name': 'Spam'})
    #         # writer.writerow({'first_name': 'Wonderful', 'last_name': 'Spam'})

    # return tempFile
def isfloat(self, x):
    try:
        a = float(x)
    except ValueError:
        return False
    else:
        return True
def isint(self, x):
    try:
        a = float(x)
        b = int(a)
    except ValueError:
        return False
    else:
        return a == b

def job(t):

    shutil.copy2('//SCONTACTSRV/Public/Commun/Export inv rpp/exportinvtxt.txt','C:/Users/developmentPC/Documents/dev/sportscontact/dbPre.json')
    shutil.copy2('//SCONTACTSRV/Public/Commun/Export inv rpp/exportinvtxt.txt','C:/Users/developmentPC/Documents/dev/sportscontact/exportinvtxt.txt')
    aList = []
    finalDict = {'Items':aList}
    orgDict = {}

    femmeOptions=['N1000538','N1000642','N1000512','N1000629','N1000505','N1000500','02','1']
    hommeOptions=['N1000513','N1000630','N1000645','N1000501','01','N1000517','N1000518']
    juniorOptions=['N1000631','N1000540','N1000502','N1000544','N1000632','N1000644','N1000646']
    juniorGirlOptions=['N1000553','N1000507','03','2']
    adulteOptions=['N1000552','N1000539','N1000643','N1000506']

    indoorOptions = ['N2000696','N2000699']
    outdoorOptions = ['N2000698', 'N2000695', 'N2000692']
    turfOptions = ['N2000700','N2000697']

    femmeStringOptions = ['FEMME', 'WMN', 'WMN\'S', 'WMNS','WOMENS']
    hommeStringOptions = ['HOMME', 'MEN']
    juniorStringOptions = [' JR ', 'JUNIOR','YOUTH',' JR']
    juniorGirlStringOptions = [' GS JR ','GIRL','FILLE']
    adulteStringOptions = [' AD ','ADULTE','ADULTS','ADULT']

    outdoorStringOptions = [' FG ',' FXG ',' FG/AG ',' EXT ','OUTDOOR', ' MG ']
    indoorStringOptions = [' IN ',' IND ',' IC ', ' CT ', ' ID ','INDOOR']
    turfStringOptions = [' TF ',' TF',' CG ',' TT ',' TURF ']

    trapOption =['LASTIC','HELLO']
    checkSizeList = []
    checkSizeDict = {'Items':checkSizeList}


    if(os.path.isfile('C:/Users/developmentPC/Documents/dev/sportscontact/dbW.json')):
        os.remove('C:/Users/developmentPC/Documents/dev/sportscontact/dbW.json')
    #data = json.load(open('dbPre.json'))
    with open('C:/Users/developmentPC/Documents/dev/sportscontact/dbW.json', 'w') as src_file:
        with open('C:/Users/developmentPC/Documents/dev/sportscontact/dbPre.json', 'r') as data_file:
            data = json.load(data_file)
            for theList, product in data.items(): #got the single list of items
                ind = -1
                for art in data[theList]: #each item in the list is a dict
                    if((art['qty_hb'] > 0) or (art['qty_sn'] > 0) ):
                        ind = ind + 1
                        #print(art['cdp'])
                        art['cdp']= art['cdp'].strip(" ")
                        art['cdp']=art['cdp'].replace(' ','')#cleanAtt(art['cdp'])
                        art['niv1']=art['niv1'].strip(" ")
                        art['niv2'] = art['niv2'].strip(" ")
                        art['dep'] = art['dep'].strip(" ")
                        art['sdep'] = art['sdep'].strip(" ")

                        art['niv1'] = art['niv1'].replace(' ', '')  # cleanAtt(art['cdp'])
                        art['niv2'] = art['niv2'].replace(' ', '')  # cleanAtt(art['cdp'])
                        art['dep'] = art['dep'].replace(' ', '')  # cleanAtt(art['cdp'])
                        art['sdep'] = art['sdep'].replace(' ', '')  # cleanAtt(art['cdp'])

                        ## id VS. ID
                        art['id'] = art['ID']
                        del art['ID']

                        ## SOULIER VS. SOULIERS
                        if art['sdep']=='SOULIER':
                            art['sdep']='SOULIERS'

                        ## CHANGE EQUIVALENT WORDS AND CODES AND STUFF
                        # NIVEAU 1
                        for x in hommeOptions:
                            if art['niv1'] == x:
                                art['niv1'] = 'HOMME'
                                break
                        for x in femmeOptions:
                            if art['niv1'] == x:
                                art['niv1'] = 'FEMME'
                                break
                        for x in juniorGirlOptions:
                            if art['niv1'] == x:
                                art['niv1'] = 'GIRL'
                                break
                        for x in juniorOptions:
                            if art['niv1'] == x:
                                art['niv1'] = 'JUNIOR'
                                break
                        for x in adulteOptions:
                            if art['niv1'] == x:
                                art['niv1'] = 'HOMME'
                                break

                        # NIVEAU 2
                        for x in outdoorOptions:
                            if art['niv2'] == x:
                                art['niv2'] = 'OUTDOOR'
                                break
                        for x in indoorOptions:
                            if art['niv2'] == x:
                                art['niv2'] = 'INDOOR'
                                break
                        for x in turfOptions:
                            if art['niv2'] == x:
                                art['niv2'] = 'TURF'
                                break


                        # TRY TO FIND NIVEAU 1 AND 2 IN CAR IF NIV1 AND NIV2 ARE ''
                        if art['niv1'] is '':
                            for x in femmeStringOptions:
                                if (str(art['car']).find(x)) > -1:
                                    art['niv1'] = 'FEMME'
                                    break
                        if art['niv1'] is '':
                            for x in hommeStringOptions:
                                if (str(art['car']).find(x)) > -1:
                                    art['niv1'] = 'HOMME'
                                    break
                        if art['niv1'] is '':
                            for x in juniorGirlStringOptions:
                                if (str(art['car']).find(x)) > -1:
                                    art['niv1'] = 'GIRL'
                                    break
                        if art['niv1'] is '':
                            for x in juniorStringOptions:
                                if (str(art['car']).find(x)) > -1:
                                    art['niv1'] = 'JUNIOR'
                                    break
                        if art['niv1'] is '':
                            for x in adulteStringOptions:
                                if (str(art['car']).find(x)) > -1:
                                    art['niv1'] = 'HOMME'
                                    break

                        # NIVEAU 2
                        if art['niv2'] is '':
                            for x in outdoorStringOptions:
                                if (str(art['car']).find(x)) >-1:
                                    art['niv2'] = 'OUTDOOR'
                                    break
                        if art['niv2'] is '':
                            for x in indoorStringOptions:
                                if (str(art['car']).find(x)) >-1:
                                    art['niv2'] = 'INDOOR'
                                    break
                        if art['niv2'] is '':
                            for x in turfStringOptions:
                                if (str(art['car']).find(x)) >-1:
                                    art['niv2'] = 'TURF'
                                    break


                        try:
                            while(1):
                                if art['dep'] in orgDict:
                                    if art['sdep'] in orgDict[art['dep']]:
                                        # WE KNOW SDEP IS IN
                                        if art['niv1'] != '':
                                            # WE KNOW SDEP IS A DICT
                                            if art['niv1'] in orgDict[art['dep']][art['sdep']]:
                                                # WE KNOW NIV1 IS IN
                                                if art['niv2'] != '':
                                                    # WE KNOW NIV1 IS A DICT
                                                    if art['niv2'] in orgDict[art['dep']][art['sdep']][art['niv1']]:
                                                        orgDict[art['dep']][art['sdep']][art['niv1']][art['niv2']].append(art)
                                                        break
                                                    else:
                                                        orgDict[art['dep']][art['sdep']][art['niv1']][art['niv2']]=[]
                                                        orgDict[art['dep']][art['sdep']][art['niv1']][art['niv2']].append(art)
                                                        break
                                                else:
                                                    # WE KNOW NIV1 IS IN AS A LIST OR DICT, FIND TYPE, APPEND OBJECT, AND BREAK
                                                    if type(orgDict[art['dep']][art['sdep']][art['niv1']]) is dict:
                                                        # IF ITS A DICT THEN WE KNOW THERE IS OBJECTS WITH NIV2, SO ADD IT IN MISC
                                                        if 'MISC' in orgDict[art['dep']][art['sdep']][art['niv1']]:
                                                            orgDict[art['dep']][art['sdep']][art['niv1']]['MISC'].append(art)
                                                            break
                                                        else:
                                                            orgDict[art['dep']][art['sdep']][art['niv1']]['MISC']=[]
                                                            orgDict[art['dep']][art['sdep']][art['niv1']]['MISC'].append(art)
                                                            break
                                                    elif type(orgDict[art['dep']][art['sdep']][art['niv1']]) is list:
                                                        orgDict[art['dep']][art['sdep']][art['niv1']].append(art)
                                                        break
                                            else:
                                                # NIV1 IS NOT IN, DO WE ADD IT AS A LIST OR DICT? DEPENDS ON NIV2
                                                if art['niv2'] == '':  # ADD IT AS LIST, APPEND OBJECT, AND BREAK
                                                    # BUT MAYBE IN THE FUTURE THERE WILL BE THE SAME NIV1 WITH A NIV2 SO WE CAN'T ADD IT AS A LIST
                                                    orgDict[art['dep']][art['sdep']][art['niv1']] = {}
                                                    orgDict[art['dep']][art['sdep']][art['niv1']]['MISC'] = []
                                                    orgDict[art['dep']][art['sdep']][art['niv1']]['MISC'].append(art)
                                                    break
                                                else:# ADD IT AS DICT, AND CONTINUE
                                                    orgDict[art['dep']][art['sdep']][art['niv1']]={}

                                        else:
                                            # WE KNOW SDEP IS IN AS A DICT, BUT THERE IS NO NIV1 SO APPEND OBJECT TO MISC, AND BREAK
                                            if 'MISC' in orgDict[art['dep']][art['sdep']]:
                                                orgDict[art['dep']][art['sdep']]['MISC'].append(art)
                                                break
                                            else:
                                                orgDict[art['dep']][art['sdep']]['MISC'] = []
                                                orgDict[art['dep']][art['sdep']]['MISC'].append(art)
                                                break
                                    else:
                                        # SDEP IS NOT IN, WE ADD IT AS A DICT
                                        if art['niv1'] == '': # ADD IT AS DICT, APPEND OBJECT, AND BREAK
                                            orgDict[art['dep']][art['sdep']] = {}
                                            orgDict[art['dep']][art['sdep']]['MISC'] = []
                                            orgDict[art['dep']][art['sdep']]['MISC'].append(art)
                                            break
                                        else:# ADD IT AS DICT, AND CONTINUE
                                            orgDict[art['dep']][art['sdep']]={}
                                else:
                                    orgDict[art['dep']]={}
                        except:
                            print(art)

                        finalDict['Items'].append(art)

        # BASEBALL X 4 - 4 NIV1
        baseAD=[]
        baseAD.extend(orgDict['BASEBALL']['SOULIERS']['MISC'])
        baseAD.extend(orgDict['BASEBALL']['SOULIERS']['HOMME']['MISC'])
        baseAD.extend(orgDict['BASEBALL']['SOULIERS']['HOMME']['INDOOR'])
        baseAD.extend(orgDict['BASEBALL']['SOULIERS']['HOMME']['N2000784'])
        baseJR=[]
        baseJR.extend(orgDict['BASEBALL']['SOULIERS']['JUNIOR']['MISC'])
        baseJR.extend(orgDict['BASEBALL']['SOULIERS']['JUNIOR']['N2000786'])
        baseGR=[]
        baseGR.extend(orgDict['BASEBALL']['SOULIERS']['GIRL']['MISC'])
        baseWM=[]
        baseWM.extend(orgDict['BASEBALL']['SOULIERS']['FEMME']['MISC'])
        baseWM.extend(orgDict['SOFTBALL']['SOULIERS']['FEMME']['MISC'])
        # SOCCER X 7 - 3 NIV1 W/ 3 NIV2
        socADIN=[]
        socADIN.extend(orgDict['SOCCER']['SOULIERS']['HOMME']['INDOOR'])
        socADOU=[]
        socADOU.extend(orgDict['SOCCER']['SOULIERS']['HOMME']['OUTDOOR'])
        socADTF=[]
        socADTF.extend(orgDict['SOCCER']['SOULIERS']['HOMME']['TURF'])
        socJRIN = []
        socJRIN.extend(orgDict['SOCCER']['SOULIERS']['JUNIOR']['INDOOR'])
        socJROU = []
        socJROU.extend(orgDict['SOCCER']['SOULIERS']['JUNIOR']['OUTDOOR'])
        socJRTF = []
        socJRTF.extend(orgDict['SOCCER']['SOULIERS']['JUNIOR']['TURF'])
        socWMOU = []
        socWMOU.extend(orgDict['SOCCER']['SOULIERS']['FEMME']['OUTDOOR'])

        # FOOTBALL x 2
        footAD = []
        footAD.extend(orgDict['FOOTBALL']['SOULIERS']['MISC'])
        footAD.extend(orgDict['FOOTBALL']['SOULIERS']['HOMME']['MISC'])
        footAD.extend(orgDict['FOOTBALL']['SOULIERS']['HOMME']['N2000784'])
        footAD.extend(orgDict['FOOTBALL']['SOULIERS']['HOMME']['N2000782'])
        footAD.extend(orgDict['FOOTBALL']['SOULIERS']['HOMME']['N2000783'])
        footJR = []
        footJR.extend(orgDict['FOOTBALL']['SOULIERS']['JUNIOR']['MISC'])
        footJR.extend(orgDict['FOOTBALL']['SOULIERS']['JUNIOR']['N2000786'])
        footJR.extend(orgDict['FOOTBALL']['SOULIERS']['JUNIOR']['N2000787'])
        footJR.extend(orgDict['FOOTBALL']['SOULIERS']['JUNIOR']['N2000784'])
        # ENTRAINMENT x 3
        entrAD=[]
        entrWM = []
        entrGR = []
        entrAD.extend(orgDict['ENTRAINEMENT']['SOULIERS']['MISC'])
        entrAD.extend(orgDict['ENTRAINEMENT']['SOULIERS']['HOMME']['MISC'])
        entrGR.extend(orgDict['ENTRAINEMENT']['SOULIERS']['GIRL']['MISC'])
        entrWM.extend(orgDict['ENTRAINEMENT']['SOULIERS']['FEMME']['MISC'])
        # BASKETBALL x 4
        bballAD = []
        bballAD.extend(orgDict['BASKETBALL']['SOULIERS']['MISC'])
        bballAD.extend(orgDict['BASKETBALL']['SOULIERS']['HOMME']['MISC'])
        bballJR = []
        bballJR.extend(orgDict['BASKETBALL']['SOULIERS']['JUNIOR']['MISC'])
        bballGR = []
        bballGR.extend(orgDict['BASKETBALL']['SOULIERS']['GIRL']['MISC'])
        bballWM = []
        bballWM.extend(orgDict['BASKETBALL']['SOULIERS']['FEMME']['MISC'])
        # RUGBY x 2
        rugbyAD = []
        rugbyAD.extend(orgDict['RUGBY']['SOULIERS']['MISC'])
        rugbyAD.extend(orgDict['RUGBY']['SOULIERS']['HOMME']['MISC'])
        rugbyAD.extend(orgDict['RUGBY']['SOULIERS']['HOMME']['N2000831'])
        rugbyAD.extend(orgDict['RUGBY']['SOULIERS']['HOMME']['N2000830'])
        rugbyAD.extend(orgDict['RUGBY']['SOULIERS']['HOMME']['N2000832'])
        rugbyJR = []
        rugbyJR.extend(orgDict['RUGBY']['SOULIERS']['JUNIOR']['MISC'])
        # COURSE A PIED x 4
        courseAD = []
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['MISC'])
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['HOMME']['MISC'])
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['HOMME']['N2000507'])
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['HOMME']['N2000509'])
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['HOMME']['N2000505'])
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['HOMME']['N2000506'])
        courseAD.extend(orgDict['COURSEAPIED']['SOULIERS']['HOMME']['N2000508'])
        courseJR = []
        courseJR.extend(orgDict['COURSEAPIED']['SOULIERS']['JUNIOR']['MISC'])
        courseGR = []
        courseGR.extend(orgDict['COURSEAPIED']['SOULIERS']['GIRL']['MISC'])
        courseGR.extend(orgDict['COURSEAPIED']['SOULIERS']['GIRL']['N2000511'])
        courseWM = []
        courseWM.extend(orgDict['COURSEAPIED']['SOULIERS']['FEMME']['MISC'])
        courseWM.extend(orgDict['COURSEAPIED']['SOULIERS']['FEMME']['N2000501'])
        courseWM.extend(orgDict['COURSEAPIED']['SOULIERS']['FEMME']['N2000504'])
        courseWM.extend(orgDict['COURSEAPIED']['SOULIERS']['FEMME']['N2000506'])
        courseWM.extend(orgDict['COURSEAPIED']['SOULIERS']['FEMME']['N2000503'])
        courseWM.extend(orgDict['COURSEAPIED']['SOULIERS']['FEMME']['N2000502'])
        # COURTETVOLLEYBALL x 2
        volleyAD = []
        #volleyAD.extend(orgDict['COURTETVOLLEYBALL']['SOULIERS']['MISC'])
        #volleyAD.extend(orgDict['VOLLEYBALL']['SOULIERS']['MISC'])
        volleyAD.extend(orgDict['COURTETVOLLEYBALL']['SOULIERS']['HOMME']['MISC'])
        volleyAD.extend(orgDict['VOLLEYBALL']['SOULIERS']['HOMME']['MISC'])
        volleyAD.extend(orgDict['TENNIS']['SOULIERS']['HOMME']['MISC'])
        volleyWM = []
        volleyWM.extend(orgDict['COURTETVOLLEYBALL']['SOULIERS']['FEMME']['MISC'])
        volleyWM.extend(orgDict['VOLLEYBALL']['SOULIERS']['FEMME']['MISC'])
        # CHEER x 2
        cheerWM = []
        cheerWM.extend(orgDict['CHEERLEADING']['SOULIERS']['MISC'])
        cheerWM.extend(orgDict['CHEERLEADING']['SOULIERS']['FEMME']['MISC'])
        cheerGR = []
        cheerGR.extend(orgDict['CHEERLEADING']['SOULIERS']['GIRL']['MISC'])
        # SANDALES x 4
        sandAD = []
        sandAD.extend(orgDict['SANDALES']['SANDALES']['MISC'])
        sandAD.extend(orgDict['SANDALES']['SANDALES']['HOMME']['MISC'])
        sandJR = []
        sandJR.extend(orgDict['SANDALES']['SANDALES']['JUNIOR']['MISC'])
        sandWM = []
        sandWM.extend(orgDict['SANDALES']['SANDALES']['FEMME']['MISC'])

        finalDict['BASEBALLHOMME']=baseAD
        finalDict['BASEBALLFEMME']=baseWM
        finalDict['BASEBALLJUNIOR']=baseJR
        finalDict['BASEBALLFILLE']=baseGR

        finalDict['SOCCERHOMMEINDOOR']=socADIN
        finalDict['SOCCERHOMMEOUTDOOR']=socADOU
        finalDict['SOCCERHOMMETURF']=socADTF
        finalDict['SOCCERJUNIORINDOOR'] = socJRIN
        finalDict['SOCCERJUNIOROUTDOOR'] = socJROU
        finalDict['SOCCERJUNIORTURF'] = socJRTF
        finalDict['SOCCERFEMMEOUTDOOR'] = socWMOU

        finalDict['FOOTBALLHOMME']=footAD
        finalDict['FOOTBALLJUNIOR']=footJR

        finalDict['RUGBYHOMME'] = rugbyAD
        finalDict['RUGBYJUNIOR'] = rugbyJR

        finalDict['BASKETBALLHOMME'] = bballAD
        finalDict['BASKETBALLFEMME'] = bballWM
        finalDict['BASKETBALLJUNIOR'] = bballJR
        finalDict['BASKETBALLFILLE'] = bballGR

        finalDict['ENTRAINEMENTHOMME'] = entrAD
        finalDict['ENTRAINEMENTFEMME'] = entrWM
        finalDict['ENTRAINEMENTFILLE'] = entrGR

        finalDict['COURSEAPIEDHOMME'] = courseAD
        finalDict['COURSEAPIEDFEMME'] = courseWM
        finalDict['COURSEAPIEDJUNIOR'] = courseJR
        finalDict['COURSEAPIEDFILLE'] = courseGR

        finalDict['SANDALESHOMME'] = sandAD
        finalDict['SANDALESFEMME'] = sandWM
        finalDict['SANDALESJUNIOR'] = sandJR

        finalDict['VOLLEYBALLHOMME'] = volleyAD
        finalDict['VOLLEYBALLFEMME'] = volleyWM

        finalDict['CHEERFILLE'] = cheerGR
        finalDict['CHEERFEMME'] = cheerWM

        finalDictForFiles = finalDict
        json.dump(finalDict,src_file)
        del finalDictForFiles['Items']


    if(os.path.isfile('C:/Users/developmentPC/Documents/dev/sportscontact/db.json')):
        os.remove('C:/Users/developmentPC/Documents/dev/sportscontact/db.json')

    shutil.copy2('C:/Users/developmentPC/Documents/dev/sportscontact/dbW.json','C:/Users/developmentPC/Documents/dev/sportscontact/db.json')

    with open('C:/Users/developmentPC/Documents/dev/sportscontact/db.json', 'r') as data_file:
        data2 = json.load(data_file)

    p = subprocess.Popen(r'start cmd /c C:/Users/developmentPC/Documents/dev/sportscontact/cmdForPush.bat', shell=True)
    p.wait()
    print('Done App Update: '+ str(datetime.now()))

    try:
        makeInvFile(data=finalDictForFiles,chORsn="HB") #charlesbourg
        makeInvFile(data=finalDictForFiles,chORsn="SN") #saint-nicolas
        print('Done File Update: '+ str(datetime.now()))
    except Exception as e:
        print(e," in Job")



job(time.localtime(1))

schedule.every().day.at("09:05").do(job, 'It is 09:05')
schedule.every().day.at("10:05").do(job, 'It is 10:05')
schedule.every().day.at("11:05").do(job, 'It is 11:05')
schedule.every().day.at("12:05").do(job, 'It is 12:05')
schedule.every().day.at("13:05").do(job, 'It is 13:05')
schedule.every().day.at("14:05").do(job, 'It is 14:05')
schedule.every().day.at("15:05").do(job, 'It is 15:05')
schedule.every().day.at("16:05").do(job, 'It is 16:05')
schedule.every().day.at("17:05").do(job, 'It is 17:05')
schedule.every().day.at("18:05").do(job, 'It is 18:05')
schedule.every().day.at("19:05").do(job, 'It is 19:05')
schedule.every().day.at("20:05").do(job, 'It is 20:05')
schedule.every().day.at("21:05").do(job, 'It is 21:05')

while True:
    schedule.run_pending()
    time.sleep(60)  # wait one minute

