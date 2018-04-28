
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
#import git

#excelFolder='//SCONTACTSRV/Public/invCustom/'

#rw_dir='C:/Users/developmentPC/Documents/dev/sportscontact/'
#repo = Repo(rw_dir)

def makeInvFile(data,chORsn):  # DictResAttribute):

    # FILE CONTROL
    if (chORsn is "SN"):
        oldFile = o.scSoulierFileSN  # + str(datetime.now()) + '.csv'
    elif (chORsn is "CH"):
        oldFile = o.scSoulierFileCH  # + str(datetime.now()) + '.csv'
    tempFile = o.excelFolder
    tempFile = tempFile + oldFile
    rwh = ui.RWHANDLE(tempFile,mul=True)
    #headerFile = '.\\ui_template\\ui_ImportRessources'
    # rwh = RWHANDLE(headerFile, 0, 0, 0, 0, 0)  # ,'rb') as csvfile:
    # fieldnames = rwh.collectFromDB(typeOfData=o.listRow, rowOfKeys=0, rowOfValues=0)
    # print(listOfHeaders)
    for x in data['Items']:
        rwh.mulSheetWrite(x,chORsn)



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

def updateShoesInvFile(self, data2):
    print("yeah")

def job(t):


    shutil.copy2('//SCONTACTSRV/Public/Commun/Export inv rpp/exportinvtxt.txt','C:/Users/developmentPC/Documents/dev/sportscontact/dbPre.json')
    shutil.copy2('//SCONTACTSRV/Public/Commun/Export inv rpp/exportinvtxt.txt','C:/Users/developmentPC/Documents/dev/sportscontact/exportinvtxt.txt')
    aList = [];
    finalDict = {'Items':aList};

    orgDict = {}#'dep':{},'sdep':{},'niv1':{},'niv2':{}};

    checkSizeList = [];
    checkSizeDict = {'Items':checkSizeList};
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
                        art['id']=art['ID']
                        del art['ID']
                        art['niv1']=art['niv1'].strip(" ")
                        art['niv2'] = art['niv2'].strip(" ")
                        art['dep'] = art['dep'].strip(" ")
                        art['sdep'] = art['sdep'].strip(" ")

                        art['niv1'] = art['niv1'].replace(' ', '')  # cleanAtt(art['cdp'])
                        art['niv2'] = art['niv2'].replace(' ', '')  # cleanAtt(art['cdp'])
                        art['dep'] = art['dep'].replace(' ', '')  # cleanAtt(art['cdp'])
                        art['sdep'] = art['sdep'].replace(' ', '')  # cleanAtt(art['cdp'])

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

                        #print(ind,'- ', art['cdp'],'is', art['id'],'[ch#',art['qty_hb'],'][sn#',art['qty_sn'],']')
                        finalDict['Items'].append(art)
                        # SORT
                        #if art['sdep'] is "SOULIERS" or "SOULIER":

                        #orgDict[art['dep']][art['sdep']][art['niv1']][art['niv2']].append(art)
                        #checkSizeDict['Items'].append(art)

                        #if ind == 1000:
                         #   break
                #if ind == 1000:
                 #   break
            json.dump(finalDict,src_file)


    if(os.path.isfile('C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json')):
        os.remove('C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json')

    shutil.copy2('C:/Users/developmentPC/Documents/dev/sportscontact/dbW.json','C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json')

    with open('C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json', 'r') as data_file:
        data2 = json.load(data_file)



    #makeInvFile(data=data2,"CH") #charlesbourg
    #makeInvFile(data=data2,"SN") #saint-nicolas

    p = subprocess.Popen(r'start cmd /c C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/cmdForPush.bat', shell=True)
    p.wait()

    print('Done: '+ str(datetime.now()))
    # repo.git.commit("commit time: "+time.localtime(secs))
    # origin = repo.remote(name='origin')
    # origin.push()


job(time.localtime(1))
#with open('C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json', 'r') as data_file:
 #   data2 = json.load(data_file)
#makeInvFile(data=data2)

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

