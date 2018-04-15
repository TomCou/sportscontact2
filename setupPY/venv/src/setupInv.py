
#!/usr/bin/python


import pip
pip.main(['install', 'schedule'])
pip.main(['install', 'gitpython'])
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
#import git


#rw_dir='C:/Users/developmentPC/Documents/dev/sportscontact/'
#repo = Repo(rw_dir)

def job(t):
    shutil.copy2('//SCONTACTSRV/Public/Commun/Export inv rpp/exportinvtxt.txt','C:/Users/developmentPC/Documents/dev/sportscontact/dbPre.json')
    shutil.copy2('//SCONTACTSRV/Public/Commun/Export inv rpp/exportinvtxt.txt','C:/Users/developmentPC/Documents/dev/sportscontact/exportinvtxt.txt')
    aList = [];
    finalDict = {'Items':aList};
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
                        art['cdp']= art['cdp'].rstrip()#cleanAtt(art['cdp'])
                        art['id']=art['ID']
                        del art['ID']
                        del art['price']
                        del art['car']
                        #print(ind,'- ', art['cdp'],'is', art['id'],'[ch#',art['qty_hb'],'][sn#',art['qty_sn'],']')
                        finalDict['Items'].append(art)
                        if(art['size'] is not ''):
                            checkSizeDict['Items'].append(art)
                        #if ind == 1000:
                         #   break
                #if ind == 1000:
                 #   break
            json.dump(finalDict,src_file)


    if(os.path.isfile('C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json')):
        os.remove('C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json')

    shutil.copy2('C:/Users/developmentPC/Documents/dev/sportscontact/dbW.json','C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/db.json')

    #with open('C:/Users/developmentPC/Documents/dev/sportscontact/db.json', 'r') as data_file:
     #   data = json.load(data_file)
    p = subprocess.Popen(r'start cmd /c C:/Users/developmentPC/Documents/dev/sportscontact2/sportscontact/cmdForPush.bat', shell=True)
    p.wait()
    print('Done: '+ str(datetime.now()))
    # repo.git.commit("commit time: "+time.localtime(secs))
    # origin = repo.remote(name='origin')
    # origin.push()


def cleanAtt(att):
    final = '';
    for x in len(att):
        y=x;
        while(att[y] is ' '):
            y=y+1
            if (x - y) is 2 :
                final=split(att,x)

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

