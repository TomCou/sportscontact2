#!/usr/bin/python

import opt
o=opt.O
#opt.ImportCheck()

import os
import tkinter
import csv
from tkinter import *
#import ui_orderScanner
import openpyxl
from inspect import currentframe, getframeinfo
from xlutils.copy import copy # http://pypi.python.org/pypi/xlutils
from xlwt import easyxf # http://pypi.python.org/pypi/xlwt
import shutil
import os
import xlwt
from xlrd import open_workbook
from openpyxl import * #load_workbook
#import openpyxl
import glob
import datetime
import threading
import time
#import ui_browser
#import rMan
#import miMan
#import fpMan
from openpyxl import Workbook
from openpyxl.utils import coordinate_from_string, column_index_from_string
from openpyxl.utils import (get_column_letter)
from openpyxl.reader.excel import load_workbook, InvalidFileException
import xlrd
from os import listdir
from os.path import isfile, join


def debug(strInfo):
    if o.DEBUG:
        print(strInfo)

class DB(object):
    def __init__(self,myOrder,option,forWhichButton):
        self.myOrder=myOrder
        self.option=option
        self.forWhichButton=forWhichButton
    def updateDatabase(self,myOrder=None, option=None, forWhichButton=None):
        self.myOrder = myOrder
        self.option = option
        self.forWhichButton = forWhichButton
        return self.update()
    def update(self,arg1=None,arg2=None,arg3=None):
        def updateDatabaseThread(myOrder,option,forWhichButton): # if upload button update, myOrder = ID, and forWhichButton is the called button
            if option == o.fromKevinExcelSheet:
                rwh= RWHANDLE(filePath=o.orderDatabase, writingRow=None, po=forWhichButton, readingRow=o.fromKevinExcelSheet, idr=myOrder.id)
                rwh.po=myOrder.po
                listOfrMan=rwh.update(myOrder.order, o.fromKevinExcelSheet)
                rwh.saveFile()
                rwh2=RWHANDLE(filePath=o.itemDatabase, writingRow=None, po=forWhichButton)
                if rwh2.update(data=listOfrMan,fromWhere=o.fromWebCheckout)==True:
                    rwh2.saveFile()
                updateDatabaseThread(myOrder.po, o.uploadButton, o.fromKevinExcelSheet)
                return rwh.fpmList
            elif option == o.newOrder:
                rwh = RWHANDLE(filePath=myOrder.currentDatabase, readingRow=o.fromMainList, writingRow=None)
                rwh.po=myOrder.po
                if (rwh.update(myOrder.order, o.newOrder) == True):
                    rwh.saveFile()
                return rwh
            elif option == o.update:
                rwh = RWHANDLE(filePath=o.orderDatabase, readingRow=o.fromMainList, writingRow=None,po=forWhichButton)
                if (rwh.update(myOrder.order, o.update) == True):
                    rwh.saveFile()
                return rwh
            elif option == o.uploadButton:
                rwh = RWHANDLE(o.orderDatabase, readingRow=o.defaultRowID, writingRow=None,po=myOrder)
                rwh.wc= forWhichButton
                rwh.wr= myOrder
                try:
                    rwh.update('Y', o.uploadButton)
                except:
                    return False
                rwh.saveFile()
                return True
            elif option == o.fromWebCheckout:
                rwh = RWHANDLE(filePath=o.itemDatabase, readingRow=o.fromWebCheckout, po=forWhichButton)
                if (rwh.update(myOrder,o.fromWebCheckout) == True):
                    rwh.saveFile()
                return rwh
            elif option == o.deleteOrder:
                rwh = RWHANDLE(o.infoFile, 0, 0, idr=0, writingCol=0, readingCol=0)
                rwh.deleteOrder(myOrder)
                rwh.saveFile()
                rwh = RWHANDLE(o.orderDatabase, 0, 0, idr=0, writingCol=0, readingCol=0)
                rwh.deleteOrder(myOrder)
                rwh.saveFile()
                rwh = RWHANDLE(o.itemDatabase, 0, 0, idr=0, writingCol=0, readingCol=0,po=myOrder)
                rwh.deleteOrder(myOrder)
                rwh.saveFile()
        return updateDatabaseThread(self.myOrder, self.option, self.forWhichButton)

class RWHANDLE(object):
    def __init__(self, filePath, readingRow=0, writingRow=0, idr=None, writingCol=0, readingCol=0,po=0,mul=False):
        #self.newItemFile=False
        # if o.itemDatabase in filePath:
        #     for i in os.listdir(o.itemDatabase):
        #         if i.endswith(str(po)+".xlsx"):
        #             #workbook = load_workbook(o.itemDatabase+i, data_only=True)
        #             filePath=o.itemDatabase+"PO"+str(po)
        #             self.newItemFile=True
        #             break
        #     if self.newItemFile == False:
        #         filePath=o.itemDatabase+"PO"
        self.filePath = filePath
        self.rr = readingRow
        self.wr = writingRow
        self.wc = writingCol
        self.rc = readingCol
        self.fpmList = []
        self.po = po
        self.idr = idr
        self.mul = mul

        if(mul):
            puma="PUMA-DIADORA"
            nike="NIKE"
            joma="JOMA"
            adidas="ADIDAS"
            self.listDep = o.listOfDepartment
            self.listDepR = o.listOfDepartmentR
            self.dictDepNvR =o.dictNameOfDepNtoR
            self.dictDepRvN =o.dictNameOfDepRtoN

            try:
                self.wb = openpyxl.load_workbook(filePath + '.xlsx', data_only=True)  # ,formatting_info=True)
                # debug(str(e) + " I am in ui, __init__()")
                # self.wb = self.xls2xlsx(filePath + '.xls')
                self.rss = {}
                for nameR in self.listDepR:
#                    self.rss[name] = self.wb[dictDepNvW[name]]
                    #abc=self.dictDepRvN[nameR]
                    tmprs = self.wb[self.dictDepRvN[nameR]]#.get_sheet_by_name(self.dictDepRvN[nameR]) #self.wb[name]
                    tmprs.title = nameR
                    self.rss[nameR] =  tmprs #self.wb[name]


                #self.rss[puma] = self.wb.get_sheet_by_name(pumaW)
                #self.rss[adidas] = self.wb.get_sheet_by_name(adidasW)
                #self.rss[joma] = self.wb.get_sheet_by_name(jomaW)
                # self.rss[nike] = self.wb.get_sheet_by_name(nikeW)
                # self.rss[puma] = self.wb.get_sheet_by_name(pumaW)
                # self.rss[adidas] = self.wb.get_sheet_by_name(adidasW)
                #self.rss.get(joma).title(joma)
                #for file in self.rss:
                #dataList = [{'a': 1}, {'b': 3}, {'c': 5}]
                # for name in self.rss:
                #     for nSheet in self.rss[name]:
                #
                #         nSheet.title(name)
                self.wss={}
                for name in self.listDep:
                    tmpws=self.wb.copy_worksheet(self.rss[self.dictDepNvR[name]])
                    tmpws.title = name
                    self.wss[name] = tmpws #self.wb.copy_worksheet(self.rss[dictDepWvN[nameW]])

                # self.wss[nikeW] = self.wb.copy_worksheet(self.rss[nike])  # get_sheet(0)                          # the sheet to write to within the writable copy
                # self.wss[pumaW] = self.wb.copy_worksheet(self.rss[puma])
                # self.wss[adidasW] = self.wb.copy_worksheet(self.rss[adidas])  # get_sheet(0)                          # the sheet to write to within the writable copy
                # self.wss[jomaW] = self.wb.copy_worksheet(self.rss[joma])
                # for name in self.wss:
                #     for nSheet in self.rss[name]:
                #         self.wss[name][nSheet].title(name)
                self.rs = self.rss['socJR_tur_R']
                self.ws = self.wss['socJR_tur']

                #self.saveFile()


            except Exception as e:
                debug(str(e) + ", Location -- ui, RWHANDLE, __init__()")
            #     self.wb = self.xls2xlsx(filePath + '.xls')
            #     self.rs = self.wb.get_sheet_by_name('mainTracker')
            #     self.rs.title = 'mainTrackerRead'
            #     self.ws = self.wb.copy_worksheet(self.rs) # the sheet to write to within the writable copy
            #     self.ws.title = 'mainTracker'
            # if o.orderDatabase in filePath:
            #     self.wr = self.fetchRinC(target=po,colOfKey=0,index=0)
            #     if self.wr == -1:
            #         self.wr=self.countRows(0)
            #         debug('Writting row: '+str(self.wr))

    def mulSheetWrite(self,dict,chORsn,addIfNotFound):
        cdp= dict['cdp']
        size=dict['size']
        if(dict['dep']=='SOCCER'):
            cat=dict['dep']+'_'+dict['niv1']+'_'+dict['niv2']
        else:
            cat = dict['dep']+'_'+dict['niv1']

        if(chORsn is "SN"):
            totalQty = int(dict['qty_sn'])
        elif(chORsn is "HB"):
            totalQty = int(dict['qty_hb'])

        #totalQty=int(dict['qty_hb'])+int(dict['qty_sn'])
        try:
            for nSheet in self.rss:
                self.rs = self.rss[nSheet]
                w_nSheet=nSheet.split('_R')
                wSheetName=w_nSheet[0]
                a=str(self.getSingle(0,0))
                if(a == cat):
                    try:
                        while(1):
                            ind_r = self.fetchRinC(cdp,o.cdpCol,o.sizeStRow)
                            if(ind_r > 0):
                                ind_c = self.fetchCinR(size,o.sizeStCol,o.sizeStRow)
                                if(ind_c > 0):
                                    self.ws = self.wss[wSheetName]
                                    self.setSingle(ind_r,ind_c,int(totalQty))
                                    return True
                                    #break
                            elif(addIfNotFound):
                                self.insertRow(o.unbrandedItemAddingRow,1,wSheetName,cdp)
                                #tmprs = self.wb.get_sheet_by_name(self.dictDepRvN[nameR])  # self.wb[name]
                                self.wb.remove(self.wb[nSheet])#.get_sheet_by_name(nSheet))
                                tmpSh=self.wb.copy_worksheet(self.wss[wSheetName])
                                tmpSh.title = nSheet
                                self.rss[nSheet]=tmpSh
                                self.rs = self.rss[nSheet]

                    except Exception as e:
                        debug(str(e) + ", Location -- RWHANDLE, mulSheetWrite")
        except Exception as e:
            debug(str(e) + ", Location -- RWHANDLE, mulSheetWrite")

    def writeItems(self,filePath):
        ### Open source workbook ###

        workbook = open_workbook(filePath)
        sheet = workbook.sheet_by_index(0)

        ### Read data from source column ###

        # Initialise empty lists which will store values
        worktick = []
        ordernum = []
        siteadd = []
        locid = []
        leadlen = []

        def readsource(lst, colnum):
            source = []
            for i in range(sheet.max_row):
                lstval = sheet.cell_value(i, colnum)
                source += [lstval]
            source = source[1:]
            return source

        # Function for reading data from a selected column (colnum) to a predefined list (lst)
        worktick = readsource(worktick, 0)
        ordernum = readsource(ordernum, 1)
        siteadd = readsource(siteadd, 2)
        locid = readsource(locid, 3)
        leadlen = readsource(leadlen, 7)

        ### Open destination workbook ###

        # masterdoc = 'C:\where\is\the\file'
        masterdoc = 'Google docs Master Workbook.xlsx'
        destwb = load_workbook(masterdoc)
        ws = destwb.active

        # Find length of input row ## FOR APPENDING TO CURRENT VALUES
        rowx = 0
        cell = ws['A1']
        column_floor = []
        while cell.value != None:
            cell = ws.cell(0, 1 + rowx, 0)
            rowx += 1
            column_floor += [cell.value]
        column_floor = len(column_floor) + 1# column_floor is our start row value for writing source values

        theday = datetime.datetime.now()
        today = "%s-%s-%s" % (theday.day, theday.month, theday.year)
        added = 'Added on: ' + today

        ### Write source values ###
        def writeval(lst, colnum):
            titlecell = ws.cell(0, column_floor - 1, 0)
            titlecell.value = added
            for i, value in enumerate(lst):
                newcell = ws.cell(0, column_floor + i, colnum)
                newcell.value = value

        writeval(worktick, 0)
        writeval(ordernum, 1)
        writeval(siteadd, 2)
        writeval(locid, 3)
        writeval(leadlen, 4)

        ### Save file ###
        # Initialise variables
        current = os.getcwd()
        processed_path = current + os.sep + 'Processed' + os.sep
        # Save our edited Master Workbook
        destwb.save(masterdoc)
        # Save a copy of our edited Master Workbook, and move to a Backup dir
        backupPath = current + os.sep + 'Backup' + os.sep
        destwb.save(backupPath + 'Backup' + today + 'Master Doc.xlsx')
        # shutil.move('Backup Master Doc.xlsx', backupPath)
        # Move our source workbook to a Processed dir
        processed_path = current + os.sep + 'Processed' + os.sep
        processedtag = 'Processed on ' + today + ' '
        shutil.move(filestring, processed_path)
        # os.rename(processed_path + os.sep + filestring, processed_path + os.sep + filestring[:-4] + ' [' + processedtag + ']' + '.xls')

    def insertRow(self,indexStartRow, nRowsToAdd,currSheetTitle,itemToAdd):

        indexStartRow=indexStartRow+1
        prevSheet = self.wss[currSheetTitle]
        lastcol = prevSheet.max_column
        lastrow = prevSheet.max_row

        prevSheet.title = currSheetTitle+'_tmp'
        self.wb.create_sheet(index=0, title=currSheetTitle)
        newSheet = self.wb[currSheetTitle]#.get_sheet_by_name(currSheetTitle)
        tstSheet = self.wb[currSheetTitle+'_tmp']#.get_sheet_by_name(currSheetTitle+'_tmp')

        for row_num in range(1, indexStartRow):
            for col_num in range(1, lastcol + 1):
                newSheet.cell(row=row_num, column=col_num).value = prevSheet.cell(row=row_num, column=col_num).value
        for row_num in range(indexStartRow, lastrow + 1):
            offsetIndex = row_num + nRowsToAdd
            for col_num in range(1, lastcol + 1):
                newSheet.cell(row=offsetIndex, column=col_num).value = prevSheet.cell(row=row_num,column=col_num).value
        # lastcol = newSheet.max_column
        # lastrow = newSheet.max_row
        # print(lastrow, ",", lastcol)
        self.ws=self.wb[currSheetTitle]#.get_sheet_by_name(currSheetTitle)
        self.setSingle(indexStartRow-1,0,itemToAdd)
        self.wss[currSheetTitle]=self.wb[currSheetTitle]#.get_sheet_by_name(currSheetTitle)
        self.wb.remove(tstSheet)

    def deleteRows(self, indexStartRow, nRowsToRemove):
        indexStartRow = indexStartRow + 1
        prevSheet = self.ws
        lastcol = prevSheet.max_column
        lastrow = prevSheet.max_row
        lastrow2=self.countRows(0,None)
        if lastrow > (lastrow2+20):
            lastrow=lastrow2
        prevSheet.title = 'mainTrackerPrev'
        self.wb.create_sheet(index=0, title='mainTracker')
        newSheet = self.wb.get_sheet_by_name('mainTracker')

        for row_num in range(1, indexStartRow):
            for col_num in range(1, lastcol + 1):
                newSheet.cell(row=row_num, column=col_num).value = prevSheet.cell(row=row_num, column=col_num).value
        for row_num in range(indexStartRow+nRowsToRemove, lastrow + 1):
            negativeOffsetIndex = row_num - nRowsToRemove
            for col_num in range(1, lastcol + 1):
                newSheet.cell(row=negativeOffsetIndex, column=col_num).value = prevSheet.cell(row=row_num, column=col_num).value
        # lastcol = newSheet.max_column
        # lastrow = newSheet.max_row
        # print(lastrow, ",", lastcol)
        self.ws = self.wb.get_sheet_by_name('mainTracker')
        return
    def deleteOrder(self,POID):
        if self.filePath == o.infoFile:
            self.removeID(POID)

        elif self.filePath == o.orderDatabase:
            indexStart = self.fetchRinC(POID, 0,0)
            self.deleteRows(indexStart,1)

        elif self.filePath == o.itemDatabase:
            indexStart = self.fetchRinC(POID, 0,0)
            self.deleteRows(indexStartRow=indexStart,nRowsToRemove=self.countRows(colKey=0,startIndex=indexStart))
        else:
            return
    def xls2xlsx(self,filename):
        # first open using xlrd
        book = xlrd.open_workbook(filename)
        index = 0
        nrows, ncols = 0, 0
        while nrows * ncols == 0:
            sheet = book.sheet_by_index(index)
            nrows = sheet.nrows
            ncols = sheet.ncols
            index += 1
        # prepare a xlsx sheet
        book1 = Workbook()
        sheet1 = book1.get_active_sheet()
        sheet1.title='mainTracker'
        for row in range(1, nrows):
            for col in range(1, ncols):
                sheet1.cell(row=row, column=col).value = sheet.cell_value(row, col)
        return book1
    def removeID(self,PO):
        if PO == 1111:
            return
        else:
            prevIDs = []
            prevPOIDs = []
            # prevIDs.append(idr)
            # prevPOIDs.append(PO)
            for row in range(0, self.countRows(self.rc)):
                prevIDs.append(self.getSingle(row, self.rc))
                prevPOIDs.append(self.getSingle(row, self.rc + 1))
                # if (int(prevIDs[row + 1]) == int(idr)) or (int(prevPOIDs[row + 1]) == int(PO)):
                #     flagDuplicate = True
                #     dupIDs = 'rowID=' + str(idr) + ', POID=' + str(PO)
                #     print("can't insert new idr in the ui_curretnStatusInfo.xlsx because of duplicating the order idr.")
                #     print("The row idr value is =" + dupIDs)
                #     self.po = int(PO)
                #     self.idr = int(idr)
                #     return False
            for item in prevIDs:
                ind = prevIDs.index(item)
                if item != PO:
                    self.setSingle(ind, self.rc, item)
                    self.setSingle(ind, self.rc + 1, prevPOIDs[ind])
                else:
                    pass
            return
    def addNewID(self, idr, PO):
        flagDuplicate=False
        prevIDs=[]
        prevPOIDs=[]
        prevIDs.append(idr)
        prevPOIDs.append(PO)
        for row in range(0, self.countRows(self.rc)):
            prevIDs.append(self.getSingle(row, self.rc))
            prevPOIDs.append(self.getSingle(row, self.rc+1))
            if (int(prevPOIDs[row+1])==int(PO)):
                flagDuplicate=True
                dupIDs='rowID='+str(idr) + ', POID=' + str(PO)
                print("can't insert new idr in the ui_curretnStatusInfo.xlsx because of duplicating the order idr.")
                print("The row idr value is ="+dupIDs)
                self.po=int(PO)
                self.idr=int(idr)
                return False
        for item in prevIDs:
            ind=prevIDs.index(item)
            self.setSingle(ind, self.rc, item)
            self.setSingle(ind, self.rc+1, prevPOIDs[ind])
        self.saveFile()
        return True
    def countRows(self,colKey,startIndex=None):
        count=0
        if startIndex is None:
            for row in range(0,self.rs.max_row):
            #for col in rs.row_values(row):
                try:
                    if self.getCell(row,colKey).value:
                        count += 1
                    else:
                        break
                except:
                    try:
                        debug('there is nothing in cell {row:'+str(row)+', col:'+colKey+'}, LOCATION{ui,countRows()}')
                        break
                    except:
                        debug('Failed, ui, countRows()')
        else:
            refValue = self.getCell(startIndex,0)
            for row in range(startIndex, self.rs.max_row):
                try:
                    if self.getCell(row, colKey).value:
                        count += 1
                    else:
                        break
                except:
                    try:
                        debug('there is nothing in cell {row:' + str(row) + ', col:' + colKey + '}, LOCATION{ui,countRows()}')
                        break
                    except:
                        debug('Double Failed -- ui, countRows()')

        return count
    def makeResFile(self,listOfResMan,descForFile):#DictResAttribute):
        fileName='P0'+str(self.getPOID(self.idr))+'res.csv'
        tempFile=o.csvFolder
        newFile=(tempFile.split('.'))[1] +fileName
        tempFile=tempFile+fileName
        headerFile='.\\ui_template\\ui_ImportRessources'
        rwh=RWHANDLE(headerFile,0,0,0,0,0)#,'rb') as csvfile:
        fieldnames=rwh.collectFromDB(typeOfData=o.listRow,rowOfKeys=0,rowOfValues=0)
        #print(listOfHeaders)
        fieldnames=['resource-id', 'rtid', 'metaclass', 'checkout-center', 'barcode', 'description', 'circulating', 'condition-note', 'home-location-string', 'room-number', 'capacity', 'unserialized', 'containing-resource', 'accessories', 'purchase-date', 'purchase-order', 'vendor', 'asset-tag', 'manufacturer', 'model', 'serial-number', 'value', 'last-inventoried-date', 'replacement-date', 'hostname', 'hardware-version', 'software-version', 'ip-address', 'warranty-labor', 'warranty-parts', 'url', 'mac-address', 'biblio-id', 'call-number', 'restrictions', 'distributor', 'shelf-location']
        #readHeader = csv.DictReader(csvfile)#'.\\ui_template\\ui_ImportRessources.csv')
        with open(tempFile,'w') as csvfile2:
            writer = csv.DictWriter(csvfile2, fieldnames=fieldnames)
            writer.writeheader()
            for rman in listOfResMan:
                #writer = csv.DictWriter(csvfile2, fieldnames=rman.getDictKeys())
                writer.writerow(rman.getDictAliasWithAdjustedPrice())
                #{'first_name': 'Baked', 'last_name': 'Beans'})
                #writer.writerow({'first_name': 'Lovely', 'last_name': 'Spam'})
                #writer.writerow({'first_name': 'Wonderful', 'last_name': 'Spam'})

        return newFile
    def saveFile(self):
        current = os.getcwd()
        # a = self.wb.get_sheet_by_name("mainTrackerRead")
        # try:
        #     a1 = self.wb.get_sheet_by_name("mainTrackerRead1")
        #     self.wb.remove_sheet(a1)
        # except Exception as e:
        #     debug('No mainTrackerRead1'+ str(e))
        #     pass
        self.wb.save(self.filePath+'.xlsx')
        for nameR in self.listDepR:
            try:
                abc=self.wb.get_sheet_by_name(nameR)
                self.wb.remove(abc)
            except:
                self.wb.remove(nameR)
        self.wb.save(self.filePath + '.xlsx')

        # if o.itemDatabase in self.filePath:
        #     if self.newItemFile is False:
        #         self.filePath=self.filePath+str(self.po)
        self.wb.close()
        debug(self.filePath+ " has saved successfully.")
        # book = openpyxl.load_workbook(self.filePath + '.xlsx')
        # try:
        #     theday = datetime.datetime.now()
        #     today = "%s-%s-%s" % (theday.month, theday.day, theday.year)
        #     backupPath = current + os.sep + 'Backup' + os.sep
        #     sheet=book
        #     if o.DEBUG:
        #         if self.filePath == o.infoFile:
        #             sheet.save(backupPath + 'Backup' + today + 'currentStatusInfo.xlsx')
        #             sheet.save(o.infoFile+'_back.xlsx')
        #         elif self.filePath ==o.orderDatabase:
        #             sheet.save(backupPath + 'Backup' + today + 'orderTracker.xlsx')
        #             sheet.save(o.orderDatabase + '_back.xlsx')
        #         elif o.itemDatabase in self.filePath:
        #             sheet.save(backupPath + 'Backup' + '_PO' + str(self.po) + '.xlsx')
        #             sheet.save(o.itemDatabase + '_back.xlsx')
        #     else:
        #         if self.filePath == o.infoFile:
        #             sheet.save(backupPath + 'Backup' + today + 'currentStatusInfo.xlsx')
        #         elif self.filePath == o.orderDatabase:
        #             sheet.save(backupPath + 'Backup' + today + 'orderTracker.xlsx')
        #         elif o.itemDatabase in self.filePath:
        #             sheet.save(backupPath + 'Backup' +'_PO'+str(self.po)+'.xlsx')
        #     sheet.close()
        #     book.close()
        #     debug(self.filePath + "_back has saved successfully.")
        # except Exception as e:
        #     debug("Saving didnt work: " + str(e)+', Location ui.RWHADLE.saveFile()')
    def setSingle(self, rowKey=0, colKey=0,value=0):
        cr = str(get_column_letter(colKey + 1)) + str(rowKey + 1)
        try:
            self.ws[cr]=value
        except Exception as e:
            debug(str(e)+ ", Location ui.py, setSingle()")
    def setDictoRow(self, rowIndex, rowValue,data={}):
        lastCol = self.rs.max_column
        for col in range(0, lastCol):  # self.rs.max_column):
            currObj = self.getSingle(rowIndex, col)
            if currObj != "e":
                self.setSingle(rowKey=rowValue, colKey=col,value=data.get(currObj))
    def setListCol(self,startRow,colValue,data=[]):
        for item in range(0, len(data)):  # self.rs.max_column):
            self.setSingle(rowKey=startRow, colKey=colValue,label=data[item])
    def setListRow(self, rowValue, startCol, data=[]):
        for item in range(0, len(data)):  # self.rs.max_column):
            self.setSingle(rowKey=rowValue, colKey=startCol, label=data[item])
    def setRes(self,rman,rowValue):
        for col in range(0, self.rs.max_column):
            self.setSingle(rowValue,col,rman.getAtt(self.getSingle(rowKey=0,colKey=col)))
    def fetchRinC(self,target,colOfKey,index,safety=False):## recursive index starts at 0
        try:
            if index>o.maxShoesSheetRow: #=self.countRows(colOfKey,index):
                return -1
            elif str(target) != str(self.getSingle(index,colOfKey)):
                return self.fetchRinC(target,colOfKey,index+1)
            else:
                return index
        except:
            return -1
    def fetchCinR(self,target,index,rowOfKey,safety=False):## recursive index starts at 0
        try:
            if safety == True:
                if index>o.maxShoesSheetRow:
                    return -1

            if index>self.rs.max_column:
                return -1
            elif str(target) != str(self.getSingle(rowOfKey,index)):
                return self.fetchCinR(target,index+1,rowOfKey)
            else:
                return index
        except:
            return -1
    def fetchResPos(self,res,offset,rr,wr):
        item=res
        EOF=False
        IR=False
        while(1):
            rr +=1
            a=self.getSingle(rr, 0)

            if a is int or str:
                try:
                    if (item.getAtt('orderID') == int(self.getSingle(rr,0))):  # get the respective orderID row position
                        if (item.getAtt('groupN') == int(self.getSingle(rr,1))):  # get the respective groupN row position
                            if (item.getAtt('itemN') == int(self.getSingle(rr,2))):  # get the itemN
                                break
                            elif (rr + 1 >= self.rs.max_row):
                                EOF = True
                                break
                            elif (self.getSingle(rr + 1, 1) != item.getAtt('groupN')):
                                IR=True
                                break
                        elif (rr + 1 >= self.rs.max_row):
                            EOF = True
                            break
                        elif (self.getSingle(rr + 1, 0) != item.getAtt('orderID')):
                            IR = True
                            break
                    elif (rr + 1 >= self.rs.max_row):
                        EOF = True
                        break
                except:
                    try:
                        EOF = True
                        rr-=1
                        break
                    except:
                        print('aint working')
            else:
                EOF = True
                rr-=1
                break

        if IR:
            offset+=1
            wr+=1
            self.insertRow(wr, 1)
        if EOF:
            rr+=1
        return [offset,EOF,rr,wr]
    def setMatchCol(self, colOfKeys, colOfValues):
        target = self.idr
        self.rr = 0
        for row in range(0, int(self.getSingle(colOfKeys, 0))):  # col and row has symmetrical relationship
            a = int(float(self.getSingle(row, colOfKeys)))
            if a == target:
                return self.getSingle(row, colOfValues)
    def repInt(self,num):
        try:
            int(num)
            return True
        except:
            return False
    def update(self,data,fromWhere):
        flagItem = False
        listOfResMan = []
        listOfFpMan=[]
        if(fromWhere==o.fromKevinExcelSheet):
            data['PO#:'] = self.po
            possiblePODuplicateCheck=self.fetchRinC(self.po,0,0)
            if possiblePODuplicateCheck>-1:
                self.wr=possiblePODuplicateCheck
                self.idr=possiblePODuplicateCheck
            listOfKeys=['group_qty','group_manu','group_des','group_pri','group_datrec']
            a=self.rs.max_column
            print(self.wr)
            for col in range(0, a):
                currValue = self.getSingle(rowKey=fromWhere, colKey=col)
                if ('QTY:' in currValue):# make ressources
                    fpData=data.get('ITEMS:')
                    flagItem=True
                    qtyOfGroup=len(fpData)
                    for groupN in range(0, qtyOfGroup):
                        fpm = fpMan.fpMan(qty=fpData[groupN].get('group_qty'),orderID=self.po,groupN=groupN)
                        for att in listOfKeys:
                            fpm.setAtt(att,fpData[groupN].get(att))
                        for item in range(0,int(fpm.qty)):
                            rman = rMan.rMan()
                            rman.buildFromFpm(fpm)
                            rman.setID(self.po,groupN,item)
                            listOfResMan.append(rman)
                        listOfFpMan.append(fpm)
                elif "NOTES:" in currValue:
                   try:
                        a = data.get('NOTES:')
                        b='Notes:'
                        if len(a) >1:
                            for i in a:
                                d=a.get(i)
                                print (d)
                                b= b + str(d)
                        else:
                            b=a.get('0')
                        self.setSingle(rowKey=self.wr-1,colKey=col,value=b)
                   except TypeError as e:
                       debug(str(e) +", Location: ui, update() -fromKevin")
                elif "PO#" in currValue:
                    self.setSingle(rowKey=self.wr, colKey=col, value=int(data.get(currValue)))
                else:
                    if 'e' in currValue:
                        pass
                    else:
                        self.setSingle(rowKey=self.wr,colKey=col,value=data.get(currValue))
            if flagItem==True:
                self.fpmList=listOfFpMan
                return listOfResMan
            else:
                return 0
        elif(fromWhere == o.uploadButton):#
            self.setSingle(self.wr, self.wc,data)
            #upload
        elif(fromWhere==o.newOrder):
            try:
                for item in data:
                    for col in range(0, self.rs.max_column):  # self.rs.max_column):
                        currObj = self.rs.cell(self.rr, col).value
                        if currObj != "e":
                            self.ws.write(r=self.wr,c= col, label=data.get(currObj))
                            #data[currValue] = a  # self.rs.cell(ID, col).value
                        #elif currValue != "i":
                         #   data[currValue] = self.rs.cell(ID, col).value
            except:
                print("There is an issue in update(o.newOrder, etc...)")
                return False
            return True
        elif(fromWhere==o.update):

            for col in range(0, self.rs.max_column):  # self.rs.max_column):
                currObj = self.getSingle(self.rr, col)
                if currObj != "e":
                    if self.repInt(data.get(currObj)):
                        self.setSingle(self.wr,col,int(data.get(currObj)))
                    else:
                        self.setSingle(self.wr,col,data.get(currObj))

            return True
        elif(fromWhere==o.fromWebCheckout):
            offsetRow=0
            count=0
            EOF=False
            r = self.rs.max_row
            targetRowReader = 0
            targetRowWriter = 0
            if r < 3:
                EOF = True
                targetRowWriter = 0
            for item in data:# got a resource
                if EOF==False:
                    listOfProps=self.fetchResPos(item,offsetRow,targetRowReader,targetRowWriter)
                    offsetRow=listOfProps[0]
                    EOF=listOfProps[1]
                    targetRowReader=listOfProps[2]
                    targetRowWriter=listOfProps[2]
                else:
                    targetRowWriter+=1

                self.setSingle(rowKey=targetRowWriter,colKey=0, value=item.getAtt('orderID'))
                self.setSingle(rowKey=targetRowWriter, colKey=1, value=item.getAtt('groupN'))
                self.setSingle(rowKey=targetRowWriter, colKey=2, value=item.getAtt('itemN'))
                a=self.rs.max_column-1
                for col in range(0, a):  # self.rs.max_column):
                    col+=3
                    try:
                        currObj = self.getSingle(0, col)
                        self.setSingle(rowKey=targetRowWriter, colKey=col, value=item.getAtt(currObj))
                    except Exception as e:
                        print(str(e))
                        print("col: ",col)
                    if(col==a):
                        break
                    col-=3
                # except Exception as e:
                #     print(str(e))
                #     print("Updating Item Database Canceled, I am in ui.py , update()")
                #    return False
            return True
    def getPOID(self,orderRowID):
        rwh = RWHANDLE(filePath=o.infoFile, idr=orderRowID)
        return int(rwh.collectFromDB(typeOfData=o.matchColValue,colOfKeys=o.colOfOrderRowID, colOfValues=o.colOfPOID))
    def getCell(self,rowKey=0,colKey=0):
        cr = str(get_column_letter(colKey+1))+str(rowKey+1)
        return self.rs[cr]
    def getSingle(self,rowKey=0,colKey=0):
        try:
            cr = str(get_column_letter(colKey+1))+str(rowKey+1)
            # print('this is here'+str(cr))
            # print('rowKey: ' +str(rowKey)+' colKey: '+str(colKey))
            # print(self.rs[cr])
            return str(self.rs[cr].value)
        except Exception as e:
            debug(str(e)+', Location ui.getSingle()')
    def getDictoRow(self,rowIndex,rowValue):
        lastCol = self.rs.max_column
        data={}
        for col in range(0, lastCol):  # self.rs.max_column):
            currObj = self.getSingle(rowIndex, col)
            if currObj != "e":
                currValue = self.getSingle(rowValue, col)
                data[currObj] = currValue
        return data
    def getListCol(self,colValue,lastRow=0):
        lastRow=int(lastRow)
        if lastRow==0:
            lastRow = self.rs.max_row
        data = []
        for row in range(0, lastRow):  # self.rs.max_column):
            data.append(self.getSingle(row, colValue))

        return data
    def getListRow(self, rowIndex, lastCol=0):
        lastCol = int(lastCol)
        if lastCol == 0:
            lastCol = self.rs.max_column
        data = []
        for col in range(0, lastCol):  # self.rs.max_column):
            data.append(self.getSingle(rowIndex,col))
        return data
    def getListOfRes(self,targetID):
        while (True):
            listOfRes = []
            startIndex=self.fetchRinC(targetID,0,0)
            for rowOfIDs in range(startIndex, self.rs.max_row):
                if (str(self.getSingle(rowOfIDs, 0)) == str(targetID)):  # get the respective orderID row position
                    resMan = rMan.rMan()
                    for col in range(0, self.rs.max_column):
                        resMan.setAtt(self.getSingle(0, col), self.getSingle(rowOfIDs, col))
                    listOfRes.append(resMan)
            return listOfRes
    def getMatchCol(self,colOfKeys,colOfValues):
        target = self.idr
        self.rr=0
        for row in range(0, int(self.countRows(colOfKeys))):#col and row has symmetrical relationship
            a=int(float(self.getSingle(row, colOfKeys)))
            if  a== target:
                return self.getSingle(row, colOfValues)
    def setCollect(self, rowOfKeys=0, rowOfValues=0, colOfKeys=0, colOfValues=0, typeOfData='list'):
        self.rk=rowOfKeys
        self.rv=rowOfValues
        self.ck=colOfKeys
        self.cv=colOfValues
        self.tod=typeOfData
    def collect(self):
        return self.collectFromDB(self.rk,self.rv,self.ck,self.cv,self.tod)
    def collectFromDB(self, rowOfKeys=0, rowOfValues=0, colOfKeys=0, colOfValues=0, typeOfData='list'):   # gets the data in a general Dict
        if typeOfData==o.single:
                return self.getSingle(rowOfValues,colOfValues)
        elif typeOfData==o.listCol:
            return self.getListCol(colOfValues,self.countRows(colOfValues))
        elif typeOfData==o.listOfRes:
            return self.getListOfRes(self.po)
        elif typeOfData == o.listRow:
            return self.getListRow(rowIndex=0,lastCol=0)
        elif typeOfData==o.dicto:
            return self.getDictoRow(rowOfKeys,self.fetchRinC(rowOfValues,0,0))
        elif typeOfData==o.matchColValue:
            try:
                return self.getMatchCol(colOfKeys,colOfValues)
            except:
                print("The Row ID ",self.idr," (e.i self.idr), either doesn't have a PO ID associated to it, or the 'ui.collectFromDB' is getting a dubious error")
                print("returning defaultRowID rowID value (e.i. self.idr)")
                return self.idr

class ORDER(object):
    def __init__(self, fileToRead=None, debug=None, currentDatabase=None, getatt=None):#,readPdf):
        self.getRootAtt=getatt
        self.currentDatabase=currentDatabase
        self.debug = debug
        self.id = 0
        self.fromWho=0
        self.listOfFpMan=[]
        self.order = {}
        self.issue = ISSUE
        self.fileToRead = fileToRead
        self.debug=debug
        self.po=0

    def makeNewID(self, POID):
        rwh = RWHANDLE(o.infoFile, readingCol=o.colOfOrderRowID,po=POID)
        rowIndexForPODuplicate=rwh.fetchRinC(self.po,colOfKey=3,index=0)
        if rowIndexForPODuplicate > -1:
            newRowID=int(float(rwh.collectFromDB(rowOfKeys=rowIndexForPODuplicate,rowOfValues=rowIndexForPODuplicate,colOfKeys=2,colOfValues=2,typeOfData=o.single)))
        else:
            newRowID = 1+ int(float(rwh.collectFromDB(rowOfKeys=0,rowOfValues=0,colOfKeys=2,colOfValues=2,typeOfData=o.single)))
        self.id = newRowID
        if (rwh.addNewID(self.id,POID))==False:
            self.id=rwh.idr
            self.po=rwh.po
        return self.id

class popupWindowAttachFile(object):
    def __init__(self, master):
        top = self.top = Toplevel(master)
        self.l = Label(top, text="File to attach to footprint issue.")
        self.l2 = Label(top, text="Make sure to put the right file ext.")
        self.l.grid(row=0, column=0)  # pack()
        self.l2.grid(row=1, column=0)  # pack()
        self.e = Entry(top)
        self.e.grid(row=2, column=0)  # pack()
        self.e.bind('<Return>', self.getClean)
        self.e.bind('<Escape>', self.getEscape)
        self.b = Button(top, text='Upload', command=self.cleanup)
        self.b.grid(row=3, column=0)  # pack()
    def cleanup(self):
        self.value = self.e.get()
        self.top.destroy()
    def getClean(self, event):
        self.cleanup()
    def cancel(self):
        self.top.destroy()
    def getEscape(self, event):
        self.cancel()
class popupWindowQuery(object):
    def __init__(self, master,query=None):
        top = self.top = Toplevel(master)
        self.l = Label(top, text=query)
#        self.l2 = Label(top, text="**The order has been backed-up either way.")
        self.l.grid(row=0, column=0)  # pack()
        self.e = Entry(top)
        self.e.grid(row=2, column=0)  # pack()
        self.e.bind('<Return>', self.getClean)
        self.e.bind('<Escape>', self.getEscape)
        self.b = Button(top, text='Enter', command=self.confirm)
        self.b2 = Button(top, text='Cancel', command=self.cancel)
        self.b.grid(row=2, column=0)  # pack()
        self.b2.grid(row=2, column=1)  # pack()
        #self.top.bind('<Escape>', self.getEscape)
        #self.top.bind('<Enter>', self.getClean)
    def confirm(self):
        try:
            self.value = self.e.get()  # self.e.get()
        except:
            self.value = str(1111)
        if type(self.value) is not str:
            self.value = None
        self.top.destroy()
    def cancel(self):
        self.value = None  # self.e.get()
        self.top.destroy()
    def getEscape(self, event):
        self.cancel()
    def getClean(self, event):
        self.confirm()
class popupWindowConfirm(object):
    def __init__(self, master,POID=None):
        top = self.top = Toplevel(master)
        self.l = Label(top, text="Are you sure you want to delete the PO order "+str(POID)+"?")
        self.l2 = Label(top, text="**The order has been backed-up either way.")
        self.l.grid(row=0, column=0)  # pack()
        self.l2.grid(row=1, column=0)  # pack()
        #        self.e=Entry(top)
        #       self.e.pack()
        #      self.e.bind('<Return>',self.getClean)
        self.b = Button(top, text='Yes', command=self.cont)
        self.b2 = Button(top, text='No', command=self.cancel)
        self.b.grid(row=2, column=0)  # pack()
        self.b2.grid(row=2, column=1)  # pack()
        self.top.bind('<Escape>', self.getEscape)
    def cont(self):
        self.value = True  # self.e.get()
        self.top.destroy()
    def cancel(self):
        self.value = False  # self.e.get()
        self.top.destroy()
    def getEscape(self, event):
        self.cancel()
class popupWindowLocation(object):
    def __init__(self, master, preValue, listOfLocation):
        top = self.top = Toplevel(master)
        self.value=0
        self.a=0
        self.preValue = preValue
        self.locationSelection = StringVar()
        try:
            self.a=listOfLocation.index(self.preValue)
            if self.a>=1:
                self.locationSelection.set(listOfLocation[self.a])  # getSelectableLocations()[0])  # .set(data[0])initial value
        except:
            self.locationSelection.set(listOfLocation[0])
            pass
        #self.typeSelection.trace('w', self.option_select)
        option = OptionMenu(top, self.locationSelection, *listOfLocation)  # *choices)
        option.config(width=1)
        self.l = Label(top, text="Select a location.")
        self.l2 = Label(top, text="**If your location is not listed, input the location manually.")
        self.l.grid(row=0, column=0, sticky='we', columnspan=2)  # pack()
        self.l2.grid(row=2, column=0, sticky='we', columnspan=2)  # pack()
        self.e = Entry(top, textvariable=self.locationSelection)
        self.e.grid(row=1, column=0, sticky='we')  # pack()
        self.e.bind('<Return>', self.getClean)
        self.e.bind('<Escape>', self.getEscape)
        self.con = Button(top, text='Confirm', command=self.confirm)
        self.con.grid(row=4, column=0, columnspan=2, sticky='we')  # pack()
        self.can = Button(top, text='Cancel', command=self.cancel)
        self.can.grid(row=4, column=2, columnspan=2, sticky='we')  # pack()
        option.grid(row=1, column=1, columnspan=3, sticky='we')  # pack(side='left', padx=10, pady=10)
    def confirm(self):
    #     typeSelected = self.e.get()
    #     for i in range(0, len(self.listOfRT)):
    #         if typeSelected == self.listOfRT[i]:
    #             self.value = self.listOfRTID[i]
        self.value = self.e.get()
        self.top.destroy()
    def cancel(self):
        # self.typeSelection.set('')
        self.value = self.preValue  # self.e.get()
        self.top.destroy()
    def getEscape(self, event):
        self.cancel()
    def getClean(self, event):
        self.confirm()
class popupWindowRTID(object):
    def __init__(self, master, preValue, listOfRT=[],listOfRTID=[]):
        top = self.top = Toplevel(master)
        self.value = 0
        self.preValue=preValue
        self.listOfRT=listOfRT
        self.listOfRTID=listOfRTID
        # self.setAtt = setAtt
        self.typeSelection = StringVar()
        self.typeSelection.set(listOfRT[0])  # getSelectableLocations()[0])  # .set(data[0])initial value
        try:
            self.a=listOfRTID.index(self.preValue)
            if self.a>=1:
                self.typeSelection.set(listOfRT[self.a])  # getSelectableLocations()[0])  # .set(data[0])initial value
        except:
            self.typeSelection.set(listOfRT[0])  # getSelectableLocations()[0])  # .set(data[0])initial value
            pass
        # self.typeSelection.trace('w', self.option_select)
        option = OptionMenu(top, self.typeSelection, *listOfRT)  # *choices)
        option.config(width=1)
        self.l = Label(top, text="Ressource Type ID (rtid)")
        self.l2 = Label(top, text="**If your location is not listed, input the location manually.")
        self.l.grid(row=0, column=0, sticky='we', columnspan=2)  # pack()
        self.l2.grid(row=2, column=0, sticky='we', columnspan=2)  # pack()
        self.e = Entry(top, textvariable=self.typeSelection)
        self.e.grid(row=1, column=0, sticky='we')  # pack()
        self.e.bind('<Return>', self.getClean)
        self.e.bind('<Escape>', self.getEscape)
        self.con = Button(top, text='Confirm', command=self.confirm)
        self.con.grid(row=4, column=0, columnspan=2, sticky='we')  # pack()
        self.can = Button(top, text='Cancel', command=self.cancel)
        self.can.grid(row=4, column=2, columnspan=2, sticky='we')  # pack()
        option.grid(row=1, column=1, columnspan=3, sticky='we')  # pack(side='left', padx=10, pady=10)
    def confirm(self):
        typeSelected=self.e.get()
        for i in range(0,len(self.listOfRT)):
            if typeSelected in self.listOfRT[i]:
                self.value=self.listOfRTID[i]
        #self.value = self.e.get()
        self.top.destroy()
    def cancel(self):
        # self.typeSelection.set('')
        self.value = self.preValue  # self.e.get()
        self.top.destroy()
    def getEscape(self, event):
        self.cancel()
    def getClean(self, event):
        self.confirm()
class popupWindowGetFile(object):
    def __init__(self, master, preValue, listOfLocation):
        top = self.top = Toplevel(master)
        self.value=0
        self.a=0
        self.preValue = preValue
        self.locationSelection = StringVar()
        try:
            self.a=listOfLocation.index(self.preValue)
            if self.a>=1:
                self.locationSelection.set(listOfLocation[self.a])  # getSelectableLocations()[0])  # .set(data[0])initial value
        except:
            self.locationSelection.set(listOfLocation[0])
            pass
        #self.typeSelection.trace('w', self.option_select)
        option = OptionMenu(top, self.locationSelection, *listOfLocation)  # *choices)
        option.config(width=15)
        self.e = Entry(top, textvariable=self.locationSelection)
        self.l = Label(top, text="Select a file:", font=("Helvetica", 10))
        self.l2 = Label(top, bg='yellow',text="** IF your NEW ORDER(excel file) is not listed:", font=("Helvetica", 8),anchor=W)
        self.l3 = Label(top, bg='yellow',text="- Your file must be .xlsx -- Not .xls",font=("Helvetica", 8), anchor=W)
        self.l4 = Label(top, bg='yellow',text="- Your file must be located in the following folder:\n" + o.mspr_folder,font=("Helvetica", 8), anchor=W)
        self.con = Button(top, text='Confirm', command=self.confirm)
        self.can = Button(top, text='Cancel', command=self.cancel)
        self.l.grid(row=0, column=0,sticky=W, columnspan=2)  # pack()
        self.e.grid(row=1, column=0, sticky='we', columnspan=2)  # pack()
        option.grid(row=1, column=3, columnspan=10, sticky='we')  # pack(side='left', padx=10, pady=10)
        self.l2.grid(row=2, column=0,sticky=W,  columnspan=2)  # pack()
        self.l3.grid(row=3, column=0,sticky=W,  columnspan=2)  # pack()
        self.l4.grid(row=4, column=0,sticky=W,  columnspan=2)  # pack()
        self.con.grid(row=5, column=0, columnspan=1, sticky='we')  # pack()
        self.can.grid(row=5, column=2, columnspan=10, sticky='we')  # pack()
        self.e.bind('<Return>', self.getClean)
        self.e.bind('<Escape>', self.getEscape)
    def confirm(self):
        self.value = self.e.get()
        self.top.destroy()
    def cancel(self):
        self.value = self.preValue
        self.top.destroy()
    def getClean(self, event):
        self.confirm()
    def getEscape(self, event):
        self.cancel()
class popupWindowDate(object):
    def __init__(self, master, str_group_mth, str_group_day, str_group_year, frame_look={}, **look):
        args = dict(relief=tkinter.SUNKEN, border=1)
        args.update(frame_look)
        top = self.top = Toplevel(master, **args)
        args = {'relief': tkinter.FLAT}
        args.update(look)
        self.str_group_mth = str_group_mth
        self.str_group_day = str_group_day
        self.str_group_year = str_group_year
        self.label_mth = tkinter.Label(top, text='mm', **args)
        self.label_day = tkinter.Label(top, text='dd', **args)
        self.label_year = tkinter.Label(top, text='yyyy', **args)
        self.entry_1 = tkinter.Entry(top, width=2, **args)
        self.label_1 = tkinter.Label(top, text='/', **args)
        self.entry_2 = tkinter.Entry(top, width=2, **args)
        self.label_2 = tkinter.Label(top, text='/', **args)
        self.entry_3 = tkinter.Entry(top, width=4, **args)
        self.label_mth.grid(row=0, column=0)
        self.label_day.grid(row=0, column=2)
        self.label_year.grid(row=0, column=4)
        self.entry_1.grid(row=1, column=0)
        self.label_1.grid(row=1, column=1)
        self.entry_2.grid(row=1, column=2)
        self.label_2.grid(row=1, column=3)
        self.entry_3.grid(row=1, column=4)
        self.entry_1.bind('<KeyRelease>', self._e1_check)
        self.entry_2.bind('<KeyRelease>', self._e2_check)
        self.entry_3.bind('<KeyRelease>', self._e3_check)
        self.entry_3.bind('<Return>', self.getClean)
        self.e.bind('<Escape>', self.getEscape)
        self.con = Button(top, text='Confirm', command=self.confirm, **args)
        self.con.grid(row=2, column=0, columnspan=2)  # grid(row=4, column=0, columnspan=2, sticky='we')  # pack()
        self.can = Button(top, text='Cancel', command=self.cancel, **args)
        self.can.grid(row=2, column=3, columnspan=2)  # .grid(row=4, column=2, columnspan=2, sticky='we')  # pack()
    def _backspace(self, entry):
        cont = entry.get()
        entry.delete(0, tkinter.END)
        entry.insert(0, cont[:-1])
    def _e1_check(self, e):
        cont = self.entry_1.get()
        if len(cont) >= 2:
            self.entry_2.focus()
        if len(cont) > 2 or not cont[-1].isdigit():
            self._backspace(self.entry_1)
            self.entry_1.focus()
    def _e2_check(self, e):
        cont = self.entry_2.get()
        if len(cont) >= 2:
            self.entry_3.focus()
        if len(cont) > 2 or not cont[-1].isdigit():
            self._backspace(self.entry_2)
            self.entry_2.focus()
    def _e3_check(self, e):
        cont = self.entry_2.get()
        if len(cont) > 4 or not cont[-1].isdigit():
            self._backspace(self.entry_3)
    def confirm(self):
        # self.str_group_mth.set(self.entry_1.get())
        # self.str_group_day.set(self.entry_2.get())
        # self.str_group_year.set(self.entry_3.get())
        self.value = [self.entry_1.get(), self.entry_2.get(), self.entry_3.get()]
        self.top.destroy()
    def cancel(self):
        self.value = [self.entry_1.get(), self.entry_2.get(), self.entry_3.get()]
        self.top.destroy()
    def getClean(self, event):
        self.confirm()
    def getEscape(self, event):
        self.cancel()
def uploadingToFP(threadName, issue,updateGUI):
    fpBr = ui_browser.footprint_driver()
    fpBr.processOrder_fp(issue=issue)
    debug("upload to fp done.")
def uploadingToWC(threadName,curRowID,updateGUI,*fileNames):
    try:
        wcBr = ui_browser.webcheckout_driver()
        wcBr.processOrder_wc(fileNames)
        debug("upload to wc done.")
    except Exception as e:
        debug(str(e)+', --ui.py,  uploadingToWC()')
def uploadingToBC(threadName,curRowID,updateGUI,*fileNames):
    try:
        bcBr = ui_browser.basecamp_driver()
        bcBr.processOrder_bc(fileNames)
    except:
        print('Basecamp, trouble in paradise, in ui.py - uploadingToBC()')

class popup:

    def __init__(self):
        self.RTID
        self.location
    #  Popup for choosing project in basecamp
    # def projectOptions(self, root,setAtt, strVar):
    #     #onlyfiles = [f for f in listdir(o.mspr_folder) if isfile(join(o.mspr_folder, f))]
    #     listOfFiles=[]
    #     for file in os.listdir(o.mspr_folder):
    #         if file.endswith(".xlsx"):
    #             listOfFiles.append(str(file))
    #     debug(listOfFiles)
    #     self.winLoc = popupWindowGetFile(root, strVar.get(), listOfFiles)
    #     root.wait_window(self.winLoc.top)
    #     strVar.set(self.winLoc.value)
    def RTID(self, root, curValue, strVar):
        self.winRT=popupWindowRTID(root, curValue, self.getSelectableRTIDs(o.colRTList), self.getSelectableRTIDs(o.colRTIDList))
        root.wait_window(self.winRT.top)
        strVar.set(self.winRT.value)
        #setatt(strVar,self.winRT.value)
    def fileOptions(self, root,setAtt, strVar):
        #onlyfiles = [f for f in listdir(o.mspr_folder) if isfile(join(o.mspr_folder, f))]
        listOfFiles=[]
        for file in os.listdir(o.mspr_folder):
            if file.endswith(".xlsx"):
                listOfFiles.append(str(file))
        debug(listOfFiles)
        self.winLoc = popupWindowGetFile(root, strVar.get(), listOfFiles)
        root.wait_window(self.winLoc.top)
        strVar.set(self.winLoc.value)
    def allFileOptions(self, root, listOfFilesToAttach,strVar_filesToAttach):
        listOfFiles = [f for f in listdir(o.mspr_folder) if isfile(join(o.mspr_folder, f))]
        # for file in os.listdir(o.mspr_folder):
        #     if file.endswith(".xlsx"):
        #         listOfFiles.append(str(file))
        debug(listOfFiles)
        self.winLoc = popupWindowGetFile(root, '', listOfFiles)
        root.wait_window(self.winLoc.top)
        #strVar.set(self.winLoc.value)
        listOfFilesToAttach.append(self.winLoc.value)
        strVar_filesToAttach.set('')
        allFiles = ''
        for i in listOfFilesToAttach:
            allFiles = str(i) + '\n' + allFiles
        strVar_filesToAttach.set(allFiles)
    def location(self, root, strVar):
        self.winLoc=popupWindowLocation(root,strVar,self.getSelectableLocations())
        root.wait_window(self.winLoc.top)
        strVar.set(self.winLoc.value)
        #setatt(strVar, self.winLoc.value)
    def date(self, root,str_group_mth,str_group_day,str_group_year):
        self.dentry = popupWindowDate(root, str_group_mth, str_group_day, str_group_year,font=('Helvetica', 40, tkinter.NORMAL), border=0)
        root.wait_window(self.dentry.top)
        str_group_mth.set(self.dentry.value[0])
        str_group_day.set(self.dentry.value[1])
        str_group_year.set(self.dentry.value[2])
        return str(self.dentry.value[0])+'/'+str(self.dentry.value[1])+'/'+str(self.dentry.value[2])
    def getSelectableRTIDs(self,readCol):
        rwh= RWHANDLE(o.infoFile, readingCol=readCol)
        listOfRT= rwh.collectFromDB(typeOfData=o.listCol,colOfValues=readCol)
        return listOfRT
    def getSelectableLocations(self):
        rwh= RWHANDLE(o.infoFile, o.colHomeList)
        listOfLocations= rwh.collectFromDB(typeOfData=o.listCol,colOfValues=o.colHomeList)
        return listOfLocations

class FP_ITEM(object):
    def __init__(self,QTY,MANU,DES,PRI,DateRec,mth,day,year,orderID):
        #set(QTY,MANU,DES,PRI,DateRec)
        self.belongsTo=orderID
        self.footObj = {}
        self.footObj={QTY:0, MANU:0, DES:0, PRI:0, DateRec:0,mth:0,day:0,year:0}#fromkeys(set,0)
        self.wcObj = {}
        self.name = QTY
        self.qty = 0
        self.manu = 0
        self.des = 0
        self.pri = 0
        self.datRec = 0
        self.en = False

    def enable(self):
        self.en =True

    def updateIndepObj(self):
        self.name = self.footObj[0].key() # the qty label will serve as name
        self.qty = self.footObj[0].value()
        self.manu = self.footObj[1].value()
        self.des = self.footObj[2].value()
        self.pri = self.footObj[3].value()
        self.datRec = self.footObj[4].value()
        self.mth=self.footObj[5].value()
        self.day=self.footObj[6].value()
        self.year=self.footObj[7].value()

    def get(self,query):
        return self.footObj.get(query, default=0)

class ISSUE():
    def get(self,query):
        return self.issueItems.get(query, default=0)
    def getItem(self,query):
        self.myItem = FP_ITEM
        self.myItem = self.issueItems['ITEMS'][query]
    def addFile(self,name):
        self.attachFiles.append(name)
    def updateIndepObj(self):
        self.SHORTD = self.issueItems[0].value()
        self.Supplier = self.issueItems[1].value()
        self.Special__bInstructions = self.issueItems[2].value()
        self.Requested__bBy = self.issueItems[3].value()
        self.US__bFunds = self.issueItems[4].value()
        self.Quote__b__3 = self.issueItems[5].value()
        self.Customer__b__3 = self.issueItems[6].value()
        self.Receiver = self.issueItems[7].value()
        self.Budget = self.issueItems[8].value()
        self.Account__b__3 = self.issueItems[9].value()
        self.C__fO = self.issueItems[10].value()
        self.Fax__b__3 = self.issueItems[11].value()
        self.PRIORITY = self.issueItems[12].value()
        self.NEWSTATUS = self.issueItems[13].value()
        self.Email__baddress = self.issueItems[14].value()
        self.LADD = self.issueItems[15].value()
        self.ASSIGNTO = self.issueItems[16].value()
        self.PO__3 = self.issueItems[17].value()
        # CHECKBOXES
        self.Order__bDateTODAY = self.issueItems[18].value()
        self.MAIL_ASSIGNEES = self.issueItems[19].value()
        self.MAIL_ENDUSER = self.issueItems[20].value()
        self.MAIL_CC = self.issueItems[21].value()
    def addItem(self,item,itemDict):
        self.itemDict.append(item)
        return itemDic
    def makeItemDictList(self,n=11):
        itemDictList={}

        for i in range(0, n):
            #item = self.dictOfGroupMan.get(items)
            groupn = str(i+1)
            itemDictList['Item__b__3' + groupn] = FP_ITEM('Item__b__3' + groupn + '__bQuantity'
                                                          , 'Item__b__3' + groupn + '__bManu'
                                                          , 'Item__b__3' + groupn
                                                          , 'Item__b__3' + groupn + '__bPrice'
                                                          , 'Item__b__3' + groupn + '__bReceived__bDateTODAY'
                                                          , 'MonthInput_Item__b__3' + groupn + '__bReceived__bDate'
                                                          , 'DayInput_Item__b__3' + groupn + '__bReceived__bDate'
                                                          , 'YearInput_Item__b__3' + groupn + '__bReceived__bDate'
                                                          , self.belongTo)
        # itemDictList['Item__b__32'] = FP_ITEM('Item__b__32__bQuantity', 'Item__b__32__bManu', 'Item__b__32',
        #                                       'Item__b__32__bPrice', 'Item__b__32__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__32__bReceived__bDate',
        #                                       'DayInput_Item__b__32__bReceived__bDate',
        #                                       'YearInput_Item__b__32__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__33'] = FP_ITEM('Item__b__33__bQuantity', 'Item__b__33__bManu', 'Item__b__33',
        #                                       'Item__b__33__bPrice', 'Item__b__33__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__33__bReceived__bDate',
        #                                       'DayInput_Item__b__33__bReceived__bDate',
        #                                       'YearInput_Item__b__33__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__34'] = FP_ITEM('Item__b__34__bQuantity', 'Item__b__34__bManu', 'Item__b__34',
        #                                       'Item__b__34__bPrice', 'Item__b__34__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__34__bReceived__bDate',
        #                                       'DayInput_Item__b__34__bReceived__bDate',
        #                                       'YearInput_Item__b__34__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__35'] = FP_ITEM('Item__b__35__bQuantity', 'Item__b__35__bManu', 'Item__b__35',
        #                                       'Item__b__35__bPrice', 'Item__b__35__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__35__bReceived__bDate',
        #                                       'DayInput_Item__b__35__bReceived__bDate',
        #                                       'YearInput_Item__b__35__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__36'] = FP_ITEM('Item__b__36__bQuantity', 'Item__b__36__bManu', 'Item__b__36',
        #                                       'Item__b__36__bPrice', 'Item__b__36__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__36__bReceived__bDate',
        #                                       'DayInput_Item__b__36__bReceived__bDate',
        #                                       'YearInput_Item__b__36__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__37'] = FP_ITEM('Item__b__37__bQuantity', 'Item__b__37__bManu', 'Item__b__37',
        #                                       'Item__b__37__bPrice', 'Item__b__37__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__37__bReceived__bDate',
        #                                       'DayInput_Item__b__37__bReceived__bDate',
        #                                       'YearInput_Item__b__37__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__38'] = FP_ITEM('Item__b__38__bQuantity', 'Item__b__38__bManu', 'Item__b__38',
        #                                       'Item__b__38__bPrice', 'Item__b__38__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__38__bReceived__bDate',
        #                                       'DayInput_Item__b__38__bReceived__bDate',
        #                                       'YearInput_Item__b__38__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__39'] = FP_ITEM('Item__b__39__bQuantity', 'Item__b__39__bManu', 'Item__b__39',
        #                                       'Item__b__39__bPrice', 'Item__b__39__bReceived__bDateTODAY',
        #                                       'MonthInput_Item__b__39__bReceived__bDate',
        #                                       'DayInput_Item__b__39__bReceived__bDate',
        #                                       'YearInput_Item__b__39__bReceived__bDate', self.belongTo)
        # itemDictList['Item__b__310'] = FP_ITEM('Item__b__310__bQuantity', 'Item__b__310__bManu',
        #                                        'Item__b__310', 'Item__b__310__bPrice',
        #                                        'Item__b__39__bReceived__bDateTODAY',
        #                                        'MonthInput_Item__b__310__bReceived__bDate',
        #                                        'DayInput_Item__b__310__bReceived__bDate',
        #                                        'YearInput_Item__b__310__bReceived__bDate', self.belongTo)

        return itemDictList
    def makeIssue(self,num):
        set=('SHORTD'
                ,'Supplier'
                ,'Special__bInstructions'
                ,'Requested__bBy'
                ,'US__bFunds'
                ,'Quote__b__3'
                ,'Customer__b__3'
                ,'Receiver'
                ,'Budget'
                ,'Account__b__3'
                ,'C__fO'
                ,'Fax__b__3'
                ,'PRIORITY'
                ,'NEWSTATUS'
                ,'Email__baddress'
                ,'LADD'
                ,'ASSIGNTO'
                ,'PO__3'
                #CHECKBOXES
                ,'Order__bDateTODAY'
                ,'MAIL_ASSIGNEES'
                ,'MAIL_ENDUSER'
                ,'MAIL_CC'
                ,'DATE_S_YearInput_Order__bDate_S_Year'
                ,'DATE_S_MonthInput_Order__bDate_S_Month'
                ,'DATE_S_DayInput_Order__bDate_S_Day'
                #LIST
                ,'ITEMS')
        self.issueItems = {}
        self.issueItems.fromkeys(set, 1)
        self.issueItems['ITEMS']= self.makeItemDictList(num)
        return self.issueItems
    def __init__(self,orderID,num):#,order,priority,status,po,supplier,contactEmail,Items,request,quote,customer,receiver,budget,account,co,fax,description, offPerson,sendEmailOff,sendEmailContact):
        self.belongTo=orderID
        self.issueItems=self.makeIssue(num)
        self.itemDictList={}
        self.attachedFiles=[]
        ## BELOW IS OPTIONAL for detailed issues variables
        self.SHORTD = 0
        self.Supplier = 0
        self.Special__bInstructions = 0
        self.Requested__bBy = 0
        self.US__bFunds = 0
        self.Quote__b__3 = 0
        self.Customer__b__3 = 0
        self.Receiver = 0
        self.Budget = 0
        self.Account__b__3 = 0
        self.C__fO = 0
        self.Fax__b__3 = 0
        self.PRIORITY = 0
        self.NEWSTATUS = 0
        self.Email__baddress = 0
        self.LADD = 0
        self.ASSIGNTO = 0
        self.PO__3 = 0
        # CHECKBOXES
        self.Order__bDateTODAY = 0
        self.MAIL_ASSIGNEES = 0
        self.MAIL_ENDUSER = 0
        self.MAIL_CC = 0
        # LIST
        self.ITEMS = 0
    #Receiving data form database
    #must then lock if upload in footprint
    def registerData(self,data):# data is a dict
        size= len(data)
        i=0
        #keyss=[]
        keyss = list(data.items())
        while(1):
            try:
                if i==size:
                    break
                else:
                    obj = keyss[i][0]
                    #print(keyss[i])
                    #debug(str(obj))
                    if obj is None:
                        i+=1
                    elif type(obj) is str:
                        if 'Item' in obj:
                            for y in range(0,7):
                                if y==0:#len('Item__b__xx') == len(obj):
                                    itemName = keyss[i+2][0]
                                if y==1:
                                    obj = keyss[i][0]
                                    b=data.get(obj)
                                    i+=1
                                    obj = keyss[i][0]
                                    b+=(' '+ data.get(obj))
                                    self.issueItems['ITEMS'][itemName].footObj[obj] = b
                                    i+=1
                                else:
                                    try:
                                        obj = keyss[i][0]
                                        b = data.get(obj)
                                        self.issueItems['ITEMS'][itemName].footObj[obj] = b
                                    except Exception as e:
                                        debug(str(e) + ', Location in ui.ISSUE.registerData() ')
                                        pass
                                    i=i+1
                        elif type(obj)== str:
                            b=data.get(obj)
                            self.issueItems[obj] = b
                            i=i+1
                        else:
                            i=i+1
            except Exception as e:
                debug(str(e)+', Location in ui.ISSUE.registerData() ')
                break
    def __str__(self):
        return("Order objects:\n"
                "  order              = {0}\n"
               "  priority            = {1}\n"
               "  status              = {2}\n"
               "  po                  = {3}\n"
               "  supplier            = {4}\n"
               "  contactEmail        = {5}\n"             
               "  request             = {7}\n"
               "  quote               = {8}\n"
               "  Items               = {6}\n"
               "  customer            = {9}\n"
               "  receiver            = {10}\n"
               "  budget              = {11}\n"
               "  account             = {12}\n"
               "  co                  = {13}\n"
               "  fax                 = {14}\n"
               "  description         = {15}\n"
               "  offPerson           = {16}\n"
               "  sendEmailOff        = {17}\n"
               "  sendEmailContact    = {18}")
                #.format(self.order ,self.priority,self.status,self.po,self.supplier,self.contactEmail,self.Items,self.request,self.quote,self.customer,self.receiver,self.budget ,self.account,self.co ,self.fax,self.description,self.offPerson ,self.sendEmailOff,self.sendEmailContact)


              # .format(self.id, self.dsp_name, self.dsp_code,
               #        self.hub_code, self.pin_code, self.pptl))


        # ## DROPBOX
            # optional
            # color = orderSelection.get()
            # root['bg'] = color
        # use width x height + x_offset + y_offset (no spaces!)
        #root.geometry("%dx%d+%d+%d" % (2000, 6000, 200, 150))
        #root.title("tk.OptionMenu Testing as combobox")

