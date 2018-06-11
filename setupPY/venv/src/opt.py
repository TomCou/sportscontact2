# MISC
class O:
    sizeStRow=2
    sizeStCol=5
    cdpCol=0
    unbrandedItemAddingRow=3
    maxShoesSheetRow=300
    defaultShoeFile= 'inv_soulier'
    listOfDepartmentR=['socJR_tur_R',
                      'socJR_ind_R',
                      'socJR_out_R',
                      'socAD_tur_R',
                      'socAD_ind_R',
                      'socAD_out_R',
                      'socWAD_out_R',
                    'baseJR_R',
                    'baseAD_R',
                      'bbJR_R',
                      'bbAD_R',
                      'rugbyAD_R',
                       'rugbyJR_R',
                      'footJR_R',
                      'footAD_R',
                      'courseH_R',
                      'courseF_R',
                      'courseJR_R',
                      'vbH_R',
                      'vbF_R',
                        'entreH_R',
                        'entreF_R',
                        'entreJR_R',
                        'cheerF_R',
                        'cheerJR_R'
                       ]
    listOfDepartment = ['socJR_tur',
                         'socJR_ind',
                         'socJR_out',
                         'socAD_tur',
                         'socAD_ind',
                         'socAD_out',
                         'socWAD_out',
                        'baseJR',
                        'baseAD',
                         'bbJR',
                         'bbAD',
                         'rugbyAD',
                        'rugbyJR',
                        'footJR',
                         'footAD',
                         'courseH',
                         'courseF',
                         'courseJR',
                         'vbH',
                         'vbF',
                        'entreH',
                        'entreF',
                        'entreJR',
                        'cheerF',
                        'cheerJR'
                         ]
    dictNameOfDepNtoR = {
                        'socJR_tur': 'socJR_tur_R',
                        'socJR_ind':'socJR_ind_R',
                        'socJR_out':'socJR_out_R',
                        'socAD_tur':'socAD_tur_R',
                        'socAD_ind':'socAD_ind_R',
                        'socAD_out':'socAD_out_R',
                        'socWAD_out':'socWAD_out_R',
                        'baseJR':'baseJR_R',
                        'baseAD':'baseAD_R',
                        'bbJR':'bbJR_R',
                        'bbAD':'bbAD_R',
                        'rugbyAD':    'rugbyAD_R',
                        'rugbyJR':    'rugbyJR_R',
                        'footJR':'footJR_R',
                        'footAD':'footAD_R',
                        'courseH':'courseH_R',
                        'courseF':'courseF_R',
                        'courseJR':'courseJR_R',
                        'vbH':'vbH_R',
                        'vbF':'vbF_R',
                        'entreH':   'entreH_R',
                        'entreF': 'entreF_R',
                        'entreJR': 'entreJR_R',
                        'cheerF': 'cheerF_R',
                        'cheerJR': 'cheerJR_R'
                        }
    dictNameOfDepRtoN = {
                        'socJR_tur_R': 'socJR_tur',
                        'socJR_ind_R':'socJR_ind',
                        'socJR_out_R':'socJR_out',
                        'socAD_tur_R':'socAD_tur',
                        'socAD_ind_R':'socAD_ind',
                        'socAD_out_R':'socAD_out',
                        'socWAD_out_R':'socWAD_out',
                        'baseJR_R':'baseJR',
                        'baseAD_R':'baseAD',
                        'bbJR_R':     'bbJR',
                        'bbAD_R':     'bbAD',
                        'rugbyAD_R':    'rugbyAD',
                        'rugbyJR_R':    'rugbyJR',
                        'footJR_R':   'footJR',
                        'footAD_R':   'footAD',
                        'courseH_R':  'courseH',
                        'courseF_R':  'courseF',
                        'courseJR_R': 'courseJR',
                        'vbH_R':      'vbH',
                        'vbF_R':      'vbF',
                        'entreH_R':   'entreH',
                        'entreF_R': 'entreF',
                        'entreJR_R': 'entreJR',
                        'cheerF_R': 'cheerF',
                        'cheerJR_R': 'cheerJR'

    }

    dictOfSizes2ColumnPosition={
        "6K": 4,
        "6.5K": 5,
        "7K": 6,
        "7.5K": 7,
        "8K": 8,
        "8.5K": 9,
        "9K": 10,
        "9.5K": 11,
        "10K": 12,
        "10.5K": 13,
        "11K": 14,
        "11.5K": 15,
        "12K": 16,
        "12.5K": 17,
        "13K": 18,
        "13.5K": 19,
        "1": 20,
        "1.5": 21,
        "2": 22,
        "2.5": 23,
        "3": 24,
        "3.5": 25,
        "4": 26,
        "4.5": 27,
        "5": 28,
        "5.5": 29,
        "6": 30,
        "6.5": 31,
        "7": 32,
        "7.5": 33,
        "8": 34,
        "8.5": 35,
        "9": 36,
        "9.5": 37,
        "10": 38,
        "10.5": 39,
        "11": 40,
        "11.5": 41,
        "12": 42,
        "12.5": 43,
        "13": 44,
        "13.5": 45,
        "14": 46,
        "14.5": 47,
        "15": 48,
        "16": 49,
        "17": 50,
        "18": 51,
        "": -1}

    scSoulierFileCH='inv_soulier_HB'
    scSoulierFileSN = 'inv_soulier_SN'
    excelFolder = '//SCONTACTSRV/Public/invCustom/'
    DEBUG = True
    fromKevinExcelSheet=1
    fromFootprint = 2
    fromWebCheckout=3
    fromHalSS=4
    fromMakeProcPDF=5
    fromMainList = 6
    colOfPOID=3
    defaultRowID = 8
    getInfo = 8 # if you want the current order id and stuff
    newOrder=10
    update=11
    PO='PO'
    kevinCheck='DkevinSS'
    fpCheck='UfootPrint'
    wcCheck='UwebCheckout'
    halCheck='UhalSS'
    procPDFCheck='UmakeProcPDF'
    rowID='rowID'
    uploadButton=12
    colOfOrderRowID =2
    list='list'
    single='single'
    dicto='dicto'
    apiCol=6
    apiRow=2
    listOfList='listOfList'
    dictoOfList='dictoOfList'
    listRow='listRow'
    listCol = 'listCol'
    listOfRes='listOfRes'
    deleteOrder=13
    colRTList=6
    colRTIDList = 7
    colHomeList=4
    colHomeAlias=5
    colOfLastResID=1
    matchColValue='matchColValue'
    fp_pass="aqaSWS#1erd"
    fp_user="tcousine"
    wc_pass="scooter"
    wc_user="3474142"
    bc_pass = ""
    bc_user = "tommy.cousineau@unb.ca"
    defaultColumn=1
    sg=1
    text=2
    itFirstRow=31
    itGrSep=100
    qtyC=1
    numbOfDescriptiveLineForEachItemChild=8


    if DEBUG:
        mspr_folder='.\\MSReceivedGoods\\MSPR\\unprocessed\\'
        orderDatabase = ".\\ui_database\\ui_orderTracker"
        itemDatabase = ".\\ui_database\\ui_itemTrackerFiles\\"
        infoFile = ".\\ui_database\\ui_currentStatusInfo"
        attachFileFolder="\\MSReceivedGoods\\MSPR\\unprocessed\\"
        csvFolder='.\\ui_database\\ui_webcheckout\\'
        emailTo=['Tcousine:','kennye:']
        dontEmailTo=['hdalzell:','dkell:',"rking:","kwm:"]
#        emailTo2='kennye'
    else:
        mspr_folder = 'C:\\Users\\tcousine\\Desktop\\MSReceivedGoods\\MSPR\\'
        orderDatabase = ".\\ui_database\\ui_orderTracker"
        itemDatabase = ".\\ui_database\\ui_itemTrackerFiles\\"
        infoFile = ".\\ui_database\\ui_currentStatusInfo"
        csvFolder = '.\\ui_database\\ui_webcheckout\\'
        emailTo = ['Tcousine:', 'kennye:']
        dontEmailTo = ['hdalzell:', 'dkell:', "rking:", "kwm:"]
    sampleFile = 'currentOrder.xlsx'



    def throwError(prob,value,frameinfo):
        print(prob, value)
        print("Come and Fix Me Here:", frameinfo.filename, frameinfo.lineno)


def ImportCheck():

    import pip
    import subprocess
    import os
    import sys

    def install(package):
        try:
        #pip.main(['install', package])
            subprocess.check_call([sys.executable, '-m', 'pip', 'install', package])
        except:
            print(package+" didnt work")

    # Example
    #if __name__ == '__main__':
    #   install('argh')

    #pywinauto
    # try:
    #     import pywinauto
    # except ImportError:
    #     install('pywinauto')


    #datetime
    try:
        import datetime
    except ImportError:
        install('datetime')

    #csv
    try:
        import csv
    except ImportError:
        install('csv')

    #tkinter
    try:
        # Python3
        import tkinter
    except ImportError:
        # Python2
        import Tkinter


    #xlrd
    try:
        import xlrd
    except ImportError:
        install('xlrd')

    #os
    try:
        import os
    except ImportError:
        install('os')

    #xlutils
    try:
        import xlutils
    except ImportError:
        install('xlutils')
        install('xlwt')

    #openpyxl
    try:
        import openpyxl
    except ImportError:
        install('openpyxl')


    #selenium
    try:
        import selenium
    except ImportError:
        install('selenium')
