#!/usr/bin/env python
# -*- coding: utf-8 -*-
LOGIN_SERVER_IP="member.medicalfields.jp"
VERSION_INFO="1.1"
from functools import partial
from PyQt5.QtCore import QLocale, QTranslator, QLibraryInfo
import configparser
import os
import requests
import unicodedata
from requests.exceptions import Timeout
from requests.exceptions import ProxyError
import json
import winreg as winreg
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (QAction, QApplication, QCheckBox, QComboBox,
        QDialog, QGridLayout, QGroupBox, QHBoxLayout, QLabel, QLineEdit,
        QMessageBox, QMenu, QPushButton, QSpinBox, QStyle, QSystemTrayIcon,QFileDialog,
        QTextEdit, QVBoxLayout)
from PyQt5.QtCore import (QCoreApplication, QObject, QRunnable, QThread,
                          QThreadPool, pyqtSignal)
from PyQt5.QtGui import (QIntValidator)
from PyQt5.QtCore import Qt
import sys
import glob
import time
import csv
import systray_rc
import re
from chardet.universaldetector import UniversalDetector
#>Pyrcc5 systray/systray.qrc -o systray/systray_rc.py
global proxyIP,proxyPort,useInternetExplorerSetting,useProxy,IEProxyIP,IEProxyPort,proxyDict
global fileType,fileTypeVer,fileFolder,fileFolderTemp,fileFolderCheckType,fileSuccessFlag,updateDetect,privacyInfo
global loginName,loginPass,isLogined,userID,userName,userEmail,MFtoken,loopCount,onlyStartUpProcess
updateDetect=False
isLogined=False
onlyStartUpProcess=True
loopCount=0
fileSuccessFlag=False
fileFolderTemp=""
fileFolderCheckType=0
import base64
from os import path

FOLDER_PATH_REGEX = r'^[a-zA-Z]:\\(((?![<>:"/\\|?*]).)+((?<![ .])\\)?)*$'
NETWORK_FOLDER_PATH_REGEX = r'^\\\\(((?![<>:"/\\|?*]).)+((?<![ .])\\)?){1,}\\(((?![<>:"/\\|?*]).)+((?<![ .])\\)?)*$'
MESSAGE_NO = 0
MESSAGE_NO_MAIN_FOLDER=3
MESSAGE_NO_INDEX_FOLDER=1
MESSAGE_NOT_FIND_TXT=2

MESSAGE_NO_MAIN_FOLDER_="<b>エラー！</b>\u2029連動用フォルダに入力された場所にフォルダが存在しません\u2029" \
                         "一度レセコンより連動用のデータを出力してください\u2029" \
                         "もしレセコンから連動用のデータを出力したのにも関わらずこの画面が再び出た場合は連動用フォルダの入力が間違っている可能性が高いです\u2029" \
                         "各メーカーのNSIPSの設定画面で指定したフォルダが正しく入力または選択されているかご確認下さい。多くの場合、連動用フォルダは【sips1】や【nsips1】などの文字で終わることが多いです。\u2029" \
                        "また、すぐに連動用データを出力出来ない場合で、入力された場所に出力されることが確実な場合はフォルダを新規作成して下さい"

MESSAGE_NO_INDEX_FOLDER_="<b>エラー！</b>\u2029連動用フォルダに入力された場所にフォルダは存在しますが、レセコンと連動するために必要フォルダ（レセコンからの処方箋データがあるINDEXフォルダとDATAフォルダ）が見つかりませんでした\u2029" \
                         "一度レセコンより連動用のデータを出力してください\u2029" \
                         "もしレセコンから連動用のデータを出力したのにも関わらずこの画面が再び出た場合は連動用フォルダの入力が間違っている可能性が高いです\u2029" \
                         "各メーカーのNSIPSの設定画面で指定したフォルダが正しく入力または選択されているかご確認下さい。多くの場合、連動用フォルダは【sips1】や【nsips1】などの文字で終わることが多いです。"

SET_VISUALITY_FALSE=4
MESSAGE_INTERNET_ERROR=5
MESSAGE_OTHER_FILE_DETECT=6
MESSAGE_OTHER_FILE_DETECT_="<b>エラー！</b>\u2029Nsipsの連動設定がされているindexフォルダに別のファイルが見つかりました。\u2029通常indexフォルダにはtxtファイルのみしか存在できません。\u2029" \
                           "一度フォルダの中身を空にするか、連動設定を再構築してください。"
MESSAGE_INDEX_EXIST_BUT_DATA_NOT_EXIST=7
MESSAGE_INDEX_EXIST_BUT_DATA_NOT_EXIST_="<b>エラー！</b>\u2029Nsipsの連動設定がされているindexフォルダの内容とdataフォルダの内容が一致しません。これは連動データが破損している可能性があります\u2029" \
                                        "一度フォルダの中身を空にするか、連動設定を再構築してください。"
MESSAGE_SEND_CSV=8
MESSAGE_UPDATE_DETECT=9
MESSAGE_404_DETECT=10
MESSAGE_UPLOAD_TEXT_ERROR_DETECT=11
MESSAGE_UPLOAD_FILE_TYPE_ERROR_DETECT=12
MESSAGE_400_DETECT=13
import objgraph
import subprocess
from PyQt5 import QtWidgets

# 2重起動を防ぐ
if os.name == 'posix':
    print('on Mac or Linux')
    file_name = path.basename(__file__)
    p1 = subprocess.Popen(["ps", "-ef"], stdout=subprocess.PIPE)
    p2 = subprocess.Popen(["grep", file_name], stdin=p1.stdout, stdout=subprocess.PIPE)
    p3 = subprocess.Popen(["grep", "python"], stdin=p2.stdout, stdout=subprocess.PIPE)
    p4 = subprocess.Popen(["wc", "-l"], stdin=p3.stdout, stdout=subprocess.PIPE)
    p1.stdout.close()
    p2.stdout.close()
    p3.stdout.close()
    output = p4.communicate()[0].decode("utf8").replace('\n', '')
    if int(output) != 1:
        exit()
elif os.name == 'nt':

    import win32api
    import win32event
    import winerror
    import pywintypes
    print('on Windows')
    UNIQUE_MUTEX_NAME = 'Global\\MyProgramIsAlreadyRunning'
    handle = win32event.CreateMutex(None, pywintypes.FALSE, UNIQUE_MUTEX_NAME)

    if not handle or win32api.GetLastError() == winerror.ERROR_ALREADY_EXISTS:
        print('既に別のプロセスが実行中です。', file=sys.stderr)
        app = QApplication(sys.argv)
        app.setAttribute(Qt.AA_DisableWindowContextHelpButton)
        QMessageBox.information(None, "注意",
                "MediMonitorレセプトコンピュータ連動ソフトウェアはすでに起動しています.システムトレイをご確認ください.")
        sys.exit(-1)
    print('このプロセスだけが実行中です。')

def get_internet_key(name):
    ie_settings = winreg.OpenKey(winreg.HKEY_CURRENT_USER,
                                 r'Software\Microsoft\Windows\CurrentVersion\Internet Settings', 0,
                                 winreg.KEY_ALL_ACCESS)
    value, type = winreg.QueryValueEx(ie_settings, name)
    return value


def hasInternetProxy():
    """Returns boolean. If key doesn't exist, returns false."""
    if get_internet_key('ProxyEnable') == 1:
        return True
    else:
        return False


def getProxyAddresses():
    """Returns dictionary of protocol:address key:value pairs."""
    proxies = {}
    if hasInternetProxy():
        rawStr = get_internet_key("ProxyServer")
        separated = rawStr.split(";")
        if len(separated) == 1:
            proxies["all"] = separated[0]
        else:
            for single_proxy in separated:
                protocol, address = single_proxy.split("=")
                proxies[protocol] = address
            if 'all' not in proxies.keys() and 'http' in proxies.keys():
                proxies['all'] = proxies["http"]
    return proxies
def getIEsetting(self):
    global IEProxyIP,IEProxyPort
    proxyAll = getProxyAddresses().get("all")
    if proxyAll != None:
        proxyAllMap = proxyAll.split(":")
        IEProxyIP = proxyAllMap[0]
        IEProxyPort = proxyAllMap[1]
    else:
        IEProxyIP = ""
        IEProxyPort = ""

class Window(QDialog):

    def folderErrorCheckTemp(self, fileFolder):

        global indexDirTemp, dataDirTemp
        indexDirTemp = ""
        dataDirTemp = ""
        if os.path.exists(fileFolder + "\\index\\"):
            indexDirTemp = fileFolder + "\\index\\"
        elif os.path.exists(fileFolder + "\\INDEX\\"):
            indexDirTemp = fileFolder + "\\INDEX\\"
        elif os.path.exists(fileFolder + "\\Index\\"):
            indexDirTemp = fileFolder + "\\Index\\"
        if os.path.exists(fileFolder + "\\data\\"):
            dataDirTemp = fileFolder + "\\data\\"
        elif os.path.exists(fileFolder + "\\DATA\\"):
            dataDirTemp = fileFolder + "\\DATA\\"
        elif os.path.exists(fileFolder + "\\Data\\"):
            dataDirTemp = fileFolder + "\\Data\\"
        if len(indexDirTemp) == 0:
            return False
        if len(dataDirTemp) == 0:
            return False

        return True
    class ConcurrentlyWorker(QThread):
        finSignal = pyqtSignal(int,int)  # 扱うデータがある場合は型を指定する
        loopFlag = 1

        def run(self):
            print('threading...')
            while (True):
                global onlyStartUpProcess,fileSuccessFlag,updateDetect
                global MFtoken,sendCsv,sendCsvName
                if updateDetect:
                    self.finSignal.emit(loopCount, MESSAGE_UPDATE_DETECT)
                    print(MFtoken + "updateDetect")
                    break

                if len(MFtoken) == 0:
                    print(MFtoken + "tokenがないです")
                    break
                if onlyStartUpProcess:
                    self.connectHTTPandSendCSV("","")#これはIDのチェックのため
                    if self.folderErrorCheck() and fileSuccessFlag:
                        self.finSignal.emit(loopCount, SET_VISUALITY_FALSE)

                    onlyStartUpProcess=False
                else:
                    if (self.loopFlag == 1):
                        if self.folderErrorCheck():
                            self.folderCheckAndUpload()
                    else:
                        break
        def folderErrorCheck(self):
            global indexDir,dataDir
            indexDir = ""
            dataDir = ""
            if os.path.exists(fileFolder + "\\index\\"):
                indexDir = fileFolder + "\\index\\"
            elif os.path.exists(fileFolder + "\\INDEX\\"):
                indexDir = fileFolder + "\\INDEX\\"
            elif os.path.exists(fileFolder + "\\Index\\"):
                indexDir = fileFolder + "\\Index\\"
            if os.path.exists(fileFolder + "\\data\\"):
                dataDir = fileFolder + "\\data\\"
            elif os.path.exists(fileFolder + "\\DATA\\"):
                dataDir = fileFolder + "\\DATA\\"
            elif os.path.exists(fileFolder + "\\Data\\"):
                dataDir = fileFolder + "\\Data\\"
            if len(indexDir) == 0:
                print("noIndexDir")
                self.finSignal.emit(loopCount, MESSAGE_NO_INDEX_FOLDER)
                time.sleep(5)
                return False

            if len(dataDir) == 0:
                print("noDataDir")
                self.finSignal.emit(loopCount, MESSAGE_NO_INDEX_FOLDER)
                time.sleep(5)
                return False
            return True


        def folderCheckAndUpload(self):

            global indexDir,dataDir
            global loginName, loginPass, isLogined, userID, userName, userEmail, MFtoken, proxyDict,loopCount
            global response,fileSuccessFlag,updateDetect
            global fileType,fileTypeVer,fileFolder,privacyInfo
            print('Loop...')
            csvName=""
            dirLoopCount=0
            while (True):
                if (self.loopFlag != 1):
                    break
                files = glob.glob(indexDir+"*")
                otherFileDetect=False
                indexExistButDataNotExist=False
                for file in files:
                    if file.endswith("txt") or file.endswith("TXT") :
                        print("doUpload")
                        if not fileSuccessFlag :
                            fileSuccessFlag=True
                            configSaver.saveConfig(self)
                        basename = os.path.basename(file)
                        dataFilePath=dataDir+basename
                        csvBase64=""
                        sendCsv=""
                        if os.path.exists(dataFilePath):

                            fileTypeError=True
                            someError = True
                            lookup = ('cp932', 'utf_8_sig', 'utf_8')
                            for trying_encoding in lookup:
                                try:
                                    for line in open(dataFilePath, "r", encoding=trying_encoding):
                                        if fileType == 0:
                                            if line.startswith("VER"):
                                                fileTypeError = False
                                        print(line)
                                        break
                                    print(fileTypeError)
                                    if not fileTypeError:
                                        if privacyInfo == 0:
                                            # with open(dataFilePath, "rb") as f:
                                            #     csvBase64 = base64.b64encode(f.read())

                                            with open(dataFilePath, encoding=trying_encoding) as csvfile:
                                                s = csvfile.read()
                                                csvBase64 = base64.b64encode(
                                                    s.encode("shift_jis", errors="ignore"))
                                        else:
                                            with open(dataFilePath, encoding=trying_encoding) as csvfile:
                                                csvBase = self.deletePatientInfo(csvfile)
                                                csvBase64 = base64.b64encode(
                                                    csvBase.encode("shift_jis", errors="ignore"))

                                        self.connectHTTPandSendCSV(csvBase64, basename)
                                        if response is not None:
                                            if response.status_code == 200:
                                                r = json.loads(response.text)
                                                if "this" in r:
                                                    if r["this"] == "ok":
                                                        someError = False
                                                        okToDelete = True
                                                        while okToDelete:
                                                            try:
                                                                os.remove(file)
                                                                okToDelete = False
                                                            except PermissionError:
                                                                print('パーミッションエラーです。')
                                                                time.sleep(0.5)
                                                            except Exception:
                                                                print('その他の例外です。')
                                                                time.sleep(5)

                                                        okToDelete = True
                                                        while okToDelete:
                                                            try:
                                                                os.remove(dataFilePath)
                                                                okToDelete = False
                                                            except PermissionError:
                                                                print('パーミッションエラーです。')
                                                                time.sleep(0.5)
                                                            except Exception:
                                                                print('その他の例外です。')
                                                                time.sleep(5)
                                                    if r["this"] == "needUpdate":
                                                        updateDetect = True

                                            if response.status_code == 404:
                                                self.finSignal.emit(loopCount, MESSAGE_404_DETECT)
                                                time.sleep(300)
                                            if response.status_code == 400:
                                                someError = False
                                                self.finSignal.emit(loopCount, MESSAGE_400_DETECT)
                                                break

                                    if updateDetect:
                                        print("アップロード中にアップデートが見つかりました")
                                        self.finSignal.emit(loopCount, MESSAGE_UPDATE_DETECT)
                                        time.sleep(86400)
                                        break
                                    if fileTypeError:
                                        print("ＣＳＶのファイルにエラーがありました")
                                        self.finSignal.emit(loopCount,
                                                            MESSAGE_UPLOAD_FILE_TYPE_ERROR_DETECT)
                                        time.sleep(5)
                                    elif someError:
                                        print("ＣＳＶのアップロードにエラーがありました")
                                        self.finSignal.emit(loopCount,
                                                            MESSAGE_UPLOAD_TEXT_ERROR_DETECT)
                                        time.sleep(5)


                                    break
                                except UnicodeDecodeError:
                                    print ("UnicodeDecodeError")
                                    continue



                        else:
                            indexExistButDataNotExist=True


                        print(dataFilePath)
                    else:
                        print(file)
                        otherFileDetect=True

                if dirLoopCount == 50:
                    print("endIndexLoop"+str(dirLoopCount))
                    if otherFileDetect:
                        self.finSignal.emit(loopCount, MESSAGE_OTHER_FILE_DETECT)
                    elif updateDetect:
                        self.finSignal.emit(loopCount, MESSAGE_UPDATE_DETECT)
                    elif indexExistButDataNotExist:
                        self.finSignal.emit(loopCount, MESSAGE_INDEX_EXIST_BUT_DATA_NOT_EXIST)
                    else :
                        self.finSignal.emit(loopCount, MESSAGE_NOT_FIND_TXT)
                    break
                dirLoopCount=dirLoopCount+1
                time.sleep(0.1)

        def connectHTTPandSendCSV(self,sendCsv,sendCsvName):

            global loginName, loginPass, isLogined, userID, userName, userEmail, MFtoken, proxyDict,loopCount
            global response
            global fileType,fileTypeVer,fileFolder
            print('Loop...')

            loopCount=loopCount+1  # 0から1万の範囲で乱数生成
            httpsRequest.asyncTokenRequest(self, proxyDict, MFtoken, userID, sendCsv, sendCsvName,fileType,fileTypeVer)
            if response is None:

                while (True):
                    print(" response is None while (True)"+str(loopCount))
                    self.finSignal.emit(loopCount,MESSAGE_INTERNET_ERROR)
                    httpsRequest.asyncTokenRequest(self, proxyDict, MFtoken, userID, sendCsv, sendCsvName,fileType,fileTypeVer)
                    time.sleep(5)
                    if response is not None:
                        break

            if len(sendCsv)==0:
                self.finSignal.emit(loopCount,MESSAGE_NO)
            else:
                self.finSignal.emit(loopCount,MESSAGE_SEND_CSV)


        def deletePatientInfo(self,csvfile):
            sendCsv=""
            reader = csv.reader(csvfile)
            for row in reader:
                if row[0].startswith("VER"):# ヘッダ部
                    row[6] =""
                    row[7] =""
                    row[8] =""
                    row[9] =""
                    row[10] =""
                if row[0] == "1":  # 患者情報
                    kanaName=""
                    kanaNameSplit=row[2].split()
                    if len(kanaNameSplit)>=2:
                        kanaName = kanaNameSplit[0][0]+" "+kanaNameSplit[1][0]
                    elif len(kanaNameSplit)>=1:
                        kanaName = kanaNameSplit[0][0]
                    row[2]=kanaName
                    kanaNameFull=unicodedata.normalize("NFKC", kanaName)
                    kanaNameFull=kanaNameFull.replace(' ', '　')
                    row[3]=kanaNameFull
                    row[6]=""
                    row[7]=""
                    row[8]=""
                    row[9]=""
                    row[10]=""
                    row[11]=""
                    row[14]=""
                    row[15]=""
                    row[16]=""
                    row[17]=""
                    row[19]=""
                    row[20]=""
                    row[22]=""
                    row[23]=""
                    row[24]=""
                    row[25]=""
                    row[26]=""
                    row[27]=""
                    row[28]=""
                    row[29]=""
                    row[30]=""
                    row[31]=""
                    row[32]=""
                    row[33]=""
                    row[34]=""
                    row[35]=""
                    row[36]=""
                    row[48]=""

                if row[0] == "2":  # 処方箋情報
                    kanaName=""
                    if len(row[14])>0:
                        kanaName=row[14][0]
                    row[14]=kanaName
                    row[12]=""
                    row[13]=""
                    row[15]=""
                    row[16]=""
                    row[17]=""
                    row[18]=""
                    row[19]=""
                    row[20]=""
                    row[21]=""
                    row[22]=""
                    row[23]=""
                    row[24]=""
                    row[25]=""
                    row[26]=""
                    row[27]=""
                    row[28]=""
                    row[29]=""
                    row[30]=""
                    row[31]=""
                    row[32]=""

                sendCsv += ','.join(row)
                sendCsv += '\n'
            return sendCsv
        def onLoop(self):
            self.loopFlag = 1

        def offLoop(self):
            self.loopFlag = 0

    thread = ConcurrentlyWorker()
    def startThread(self):
        try:
            self.thread.onLoop()
            self.thread.start()  # スレッド開始
        except Exception as e:
            print(e)

    def stopThread(self):
        try:
            self.thread.offLoop()
        except Exception as e:
            print(e)
    def __init__(self):
        super(Window, self).__init__()

        global proxyIP, proxyPort, useInternetExplorerSetting, useProxy, IEProxyIP, IEProxyPort, proxyDict
        global fileType, fileTypeVer, fileFolder,privacyInfo
        global loginName, loginPass, isLogined, userID, userName, userEmail, MFtoken

        configSaver.load(self)
        self.sipsSettingWizardWindow= sipsSettingWizardWindow(2, 5,self)
        self.configWindow = configWindow(2, 5)
        self.connectWindow = connectWindow(2, 5)
        # subWindowを表示

        self.createIconGroupBox()
        self.createMessageGroupBox()
        self.creatDebugGroupBox()
        self.createSystemMessageGroupBox()
        self.createOnlineMessageGroupBox()
        self.createSettingGroupBox()

        self.sipsSettingWizardButton = QPushButton('レセコンのNSIPS(処方箋データ)との連動先フォルダ設定')
        self.sipsSettingWizardButton.setAutoDefault(False)
        self.connectButton = QPushButton('ログイン')
        self.fileTypeLabel.setMinimumWidth(self.durationLabel.sizeHint().width())
        self.createActions()
        self.createTrayIcon()

        self.sipsSettingWizardButton.clicked.connect(self.show_wizard_dialog)
        self.showMessageButton.clicked.connect(self.showMessage)
        self.showIconCheckBox.toggled.connect(self.trayIcon.setVisible)
        self.iconComboBox.currentIndexChanged.connect(self.changeFiletype)
        self.verComboBox.currentIndexChanged.connect(self.changeVertype)
        self.privacyComboBox.currentIndexChanged.connect(self.changePrivacytype)
        self.trayIcon.messageClicked.connect(self.messageClicked)
        self.trayIcon.activated.connect(self.iconActivated)
        self.btnFolder.clicked.connect(self.show_folder_dialog)
        self.connectButton.clicked.connect(self.show_connect_dialog)
        self.settingButton.clicked.connect(self.show_config_dialog)

        mainLayout = QVBoxLayout()
        mainLayout.addWidget(self.systemMessageGroupBox)
        mainLayout.addWidget(self.onlineMessageGroupBox)
        mainLayout.addWidget(self.iconGroupBox)
        mainLayout.addWidget(self.sipsSettingWizardButton)
        mainLayout.addWidget(self.messageGroupBox)
        mainLayout.addWidget(self.connectButton)
        mainLayout.addWidget(self.settingGroupBox)
        self.setLayout(mainLayout)
        self.setWindowTitle("MediMonitorレセプトコンピュータ連動ソフトウェア")
        self.resize(400, 300)
        self.thread.finSignal.connect(self.afterThreadFinished)
        self.startThread()
    # シグナルで送られたデータは引数として受け取れる
    def afterThreadFinished(self, signalData,messageDialog):
        global fileSuccessFlag
        print('thread is finished. signal: ', end='')
        print(str(signalData)+" ")
        #sleep(1)  # 1秒停止
        if signalData>0:
            objgraph.show_growth()  # メモリリークがあれば表示される
        if messageDialog==MESSAGE_NO_INDEX_FOLDER:

            print("MESSAGE_NO_INDEX_FOLDER")
            self.localMessageTextEdit.setText(MESSAGE_NO_INDEX_FOLDER_)
        elif messageDialog==MESSAGE_OTHER_FILE_DETECT:

            print("MESSAGE_OTHER_FILE_DETECT")
            self.localMessageTextEdit.setText(MESSAGE_OTHER_FILE_DETECT_)
            self.setVisible(True)

        elif messageDialog==MESSAGE_INDEX_EXIST_BUT_DATA_NOT_EXIST:
            print("MESSAGE_OTHER_FILE_DETECT")
            self.localMessageTextEdit.setText(MESSAGE_INDEX_EXIST_BUT_DATA_NOT_EXIST_)
            self.setVisible(True)
        elif messageDialog==MESSAGE_UPDATE_DETECT:

            print("MESSAGE_UPDATE_DETECT")
            self.localMessageTextEdit.setText("<b>エラー！</b>\u2029MediMonitor連動ソフトウェアをアップデートする必要があります。公式サイトより最新版よりダウンロードしてインストールしてください。")
            self.setVisible(True)
        elif messageDialog == MESSAGE_UPLOAD_TEXT_ERROR_DETECT:

            print("MESSAGE_UPLOAD_TEXT_ERROR_DETECT")
            self.localMessageTextEdit.setText("<b>エラー！</b>\u2029アップロードするテキストにエラーがあるためアップロード出来ませんでした。一度フォルダの中身を空にしてください。")

        elif messageDialog == MESSAGE_UPLOAD_FILE_TYPE_ERROR_DETECT:

            print("MESSAGE_UPLOAD_FILE_TYPE_ERROR_DETECT")
            self.localMessageTextEdit.setText("<b>エラー！</b>\u2029アップロードするテキストファイルの形式がNSIPS形式ではないためアップロード出来ませんでした。一度フォルダの中身を空にし、レセコンよりNSIPSを出力しているか確認して下さい。")

        elif messageDialog==MESSAGE_404_DETECT:

            print("MESSAGE_404_DETECT")
            self.localMessageTextEdit.setText("<b>エラー！</b>\u2029メンテナンス中の可能性があります。しばらくした後再接続します")
        elif messageDialog == MESSAGE_400_DETECT:

            print("MESSAGE_400_DETECT")
            self.localMessageTextEdit.setText("<b>エラー！</b>\u2029オンラインシステムからログアウトされました。ログインしてください。")

        elif messageDialog==MESSAGE_SEND_CSV:

            print("MESSAGE_SEND_CSV")
            self.localMessageTextEdit.setText("レセプトからデータを受信しました。オンラインシステムに連動用データを送信しています。")
            httpsRequest.viewRepaint(self)
        elif messageDialog==MESSAGE_NOT_FIND_TXT:
            print("MESSAGE_NOT_FIND_TXT")
            if isLogined:
                if fileSuccessFlag:
                    self.localMessageTextEdit.setText("オンラインシステムに接続され、システムは正常に動作しています。\u2029レセコンからの連動用ファイルの出力を待機しています")
                else:
                    self.localMessageTextEdit.setText("オンラインシステムには接続されていますが、レセコンから連動用ファイルがまだ受信されていません。\u2029"
                                                      "レセコンから連動用データを出力してください。\u2029もしデータを出力した場合でもこのメッセージが変わらない場合は他社の連動設定と間違っていないか、出力先のフォルダが間違っていないかなどをご確認ください")
            else :
                self.localMessageTextEdit.setText("オンラインシステムに接続されていません。ログインしてください。")
        elif messageDialog==SET_VISUALITY_FALSE:
            print("SET_VISUALITY_FALSE")
            httpsRequest.viewRepaint(self)
            self.setVisible(False)
        elif messageDialog==MESSAGE_INTERNET_ERROR:
            print("MESSAGE_INTERNET_ERROR")
            self.localMessageTextEdit.setText("<b>エラー！</b>\u2029インターネットに接続できませんでした。再試行しています。インターネットに接続されているかご確認ください。またプロキシーの設定が正しいかご確認ください。")



        else :
            print("viewRepaint")
            httpsRequest.viewRepaint(self)

    def setVisible(self, visible):
        self.minimizeAction.setEnabled(visible)
        self.restoreAction.setEnabled(self.isMaximized() or not visible)
        if not visible:
            self.trayIcon.showMessage(
            "MediMonitorはバックグラウンドで動作中です",
            "アプリケーションを終了するためにはシステムトレイを右クリックした後、 [終了]を "
                    "選択してください。",
            QSystemTrayIcon.Information,
            2000
        )
        super(Window, self).setVisible(visible)

    def closeEvent(self, event):
        if self.trayIcon.isVisible():
            self.hide()
            event.ignore()

    def changePrivacytype(self, index):
        global privacyInfo
        privacyInfo=index
        configSaver.saveConfig(self)
        if fileSuccessFlag:
            message = "すでに送信された患者情報を更新する\n→レセコンより上書きしたい患者を再送信\nすでに送信された患者情報を削除\n→レセコンより削除指示\nを行って下さい"
            ret = QMessageBox.information(None, "確認", message, QMessageBox.Ok)
    def changeVertype(self, index):
        global fileType,fileTypeVer
        fileTypeVer=index
        configSaver.saveConfig(self)

    def changeFiletype(self, index):
        self.setVerComboBox()
        global fileType,fileTypeVer
        fileType=index
        configSaver.saveConfig(self)

    def setVerComboBox(self):
        global fileType,fileTypeVer
        if self.iconComboBox.currentText()=="NSIPS":
            self.verComboBox.clear()
            self.verComboBox.addItem("バージョンの自動判別（基本的にはこちら）")
            self.verComboBox.addItem("Ver.1.04.01")
            self.verComboBox.addItem("ver 1.03")
            self.verComboBox.addItem("ver 1.02")
            self.verComboBox.addItem("ver 1.01")
            self.verComboBox.addItem("ver 1.00")

        if self.iconComboBox.currentText()=="JAHIS電子処方箋":
            self.verComboBox.clear()
            self.verComboBox.addItem("自動判別")
        if self.iconComboBox.currentText()=="JAHISお薬手帳CSV":
            self.verComboBox.clear()
            self.verComboBox.addItem("自動判別")

    def iconActivated(self, reason):
        if reason in (QSystemTrayIcon.Trigger, QSystemTrayIcon.DoubleClick):
            self.setVisible(True)
        elif reason == QSystemTrayIcon.MiddleClick:
            self.showMessage()

    def showMessage(self):
        icon = QSystemTrayIcon.MessageIcon(
                self.typeComboBox.itemData(self.typeComboBox.currentIndex()))
        self.trayIcon.showMessage(self.pharmIDEdit.text(),
                self.bodyEdit.toPlainText(), icon,
                self.durationSpinBox.value() * 1000)

    def messageClicked(self):
        self.setVisible(True)

    def createSystemMessageGroupBox(self):

        self.systemMessageGroupBox = QGroupBox("メッセージ")
        onlineMessageLayout = QGridLayout()
        self.localMessageTextEdit = QTextEdit()
        self.localMessageTextEdit.setText("オンラインシステムに接続されていません。ログインしてください。")
        self.localMessageTextEdit.setReadOnly(True)
        onlineMessageLayout.addWidget( self.localMessageTextEdit,2,0)
        self.systemMessageGroupBox.setLayout(onlineMessageLayout)

    def createOnlineMessageGroupBox(self):
        self.onlineMessageGroupBox = QGroupBox("オンライン状況")
        onlineMessageLayout = QGridLayout()
        self.onlineMessageLabel = QLabel("<b>ログアウト</b>")

        self.localMessageLabel = QLabel("連動状況")
        onlineMessageLayout.addWidget( self.onlineMessageLabel,0,0)
        self.onlineMessageGroupBox.setLayout(onlineMessageLayout)

    def createSettingGroupBox(self):
        self.settingGroupBox = QGroupBox("")
        iconLayout = QGridLayout()
        typeLabel = QLabel("ログインできない場合はこちらでプロキシを設定してください")
        self.settingButton = QPushButton('詳細設定')
        iconLayout.addWidget(typeLabel,0,0)
        iconLayout.addWidget(self.settingButton,0,4)
        self.settingGroupBox.setLayout(iconLayout)

    def createIconGroupBox(self):
        global fileType,fileTypeVer,fileFolder,privacyInfo
        self.iconGroupBox = QGroupBox("")

        self.fileTypeLabel = QLabel("ファイル形式:")
        self.verLabel = QLabel("バージョン:")
        self.privacyLabel = QLabel("プライバシー:")
        self.privacyComboBox = QComboBox()
        self.privacyComboBox.addItem("患者の個人情報を送信する")
        self.privacyComboBox.addItem("患者の個人情報を送信しない")
        self.iconComboBox = QComboBox()
        self.iconComboBox.addItem("NSIPS")
        self.verComboBox = QComboBox()
        self.iconComboBox.setCurrentIndex(int(fileType))
        self.setVerComboBox()
        self.verComboBox.setCurrentIndex(int(fileTypeVer))
        self.privacyComboBox.setCurrentIndex(int(privacyInfo))
        self.showIconCheckBox = QCheckBox("Show icon")
        self.showIconCheckBox.setChecked(True)
        self.txtLabel = QLabel("連動先フォルダ:")
        self.folderLabel = QLabel("注意：NsipsはDATAとINDEXが存在しているフォルダを選択してください\n【OK】 C:\\SIPS1 \n【NG】 C:\\SIPS1\\DATA")
        self.txtFolder = QLineEdit()
        self.txtFolder.setText(fileFolder)
        self.txtFolder.setReadOnly(True)
        self.btnFolder = QPushButton('参照...')

        iconLayout = QGridLayout()
        iconLayout.addWidget(self.fileTypeLabel, 2, 0)
        iconLayout.addWidget(self.iconComboBox, 2, 1, 1, 4)
        iconLayout.addWidget(self.verLabel, 3, 0)
        iconLayout.addWidget(self.verComboBox, 3, 1, 1, 4)
        iconLayout.addWidget(self.privacyLabel, 4, 0)
        iconLayout.addWidget(self.privacyComboBox, 4, 1, 1, 4)
        iconLayout.addWidget(self.txtLabel, 5, 0)
        iconLayout.addWidget(self.txtFolder, 5, 1, 1, 4)
        self.iconGroupBox.setLayout(iconLayout)

    def creatDebugGroupBox(self):
        self.debugGroupBox = QGroupBox("状況")
        bodyLabel = QLabel("本体:")
        debugLayout = QGridLayout()
        debugLayout.addWidget(bodyLabel)
        self.debugGroupBox.setLayout(debugLayout)

    def createMessageGroupBox(self):
        self.messageGroupBox = QGroupBox("")
        self.typeComboBox = QComboBox()
        self.typeComboBox.addItem("None", QSystemTrayIcon.NoIcon)
        self.typeComboBox.addItem(self.style().standardIcon(
                QStyle.SP_MessageBoxInformation), "情報",
                QSystemTrayIcon.Information)
        self.typeComboBox.addItem(self.style().standardIcon(
                QStyle.SP_MessageBoxWarning), "警告",
                QSystemTrayIcon.Warning)
        self.typeComboBox.addItem(self.style().standardIcon(
                QStyle.SP_MessageBoxCritical), "重要",
                QSystemTrayIcon.Critical)
        self.typeComboBox.setCurrentIndex(1)

        self.durationLabel = QLabel("間隔:")
        self.btnConnect = QPushButton('ログイン')
        self.durationSpinBox = QSpinBox()
        self.durationSpinBox.setRange(5, 60)
        self.durationSpinBox.setSuffix(" s")
        self.durationSpinBox.setValue(15)
        durationWarningLabel = QLabel("(ヒント)")
        durationWarningLabel.setIndent(10)
        self.btnSetting = QPushButton('詳細設定')
        titleLabel = QLabel("薬局ID:")
        self.pharmIDEdit = QLineEdit()
        self.onlyInt = QIntValidator()
        self.pharmIDEdit.setValidator(self.onlyInt)
        self.pharmIDEdit.setText(userID)
        passwordLabel = QLabel("パスワード:")
        self.passwordEdit = QLineEdit("")
        self.passwordEdit.setEchoMode(QLineEdit.Password)
        self.bodyEdit = QTextEdit()
        self.bodyEdit.setPlainText("ボディーコンテンツ")
        self.showMessageButton = QPushButton("メッセージを表示")
        self.showMessageButton.setDefault(True)
        messageLayout = QGridLayout()
        messageLayout.addWidget(titleLabel, 2, 0)
        messageLayout.addWidget(self.pharmIDEdit, 2, 1, 1, 4)
        messageLayout.addWidget(passwordLabel, 3, 0)
        messageLayout.addWidget(self.passwordEdit, 3, 1, 1, 4)
        self.messageGroupBox.setLayout(messageLayout)

    def quitInfo(self):
        ret = QMessageBox.information(None, "確認", "アプリケーションを終了してよろしいですか？終了した場合次回起動するまでデータを連動することができません。",  QMessageBox.Yes, QMessageBox.No)
        if ret == QMessageBox.Yes:
            QApplication.instance().quit()
    def showVerInfo(self):
        message = "MediMonitor連動プログラム\nCopyright © MedicalFields株式会社\n" \
                  "バージョン情報:" +VERSION_INFO+\
                  "\nこのプログラムは日本薬剤師会が提案する薬局向けコンピュータシステム間の連携システムであるNSIPSのライセンスを受けております\n" \
                  "※NSIPSは公益社団法人日本薬剤師会の登録商標です"
        ret = QMessageBox.information(None, "バージョン情報", message, QMessageBox.Ok)
    def createActions(self):
        self.minimizeAction = QAction("&最小化", self, triggered=self.hide)
        self.infoAction = QAction("&バージョン情報", self,
                triggered=self.showVerInfo)
        self.restoreAction = QAction("&表示", self,
                triggered=self.showNormal)
        self.quitAction = QAction("&終了", self,
                triggered=self.quitInfo)

    def createTrayIcon(self):
         self.trayIconMenu = QMenu(self)
         self.trayIconMenu.addAction(self.minimizeAction)
         self.trayIconMenu.addAction(self.restoreAction)
         self.trayIconMenu.addAction(self.infoAction)
         self.trayIconMenu.addSeparator()
         self.trayIconMenu.addAction(self.quitAction)
         self.trayIcon = QSystemTrayIcon(self)
         self.trayIcon.setContextMenu(self.trayIconMenu)
         self.setIcon(0)
         self.trayIcon.show()

    def show_folder_dialog(self):
        global fileFolderTemp,fileFolder

        fileFolderTempt=fileFolderTemp
        if fileFolderTempt == "":
            fileFolderTempt = fileFolder

            if fileFolderTempt == "":
                fileFolderTempt = "C:\\"

        message = "フォルダを選択する場合はレセコンより、連動用ファイルを出力してから選択してください。\n連動用ファイルを１度も出力していない場合、連動先フォルダが作成されていない可能性があります。"
        ret = QMessageBox.information(None, "注意", message, QMessageBox.Ok)

        dirname = QFileDialog.getExistingDirectory(self,
                                                   'SIPSがあるフォルダを選択してください',fileFolderTempt, QFileDialog.ShowDirsOnly)
        if dirname:
            self.dirname = dirname.replace('/', os.sep)
            if self.dirname.endswith('DATA') or self.dirname.endswith('data') or self.dirname.endswith('INDEX') or self.dirname.endswith('index') or self.dirname.endswith('Index') or self.dirname.endswith('Data'):
                message= "Nsipsのフォルダの選択の場合、DATAやINDEXがあるフォルダを選択する必要があります。\nDATAまたはINDEXのある直前のフォルダを選択してもよろしいですか？わからない場合は「Yes」を選択してください。"
                ret = QMessageBox.warning(None,"確認",message, QMessageBox.Yes, QMessageBox.No)

                if ret == QMessageBox.Yes:
                    print('Yes が クリックされました')
                    if self.dirname.endswith('DATA')  or  self.dirname.endswith('data')  or  self.dirname.endswith('Data') :
                        self.dirname=self.dirname[:-5]
                    if  self.dirname.endswith('INDEX')  or  self.dirname.endswith('index') or  self.dirname.endswith('Index'):
                        self.dirname=self.dirname[:-6]

            showOtherFolderErrorMessage=False
            files = glob.glob(self.dirname+"\\" + "*")
            for file in files:
                print("show_folder_dialog")
                if file.endswith('DATA') or file.endswith('data')  or file.endswith('Data')  or file.endswith('INDEX')  or file.endswith('index')  or file.endswith('Index') :
                    print(file)
                else:
                    print(file)
                    showOtherFolderErrorMessage=True
            if showOtherFolderErrorMessage:
                message = "現在選択したフォルダには違う名前のファイルまたはフォルダが含まれています。通常NsipsのフォルダにはDATAとINDEXのフォルダのみしか存在しないはずです。このフォルダで続行してもよろしいですか？"
                ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes, QMessageBox.No)

                if ret == QMessageBox.No:
                    return
            fileFolderTemp=self.dirname

            try:
                self.txtFolderW.setText(self.dirname)
            except AttributeError:
                print("AttributeError")

            self.step = 0
    def logoutProcess(self,doPasswordClear=True):
        global isLogined,MFtoken,userEmail,userName,userID
        MFtoken = ""
        userID = ""
        userName = ""
        userEmail = ""
        isLogined = False
        self.onlineMessageLabel.setText("<b>ログアウト</b>")
        self.connectButton.setText("ログイン")
        self.localMessageTextEdit.setText("オンラインシステムに接続されていません。ログインしてください。")
        self.passwordEdit.setEnabled(True)
        if doPasswordClear:
            self.passwordEdit.setText("")
        self.pharmIDEdit.setEnabled(True)
        self.connectButton.setEnabled(True)
        self.repaint()
        configSaver.saveConfig(self)
        self.stopThread()
    def show_connect_dialog(self):
        global isLogined,MFtoken,userEmail,userName,userID
        if isLogined :
            message = "ログアウトしてよろしいですか？"
            ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes, QMessageBox.No)

            if ret == QMessageBox.Yes:
                print('Yes が クリックされました')
                self.logoutProcess()
        else:
            if self.txtFolder.text() == "":
                message = "「レセコンのNSIPS(処方箋データ)との連動先フォルダ設定」をクリックし、初期設定を行ってください"
                ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                return
            if self.pharmIDEdit.text() == "":
                message = "IDを入力してください。"
                ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                return
            if self.passwordEdit.text() == "":
                message = "パスワードを入力してください。"
                ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                return
            httpsRequest.loginRequest(self, self.pharmIDEdit.text(), self.passwordEdit.text(), "pc")

    def show_wizard_dialog(self):

        global fileFolder
        if fileFolder !="":
            ret = QMessageBox.information(None, "確認", "既にNSIPSのフォルダが指定されています。もう一度再設定しますか？",
                                          QMessageBox.Yes, QMessageBox.No)
            if ret == QMessageBox.Yes:
                self.sipsSettingWizardWindow.show()
        else:
            self.sipsSettingWizardWindow.show()

    def show_config_dialog(self):
        configSaver.saveConfig(self)
        self.configWindow.show()
    def setIcon(self,normal0_connect1_bad2_error3):
        if normal0_connect1_bad2_error3 == 0:
            self.trayIcon.setIcon(QIcon(":/images/icon.png"))
            self.setWindowIcon(QIcon(':/images/icon.png'))
        elif normal0_connect1_bad2_error3 == 1:
            self.trayIcon.setIcon(QIcon(":/images/connect.png"))
            self.setWindowIcon(QIcon(':/images/connect.png'))
        elif normal0_connect1_bad2_error3 == 2:
            self.trayIcon.setIcon(QIcon(":/images/bad.png"))
            self.setWindowIcon(QIcon(':/images/bad.png'))
        elif normal0_connect1_bad2_error3 == 3:
            self.trayIcon.setIcon(QIcon(":/images/error.png"))
            self.setWindowIcon(QIcon(':/images/error.png'))

class MagicWizard(QtWidgets.QWizard):
    def __init__(self, parent=None):
        super(MagicWizard, self).__init__(parent)
        self.parents=parent
        self.addPage(Page1(self))
        self.addPage(Page2(self))

        self.setWindowTitle("NSIPS(処方箋データ)とMediMonitorの連動設定ウィザード")
        self.setButtonText(QtWidgets.QWizard.BackButton,"戻る")
        self.setButtonText(QtWidgets.QWizard.NextButton,"次へ")
        self.setButtonText(QtWidgets.QWizard.CancelButton,"キャンセル")
        self.setButtonText(QtWidgets.QWizard.FinishButton,"連動先フォルダをチェックしてください")

        self.button(QtWidgets.QWizard.NextButton).clicked.connect(self._checkButton)
        self.button(QtWidgets.QWizard.FinishButton).clicked.connect(self._doSomething)
        self.resize(640,480)

    def _doSomething(self):
        global fileFolder,fileFolderCheckType
        if fileFolderCheckType==0:

            ret = QMessageBox.information(None, "確認", "フォルダーチェックに成功していません。本当に登録してもよろしいですか？",
                                          QMessageBox.Yes, QMessageBox.No)
            if ret == QMessageBox.Yes:
                print("");

            if ret == QMessageBox.No:
                return

        t=self.parents.txtFolderW.text()
        fileFolder=t
        self.parents.txtFolder.setText(t)
        configSaver.saveConfig(self.parents)

    def _checkButton(self):
        global fileFolder
        self.button(QtWidgets.QWizard.FinishButton).setEnabled(False)


class Page1(QtWidgets.QWizardPage):
    def openURL(self,url):
        import webbrowser
        webbrowser.open(url)
    def __init__(self, parent=None):
        super(Page1, self).__init__(parent)
        layout = QtWidgets.QVBoxLayout()

        messageGroupBox = QGroupBox("レセプトメーカー一覧")
        messageLayout = QGridLayout()
        addressLabel = QLabel("レセプトメーカーによってNSIPSの設定は異なります。\n"
                              "以下のメーカーをクリックすると説明書が開きます。\n"
                              "レセコン側でNSIPSの連動設定を行った後、次へをクリックしてください。")

        self.receiptFolder = QLineEdit()
        self.addressLabel = QLabel("フォルダ:")
        pybutton   = QPushButton('NO@H FOR THE PHARMACY【ノアメディカルシステム】', self)
        pybutton2  = QPushButton('ReceptyNEXT【EMシステムズ】', self)
        pybutton3  = QPushButton('ぶんぎょうめいと【EMシステムズ】', self)
        pybutton4  = QPushButton('Apobahn【EMシステムズ】', self)
        pybutton5  = QPushButton('Pharma-SEED【日立ヘルスケアシステムズ】', self)
        pybutton6  = QPushButton('調剤Melphin【三菱電機】', self)
        pybutton7  = QPushButton('メディコム ファーネス【ＰＨＣ】', self)
        pybutton8  = QPushButton('エリシア【シグマソリューションズ】', self)
        pybutton9  = QPushButton('P-CUBE【ユニケソフトウェアリサーチ】', self)
        pybutton10 = QPushButton('Pharmacy Aid【ハイブリッジ】', self)
        pybutton11 = QPushButton('MEDI-ECHO P【ICソリューションズ】', self)
        pybutton12 = QPushButton('調剤王Ⅲ・調剤王ⅢLite【エムウィンソフト】', self)
        pybutton13 = QPushButton('調剤くん【ネグジット総研】', self)
        pybutton14 = QPushButton('Pharao 7Edition【両毛システムズ】', self)
        pybutton15 = QPushButton('調剤用DRS【メディカルアイリード】', self)
        pybutton16 = QPushButton('PRESUS【メディセオ】', self)
        pybutton17 = QPushButton('Gennai7（ゲンナイセブン）【ズー】', self)
        pybutton18 = QPushButton('ファーミー【モイネットシステム】', self)
        pybutton19 = QPushButton('GOHL2, OKISS【シンク】', self)
        pybutton20 = QPushButton('IpharmaⅡ【インフォテクノ】', self)
        pybutton21 = QPushButton('調剤在庫システム PAWAFULⅢ【ワンズ・システム】', self)
        pybutton22 = QPushButton('保険薬局システム Pharm-i【アイテックソフトウェア】', self)
        pybutton23 = QPushButton('Premacy Neo【コラソンシステムズ】', self)
        pybutton24 = QPushButton('調剤支援【たちばな薬局】', self)
        pybutton25 = QPushButton('その他', self)
        pybutton.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_noah/"))
        pybutton2.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_recepty_next/"))
        pybutton3.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_bungyo/"))
        pybutton4.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_apobahn/"))
        pybutton5.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pharma_seed/"))
        pybutton6.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_melphin/"))
        pybutton7.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pharnes/"))
        pybutton8.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_elixir/"))
        pybutton9.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_p_cube/"))
        pybutton10.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pharmacy_aid/"))
        pybutton11.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_medi_echo/"))
        pybutton12.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_tyouzai_ou/"))
        pybutton13.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_tyouzai_kun/"))
        pybutton14.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pharao/"))
        pybutton15.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_drs/"))
        pybutton16.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_presus/"))
        pybutton17.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_gennai7/"))
        pybutton18.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pharmy/"))
        pybutton19.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_gohl2/"))
        pybutton20.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_ipharma2/"))
        pybutton21.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pawaful3/"))
        pybutton22.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_pharm_i/"))
        pybutton23.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_premacy_neo/"))
        pybutton24.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_tatibana/"))
        pybutton25.clicked.connect(partial(self.openURL,"https://medicalfields.jp/how_to_setup_other/"))

        messageLayout.addWidget(pybutton,1, 0)
        messageLayout.addWidget(pybutton2,1, 1)
        messageLayout.addWidget(pybutton3,2, 0)
        messageLayout.addWidget(pybutton4,2, 1)
        messageLayout.addWidget(pybutton5,3, 0)
        messageLayout.addWidget(pybutton6,3, 1)
        messageLayout.addWidget(pybutton7,4, 0)
        messageLayout.addWidget(pybutton8,4, 1)
        messageLayout.addWidget(pybutton9,5, 0)
        messageLayout.addWidget(pybutton10,5, 1)
        messageLayout.addWidget(pybutton11,6, 0)
        messageLayout.addWidget(pybutton12,6, 1)
        messageLayout.addWidget(pybutton13,7, 0)
        messageLayout.addWidget(pybutton14,7, 1)
        messageLayout.addWidget(pybutton15,8, 0)
        messageLayout.addWidget(pybutton16,8, 1)
        messageLayout.addWidget(pybutton17,9, 0)
        messageLayout.addWidget(pybutton18,9, 1)
        messageLayout.addWidget(pybutton19,10, 0)
        messageLayout.addWidget(pybutton20,10, 1)
        messageLayout.addWidget(pybutton21,11, 0)
        messageLayout.addWidget(pybutton22,11, 1)
        messageLayout.addWidget(pybutton23,12, 0)
        messageLayout.addWidget(pybutton24,12, 1)
        messageLayout.addWidget(pybutton25,13, 0)
        messageGroupBox.setLayout(messageLayout)
        layout.addWidget(addressLabel)
        layout.addWidget(messageGroupBox)
        # layout.addWidget(pybutton)
        # layout.addWidget(pybutton2)
        # layout.addWidget(pybutton3)
        # layout.addWidget(pybutton4)
        # layout.addWidget(pybutton5)
        # layout.addWidget(pybutton6)
        # layout.addWidget(pybutton7)
        # layout.addWidget(pybutton8)
        # layout.addWidget(pybutton9)
        # layout.addWidget(pybutton10)
        # layout.addWidget(pybutton11)
        # layout.addWidget(pybutton12)
        # layout.addWidget(pybutton13)
        # layout.addWidget(pybutton14)
        # layout.addWidget(pybutton15)
        # layout.addWidget(pybutton16)
        # layout.addWidget(pybutton17)
        # layout.addWidget(pybutton18)
        self.setLayout(layout)


class Page2(QtWidgets.QWizardPage):
    def __init__(self, parent=None):
        super(Page2, self).__init__(parent)
        self.parents=parent.parents
        global fileFolder
        layout = QtWidgets.QVBoxLayout()

        messageGroupBox = QGroupBox("")
        messageLayout = QGridLayout()
        addressLabel = QLabel("前項目でレセコンに設定したNSIPSの出力先フォルダを入力または参照し,「連動先フォルダをチェックする」を押してください")

        txtLabel = QLabel("連動先フォルダ:")
        self.parents.txtFolderW = QLineEdit()
        self.parents.txtFolderW.setText(fileFolder)
        self.parents.txtFolderW.setReadOnly(False)
        self.parents.txtFolderW.textChanged[str].connect(partial(self.disableFinishButton,parent))
        btnFolder = QPushButton('参照...')

        btnFolder.clicked.connect(partial(Window.show_folder_dialog,self.parents))
        btnFolderCheck = QPushButton('連動先フォルダをチェックする')

        btnFolderCheck.clicked.connect(partial(self._doChecked,parent))
        messageLayout.addWidget(txtLabel, 2, 0)
        messageLayout.addWidget(self.parents.txtFolderW, 2, 1, 1, 3)
        messageLayout.addWidget(btnFolder, 2, 4, 1, 1)

        messageLayout.addWidget(QLabel("※半角の【＼】と【￥】は同じ文字として扱われます 入力をする場合は半角で【\】と入力してください"),3, 1, 1, 4)
        messageLayout.addWidget(QLabel("例: ¥¥emscr01¥RECEPTYN¥SIPS3"),4, 1, 1, 4)
        messageLayout.addWidget(QLabel("例: C:\SIPS2"), 5, 1, 1, 4)
        messageGroupBox.setLayout(messageLayout)
        layout.addWidget(addressLabel)
        layout.addWidget(messageGroupBox)
        layout.addWidget(btnFolderCheck)

        systemMessageGroupBox = QGroupBox("")

        onlineMessageLayout = QGridLayout()
        self.parents.localFolderMessageTextEdit = QTextEdit()
        self.parents.localFolderMessageTextEdit.setText("レセプトコンピューターにて設定したNsips（Sips）を出力するフォルダを「連動先フォルダ」に直接入力（右クリックでコピーしてPaste（貼り付ける）こともできます）または参照をクリックしてフォルダーを"
                                                "特定してください\n"
                                                "※フォルダを参照して選択する場合は、一度レセコンより連動用ファイルを出力させて下さい"
                                                "（一度も出力していない場合は出力先のフォルダが作成されていない場合があります）\n\n"
                                                "連動先フォルダを入力または選択が終わりましたら「連動先フォルダをチェックする」をクリックして指定したフォルダに連動可能なNSIPSがあるかチェックしてください")
        self.parents.localFolderMessageTextEdit.setReadOnly(True)
        onlineMessageLayout.addWidget( self.parents.localFolderMessageTextEdit,2,0)
        systemMessageGroupBox.setLayout(onlineMessageLayout)
        layout.addWidget(systemMessageGroupBox)
        self.setLayout(layout)

    def disableFinishButton(self,parent):
        self.setButtonText(QtWidgets.QWizard.FinishButton, "連動先フォルダをチェックしてください")
        parent.button(QtWidgets.QWizard.FinishButton).setEnabled(False)
    def _doChecked(self,parent):

        newFolderPath=self.parents.txtFolderW.text()
        if len(newFolderPath)==0:
            message="<b>エラー！</b>\u2029連動先フォルダに入力がされていません。連動先フォルダを直接入力（右クリックでコピーしてPaste（貼り付ける）こともできます）または参照よりフォルダ選択して下さい"
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
            return

        if (re.match(r".+\\\\.*", newFolderPath) ):
            message = "<b>エラー！</b>\u2029連動先フォルダに入力されているパスがフォルダパスとして正しくありません。【\】は1つだけで入力して下さい。\u2029【☓ \\\\】\u2029【○ \\】\u2029\u2029¥¥emscr01¥RECEPTYN¥SIPS3\u2029C:\\SIPS\u2029のように入力してください\n" \
                      ""
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
            return
        if not (re.match(FOLDER_PATH_REGEX, newFolderPath) or re.match(NETWORK_FOLDER_PATH_REGEX, newFolderPath) ):
            message = "<b>エラー！</b>\u2029連動先フォルダに入力されているパスがフォルダパスとして正しくありません。空白が間違って入ってないか【:】【\】が間違って入力されていないか 【\】は半角で入力されているかご確認下さい。また【＼】と【/】は【\】として入力してください\u2029\u2029¥¥emscr01¥RECEPTYN¥SIPS3\u2029C:\\SIPS\u2029のように入力してください\n" \
                      ""
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
            return

        f=0
        newParentFolderPath=str('\\'.join(os.path.abspath(newFolderPath).split('\\')[0:-1 - f]))
        if not os.path.exists(newParentFolderPath):
            newParentFolderPathCount=newParentFolderPath.count('\\')
            print (newParentFolderPathCount)
            if (newParentFolderPathCount>2):
                print(newParentFolderPath)
                message = "<b>エラー！</b>\u2029親フォルダが存在しないため、入力された連動先フォルダのパスは間違っています。\u2029" \
                          "各メーカーのNSIPSの設定画面で指定したフォルダが正しく入力または選択されているかご確認下さい。間違った箇所に空白が入っていたりすると動作しません。多くの場合、連動用フォルダは【sips1】や【nsips1】などの文字で終わることが多いです。" \
                          "\u2029親フォルダ：" + newParentFolderPath
                self.parents.localFolderMessageTextEdit.setText(message)
                ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
                return


        if not os.path.exists(newFolderPath):
            message =MESSAGE_NO_MAIN_FOLDER_
            self.parents.localFolderMessageTextEdit.setText(message)

            box = QMessageBox()
            box.setIcon(QMessageBox.Warning)
            box.setWindowTitle('確認')
            box.setText(message)
            box.setStandardButtons(QMessageBox.Yes | QMessageBox.No)
            buttonY = box.button(QMessageBox.Yes)
            buttonY.setText('フォルダを入力した場所に新規作成します(レセコンで連動用データが出力できなく,連動先フォルダが確定してる時のみ)')
            buttonN = box.button(QMessageBox.No)
            buttonN.setText('キャンセル')
            box.exec_()

            if box.clickedButton() == buttonY:
                try:
                    os.mkdir(newFolderPath)
                except:
                    print("An mkdir exception occurred")
                if os.path.exists(newFolderPath):
                    ret = QMessageBox.information(None, "確認", "フォルダーを"+str(newFolderPath)+"に新規作成しました。再びフォルダのチェックを行ってください", QMessageBox.Yes)
                else:
                    ret = QMessageBox.warning(None, "確認", "フォルダー："+str(newFolderPath)+"の作成に失敗しました。\u2029"
                                                                                      "連動先に入力されたデータが間違っています。もう一度入力内容をよくご確認下さい。", QMessageBox.Yes)
            return


        if not self.parents.folderErrorCheckTemp(newFolderPath):
            message =MESSAGE_NO_INDEX_FOLDER_ + "\u2029現在レセコンより連動用のデータが出力できず、連動先フォルダの場所が入力した値で確定してる時は画面下にある「連動先のフォルダにデータはないが強制的に登録する」をクリックすると、将来的にフォルダに連動データが転送された際に自動的に連動されます"
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)

            self.setButtonText(QtWidgets.QWizard.FinishButton, "連動先のフォルダにデータはないが強制的に登録する")
            parent.button(QtWidgets.QWizard.FinishButton).setEnabled(True)
            return

        files = glob.glob(indexDirTemp + "*")
        fileSuccessFlag = False
        otherFileDetect = False
        indexExistButDataNotExist = False
        for file in files:
            if file.endswith("txt") or file.endswith("TXT"):
                basename = os.path.basename(file)
                dataFilePath = dataDirTemp + basename
                csvBase64 = ""
                if os.path.exists(dataFilePath):
                    fileSuccessFlag=True

                else:
                    indexExistButDataNotExist = True

                print(dataFilePath)
            else:
                print(file)
                otherFileDetect = True
        if otherFileDetect:
            message ="<b>エラー！</b>\u2029Nsipsの連動設定がされているindexフォルダに別のファイルが見つかりました。\u2029通常indexフォルダにはtxtファイルのみしか存在できません。\u2029" \
                           "一度フォルダ"+newFolderPath+"のINDEXとDATAフォルダの中身を空にしてください。また入力された連動先フォルダが間違っている可能性もあります。入力内容をもう一度ご確認下さい。"
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
            return

        if indexExistButDataNotExist:
            message ="<b>エラー！</b>\u2029Nsipsの連動設定がされている連動先のindexフォルダに存在するファイルとdataフォルダに存在するファイルが一致しません。連動データが破損している可能性が高いです。\u2029" \
                           "一度フォルダ"+newFolderPath+"のINDEXとDATAフォルダの中身を空にしてください。また入力された連動先フォルダが間違っている可能性もあります。入力内容をもう一度ご確認下さい。"
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
            return

        if fileSuccessFlag:
            message ="<b>成功</b>\u2029Nsipsの連動設定が成功しました。完了をクリックして連動先フォルダの登録を行ってください"
            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)

            self.setButtonText(QtWidgets.QWizard.FinishButton, "完了")
            parent.button(QtWidgets.QWizard.FinishButton).setEnabled(True)
            global fileFolderCheckType
            fileFolderCheckType = 1
            return
        else :
            message ="<b>注意</b>\u2029Nsipsの連動設定が成功してると思われますが、連動用の処方箋データ（NSIPSデータ）は見つかりませんでした。\u2029" \
                     "連動用データを一度出力してもう一度「連動先フォルダをチェック」を行ってください。\u2029" \
                     "連動用ファイルを出力している場合でこの画面が何度も出る場合は、入力された連動用フォルダが他のNSIPS連動メーカーの連動先フォルダに設定されている可能性があります。\u2029" \
                     "連動先フォルダのフォルダ名に数字などがある場合、新しく出力される連動先の正しい数字が入力されているかを確認してください。\u2029" \
                     "またメーカによっていは特定の操作を行わないとNSIPSの連動用データが出力されない場合があります。\u2029" \
                     "各レセコンメーカにNSIPSの出力についてお問い合わせ下さい。\u2029" \
                     "もし既に連動設定が完了している場合もこの画面が出ることがあります。アプリをご確認し既にデータが転送されていないかご確認下さい" \
                     "\u2029現在レセコンより連動用のデータが出力できず、連動先フォルダの場所が入力した値で確定してる時は画面下にある「連動先のフォルダにデータはないが強制的に登録する」をクリックすると、将来的にフォルダに連動データが転送された際に自動的に連動されます\u2029"

            self.parents.localFolderMessageTextEdit.setText(message)
            ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)

            self.setButtonText(QtWidgets.QWizard.FinishButton, "連動先のフォルダにデータはないが強制的に登録する")
            parent.button(QtWidgets.QWizard.FinishButton).setEnabled(True)
            return

    def initializePage(self):
        print("initializePage");

class Page3(QtWidgets.QWizardPage):
    def __init__(self, parent=None):
        super(Page3, self).__init__(parent)
        layout = QtWidgets.QVBoxLayout()
        self.receiptFolder = QLineEdit()
        layout.addWidget(self.receiptFolder)
        self.setLayout(layout)

class sipsSettingWizardWindow:

     def __init__(self, x, y, parent=None):
         print ("init");
         self.parent_=parent


     def show(self):
         self.w = MagicWizard(self.parent_)
         self.w.exec_()

class connectWindow:
     def __init__(self, x, y, parent=None):
         # QDialogのインスタンス
         self.w = QDialog(parent)
         self.messageGroupBox = QGroupBox("ログイン")
         self.x = x
         self.y = y
         self.z = self.x * self.y
         self.sub_label1 = QLabel()
         self.sub_label1.setText(str(self.x) + "×" + str(self.y) + "=" + str(self.z))
         messageLayout = QGridLayout()

         bodyEdit = QTextEdit()
         bodyEdit.setPlainText("詳細なメッセージが必要ならばこちらを選択して下さい")
         addressLabel = QLabel("メッセージ:")
         messageLayout.addWidget(addressLabel,0,0)
         messageLayout.addWidget(bodyEdit,1,0)
         self.w.setLayout(messageLayout)


     def show(self):
         self.w.exec_()

class configWindow:
    def __init__(self,x,y,parent=None):
        #QDialogのインスタンス
        self.w = QDialog(parent)
        self.w.setWindowTitle("詳細設定")
        self.messageGroupBox = QGroupBox("プロキシサーバー")
        self.x = x
        self.y = y
        self.z = self.x*self.y
        self.sub_label1 = QLabel()
        self.sub_label1.setText(str(self.x)+"×"+str(self.y)+"="+str(self.z))

        #レイアウトを設定する
        self.proxyComboBox = QComboBox()
        self.proxyComboBox.addItem("InternetExplorerの設定を使用する")
        self.proxyComboBox.addItem("自分で設定する")

        self.proxyComboBox.currentIndexChanged.connect(self.changeProxyType)
        self.useProxyCheckBox = QCheckBox("Proxyを使用する")
        messageLayout = QGridLayout()
        addressLabel = QLabel("アドレス:")

        self.useProxyCheckBox.stateChanged.connect(self.proxyCheckBoxOnChanged)

        self.ipEdit = QLineEdit()
        portLabel = QLabel("ポート番号:")
        validator = QIntValidator(0, 65536)

        self.proxyPortEdit = QLineEdit()
        self.proxyIPEdit = QLineEdit()
        self.proxyPortEdit.setValidator(validator)
        messageLayout.addWidget(self.proxyComboBox,0, 0, 1, 5)
        messageLayout.addWidget(self.useProxyCheckBox,1, 0, 1, 5)
        messageLayout.addWidget(addressLabel, 2, 0)
        messageLayout.addWidget(self.proxyIPEdit, 2, 1, 1, 4)
        messageLayout.addWidget(portLabel, 3, 0)
        messageLayout.addWidget(self.proxyPortEdit, 3, 1, 1, 4)

        layout = QVBoxLayout()
        layout.addWidget(addressLabel)
        layout.addWidget(self.proxyIPEdit)
        layout.addWidget(portLabel)
        layout.addWidget(self.proxyPortEdit)
        mainLayout = QVBoxLayout()
        self.messageGroupBox.setLayout(messageLayout)
        mainLayout.addWidget(self.messageGroupBox)

        self.saveConfigButton = QPushButton('保存')
        self.saveConfigButton.clicked.connect(self.saveConfig)
        mainLayout.addWidget(self.saveConfigButton)
        self.w.setLayout(mainLayout)
        self.changeProxyType(self.proxyComboBox.currentIndex)

    def show(self):
        self.loadConfig()
        self.w.exec_()

    def saveConfig(self):
        global useInternetExplorerSetting, useProxy,proxyPort,proxyIP
        if self.useProxyCheckBox.isChecked():
            if self.proxyIPEdit.text() == "":
                message = "ホスト名を入力してください。"
                ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                return

            if self.proxyPortEdit.text() == "":
                message = "ポート番号を入力してください。"
                ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                return

        if self.proxyComboBox.currentIndex()==0:
            useInternetExplorerSetting = True
        else:
            useInternetExplorerSetting = False

        useProxy = self.useProxyCheckBox.isChecked()
        proxyPort= self.proxyPortEdit.text()
        proxyIP= self.proxyIPEdit.text()
        configSaver.saveConfig(self)
        message = "設定を保存しました。"
        ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
        self.w.close()
    def loadConfig(self):
        global useInternetExplorerSetting, useProxy,proxyPort,proxyIP
        global userID

        configSaver.load(self)
        if useInternetExplorerSetting:
            self.proxyComboBox.setCurrentIndex(0)
        else:
            self.proxyComboBox.setCurrentIndex(1)
        self.useProxyCheckBox.setChecked(useProxy)
        self.proxyPortEdit.setText(proxyPort)
        self.proxyIPEdit.setText(proxyIP)

    def proxyCheckBoxOnChanged(self, state):
        # チャックが入っているとき
        index=self.proxyComboBox.currentIndex()
        if index==0:
            return
        if state == Qt.Checked :
            self.proxyPortEdit.setEnabled(True)
            self.proxyIPEdit.setEnabled(True)

        # チャックが入っていないとき
        else:
            # 何も引数を指定しないとタイトルがpythonになるみたい
            self.proxyPortEdit.setEnabled(False)
            self.proxyIPEdit.setEnabled(False)

            self.proxyPortEdit.setText("")
            self.proxyIPEdit.setText("")

    def changeProxyType(self, index):
        if(index==1):#自分で設定

            self.useProxyCheckBox.setEnabled(True)
            self.proxyPortEdit.setEnabled(True)
            self.proxyIPEdit.setEnabled(True)
            self.proxyCheckBoxOnChanged(self.useProxyCheckBox.checkState())

        else:
            self.useProxyCheckBox.setEnabled(False)
            self.proxyPortEdit.setEnabled(False)
            self.proxyIPEdit.setEnabled(False)

        if (IEProxyIP != ""):
            self.useProxyCheckBox.setChecked(True)
            self.proxyPortEdit.setText(IEProxyPort)
            self.proxyIPEdit.setText(IEProxyIP)
        else:

            self.useProxyCheckBox.setChecked(False)
            self.proxyPortEdit.setText("")
            self.proxyIPEdit.setText("")

class configSaver:
    def saveConfig(self):

        global useInternetExplorerSetting,useProxy,proxyIP,proxyPort,proxyDict
        config = configparser.ConfigParser()
        config.read(os.getenv('APPDATA')+'\\MediMonitor\\config.ini')


        config['proxy'] = {
            'useInternetExplorerSetting': useInternetExplorerSetting,
            'useProxy': useProxy,
            'proxyIP': proxyIP,
            'proxyPort': proxyPort
        }
        config['main'] = {
            'fileFolder': fileFolder,
            'fileType': fileType,
            'fileTypeVer': fileTypeVer,
            'privacyInfo': privacyInfo,
            'fileSuccessFlag': fileSuccessFlag

        }
        config['mfnet'] = {
            't': MFtoken,
            'userid': userID,
            'username': userName,
            'useremail': userEmail
        }

        if useProxy :
            proxyDict={ "http"  : "http://"+proxyIP+":"+proxyPort,  "https" : "https://"+proxyIP+":"+proxyPort}
        else :
            proxyDict=None

        print("saveConfigを保存しました")
        try:

            if not os.path.exists(os.getenv('APPDATA') + "\\MediMonitor"):
                os.mkdir(os.getenv('APPDATA') + "\\MediMonitor")

            with open(os.getenv('APPDATA')+'\\MediMonitor\\config.ini', 'w') as f:
                config.write(f)
        except:
            print("saveConfigの保存にエラーが起きました")

    def load(self):
        localAppIni=os.getenv('APPDATA')+'\\MediMonitor\\config.ini'
        global useInternetExplorerSetting,useProxy,proxyIP,proxyPort,proxyDict
        global fileType,fileTypeVer,fileFolder,fileSuccessFlag,privacyInfo
        global loginName,loginPass,isLogined,userID,userName,userEmail,MFtoken
        getIEsetting(self)
        print (localAppIni)
        if not os.path.exists(localAppIni):
            configSaver.createDefaultConfig(self)

        config = configparser.ConfigParser()
        config.read(localAppIni)
        try:
            config["proxy"]
        except:
            configSaver.createDefaultConfig(self)
            configSaver.load(self)
            return


        if config["proxy"]["useInternetExplorerSetting"] =="True" :
            useInternetExplorerSetting=True

            if (IEProxyIP != ""):
                useProxy= True
                proxyIP=IEProxyIP
                proxyPort=IEProxyPort
            else:
                useProxy = False
                proxyIP = ""
                proxyPort = ""

        else :
            useInternetExplorerSetting = False
            if config["proxy"]["useProxy"] == "True":
                useProxy = True
            else:
                useProxy = False

            try:
                proxyIP = config["proxy"]["proxyIP"]
                proxyPort = config["proxy"]["proxyPort"]
            except:
                print("No Key proxy.An exception occurred")
                proxyIP = ""
                proxyPort = ""


        if useProxy :
            proxyDict={ "http"  : "http://"+proxyIP+":"+proxyPort,  "https" : "https://"+proxyIP+":"+proxyPort}
        else :
            proxyDict=None
        try:
            config["main"]
            fileFolder= config["main"]["fileFolder"]
            fileType = int(config["main"]["fileType"])
            fileTypeVer= int(config["main"]["fileTypeVer"])
            privacyInfo = int(config["main"]["privacyInfo"])

            if config["main"]["fileSuccessFlag"] == "True":
                fileSuccessFlag=True
            else:
                fileSuccessFlag=False

        except:

            fileFolder= ""
            fileType = 0
            fileTypeVer= 0
            privacyInfo= 0
        try:
            config["mfnet"]
            MFtoken = config["mfnet"]["t"]
            userID= config["mfnet"]["userid"]
            userName= config["mfnet"]["username"]
            userEmail= config["mfnet"]["useremail"]
        except:
            MFtoken = ""
            userID= ""
            userName= ""
            userEmail= ""
    def createDefaultConfig(self):
        global proxyPort,proxyIP
        config = configparser.ConfigParser()

        config['proxy'] = {
            'useInternetExplorerSetting': "True",
            'useProxy': "False",
            'proxyIP': "",
            'proxyPort': ""
        }
        config['main'] = {
            'fileType': 0,
            'fileTypeVer':0,
            'privacyInfo': 0,
        }
        config['mfnet'] = {
            't': "",
            'userid': "",
            'username': "",
            'useremail': "",

        }
        if not os.path.exists(os.getenv('APPDATA')+"\\MediMonitor"):
            os.mkdir(os.getenv('APPDATA')+"\\MediMonitor")

        with open(os.getenv('APPDATA')+'\\MediMonitor\\config.ini', 'w') as config_file:
            config.write(config_file)

class httpsRequest(QThread):

    def asyncTokenRequest(self,proxies,token,user_id,csv="",csvName="",csvType="0",csvVersion="0"):
        global response
        android_id="pc"
        nonce=int(time.time()*1000)
        response = httpsRequest.customHTTPrequest(self,
                                                  url='https://' + LOGIN_SERVER_IP + '/php/medi/mf_system_core/post_csv_from_pc.php',
                                                  data={'MFNetUserID': user_id, 'MFNetToken': token,
                                                        'MFNetAndroidID': android_id, 'MFNetCSV': csv, 'MFNetCSVVersion': csvVersion, 'MFNetCSVType': csvType, 'MFNetNonce': nonce,
                                                        'MFNetCSVname': csvName,'pc_version': VERSION_INFO}, proxies=proxies, showMessage=False)

    def viewRepaint(self):
        global proxyDict,userID

        if response is None :

            self.onlineMessageLabel.setText("<b>接続エラー 再試行します</b>")
            self.repaint()
            return
        if response.status_code==404:
            print("token 404")
            self.onlineMessageLabel.setText("<b>サーバー上にエラー？</b>（メンテナンス中の可能性があります）")

            self.passwordEdit.setEnabled(False)
            self.passwordEdit.setText("00000000")
            self.pharmIDEdit.setEnabled(False)
            self.pharmIDEdit.setText(userID)
            self.connectButton.setText("ログアウト")

            self.connectButton.setEnabled(True)
            self.repaint()
            return


        if response.status_code is not 200 :
            print("Not200")
            self.logoutProcess(False)
            message = "システムからログアウトされました。\n再度ログインしてください。"
            ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
            self.setVisible(True)
            return
        print("Tokenは認証されました")
        print(response.status_code)  # HTTPのステータスコード取得
        print(response.text)  # レスポンスのHTMLを文字列で取得
        r=json.loads(response.text)
        if "this" in r :
            if r["this"] == "kara":
                print("kara")
                httpsRequest.loginSuccessProcess(self,MFtoken,userID,userEmail,userName,False)
            elif r["this"] == "needUpdate":
                message = "本ソフトウェアを最新版にアップデートする必要があります。公式サイトより最新版をダウンロードし、インストールしてください。"
                ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
                import webbrowser
                webbrowser.open('https://medicalfields.jp/')
                QApplication.instance().quit()


            elif r["this"] == "ok":
                print("kara")
                self.localMessageTextEdit.setText("レセコンから受信した連動用ファイルをオンラインシステムへ転送しました。")
                self.repaint()

    def customHTTPrequest(self,url,data,proxies,showMessage=True):
        response=None
        try:
            response = requests.post(url,
                                     data=data,
                                     proxies=proxies)
        except ProxyError:
            print("ProxyError")
            if showMessage:
                message = "プロキシーが有効になっていますが、サーバーからの応答がありません。プロキシーを無効にするか、インターネットに接続されているかをご確認ください。"
                ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
        except requests.exceptions.RequestException as e:
            print(e)
            if showMessage:
                message = "サーバーからの応答がありません。インターネットに接続されているかをご確認ください。またプロキシーの設定が必要な場合は詳細設定よりプロキシーの設定を行ってください。"
                ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
        except Timeout:
            print("Timeout")
            if showMessage:
                message = "アクセスが集中してサーバーが応答していない可能性があります。しばらくしてから実行してください。"
                ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
        return response

    def loginRequest(self,user_id,user_pw,android_id):

        configSaver.saveConfig(self)
        global loginName, loginPass, isLogined, userID, userName, userEmail, MFtoken
        global proxyDict

        self.connectButton.setText("接続中・・・")
        self.localMessageTextEdit.setText("オンラインシステムへ接続中です。しばらくお待ち下さい。")
        self.passwordEdit.setEnabled(False)
        self.pharmIDEdit.setEnabled(False)
        self.connectButton.setEnabled(False)
        self.repaint()

        response = httpsRequest.customHTTPrequest(self,url='https://'+LOGIN_SERVER_IP+'/php/medi/mf_system_core/mf_user_id_login.php', data={'user_id': user_id, 'user_pw': user_pw , 'android_id': android_id,'pc_version': VERSION_INFO}, proxies=proxyDict)
        if response ==None :
            self.logoutProcess(False)
            return
        print(response.status_code)  # HTTPのステータスコード取得
        print(response.text)  # レスポンスのHTMLを文字列で取得
        if response.status_code==200:
            r = json.loads(response.text)
            if "this" in r:
                if r["this"] == "invalid":
                    message = "ログインに失敗しました。薬局IDとパスワードが正しいかご確認ください。"
                    ret = QMessageBox.warning(None, "失敗", message, QMessageBox.Yes)
                elif r["this"] == "locked":
                    message = "ログインに失敗しました\nパスワードの上限試行回数に達したか、その他のセキュリティ上の理由でアカウントがロックされました。パスワードを再発行してください。\n" \
                              "パスワードを再発行した場合、現在ログイン中の端末はすべてログアウトされます。"
                    ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
                elif r["this"] == "needUpdate":
                    message = "本ソフトウェアを最新版にアップデートする必要があります。公式サイトより最新版をダウンロードし、インストールしてください。"
                    ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
                elif r["this"] == "regionError":
                    yourIP=""
                    if "ip" in r:
                        yourIP=r["ip"]
                    message = "ログインに失敗しました\n申し訳ございませんが、お使いの地域ではログインすることができません\n" \
                              "日本国内でこのエラーが出る場合はこちらのIPアドレス["+yourIP+"]をサポートセンターまでご連絡下さい。"
                    ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                elif r["this"] == "no_serial_key":
                    message = "ログインに失敗しました\nシリアルキーが登録されていません。アプリより有効なシリアルキーの登録か、サブスクリプションを登録した後、実行してください"
                    ret = QMessageBox.warning(None, "確認", message, QMessageBox.Yes)
                elif r["this"] == "no_valid_key_left":
                    message = "ログインに失敗しました\nこのアカウントで認証可能なライセンス数の上限に達しています。他の端末を解除してから再ログインしてください。\n" \
                              "またログインできる端末がわからない場合はパスワードを再発行してください。パスワードを再発行した場合、現在ログイン中の端末はすべてログアウトされます。"
                    ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                elif r["this"] == "no_client_left":
                    message = "ログインに失敗しました\nこのアカウントで許可されている端末の上限に達しています。他の端末からログアウトするか、端末の上限を拡大してから再ログインしてください。\n" \
                              "またログインできる端末がわからない場合はパスワードを再発行してください。パスワードを再発行した場合、現在ログイン中の端末はすべてログアウトされます。"
                    ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                self.logoutProcess(False)

                return

            httpsRequest.loginSuccessProcess(self, r["token"], r["user_id"], r["user_mail_address"], r["user_name"],
                                                  True)
        else :
            message = "サーバーから応答はありましたが、データが受理されませんでした。\nメンテナンス中の可能性があります。しばらくした後実行してください。"
            ret = QMessageBox.warning(None, "エラー", message, QMessageBox.Yes)
            self.logoutProcess(False)

    def loginSuccessProcess(self,token,user_id,user_email,user_name,showMessage):

        global loginName, loginPass, isLogined,MFtoken,userID,userEmail,userName
        MFtoken=token
        userID=user_id
        userEmail=user_email
        userName=user_name
        isLogined=True
        message = "ログインに成功しました\nバックグランドで動作します"
        configSaver.saveConfig(self)

        self.passwordEdit.setEnabled(False)
        self.passwordEdit.setText("00000000")
        self.pharmIDEdit.setEnabled(False)
        self.pharmIDEdit.setText(user_id)
        self.connectButton.setText("ログアウト")

        self.connectButton.setEnabled(True)
        self.onlineMessageLabel.setText("<b>ログイン中</b> ("+userName+"様："+userEmail+")")
        self.localMessageTextEdit.setText("オンラインシステムのログインに成功しました。連動先フォルダのチェックを行います。")
        self.repaint()

        if showMessage:
            self.startThread()
            if fileSuccessFlag:
                ret = QMessageBox.information(None, "確認", message, QMessageBox.Yes)
                self.setVisible(False)
            else :
                ret = QMessageBox.information(None, "確認", "オンラインシステムとの接続は完了しました。レセプトコンピューターから連動用のデータを出力してください。", QMessageBox.Yes)

if __name__ == '__main__':
    import sys
    app = QApplication(sys.argv)
    translator = QTranslator()
    translator.load('qt_' + QLocale.system().name(), QLibraryInfo.location(QLibraryInfo.TranslationsPath))
    app.installTranslator(translator)
    app.setAttribute(Qt.AA_DisableWindowContextHelpButton)

    if not QSystemTrayIcon.isSystemTrayAvailable():
        QMessageBox.critical(None, "システムトレイ",
                "このシステムではシステムトレイを作成することができませんでした")
        sys.exit(1)
    QApplication.setQuitOnLastWindowClosed(False)
    window = Window()
    window.show()
    global fileFolder
    if fileFolder == "":
        window.show_wizard_dialog()
    sys.exit(app.exec_())
