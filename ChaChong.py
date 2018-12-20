#!/usr/bin/env python3
# -*- coding: UTF-8 -*-
import os
import sys
import time
import codecs
from win32com.client import Dispatch

if len(sys.argv) != 2:
    print("命令参考：python chachong 文件包位置不要加斜线")
rootDir=sys.argv[1]
rtdirname=os.path.split(sys.argv[1])
tgdirname=os.getcwd()

def HuiDuanLuoShu(WenJian):
    if WenJian[0] == r"~" :
        return 0
    elif WenJian[0] == r".":
        return 0
    else:
        wordApp = Dispatch('word.Application')
        wordApp.Visible = False
        wordApp.DisplayAlerts = 0
        myDoc = wordApp.Documents.Open(FileName = os.path.join(rootDir, WenJian))
#        myRange = myDoc.Range(0, 0)
        dLs = myDoc.Paragraphs.Count
        myDoc.Close()
        wordApp.Quit()
        return dLs

def HuiDuanLuoWenZi(WenJian, duanluoHao):
    if WenJian[0] == r"~" :
        return ""
    elif WenJian[0] == r".":
        return ""
    else:
        wordApp = Dispatch('word.Application')
        wordApp.Visible = False
        wordApp.DisplayAlerts = 0
        myDoc = wordApp.Documents.Open(FileName = os.path.join(rootDir, WenJian))
#        wordApp.Documents.Selection.Find.Execute(" ", False, False, False, False, False, True, 1, True, "", 2)
#        myRange = myDoc.Range(0, 0)
        WenZi= myDoc.Paragraphs(duanluoHao).Range.text
        WenZi.replace(" ", "")
        myDoc.Close()
        wordApp.Quit()
        return WenZi

def BiJiaoLa(WenZi, WenJian):
    if WenJian[0] == r"~":
        return False
    elif WenJian[0] == r".":
        return False
    else:
        wordApp = Dispatch('word.Application')
        wordApp.Visible = False
        wordApp.DisplayAlerts = 0
        myDoc02 = wordApp.Documents.Open(FileName = os.path.join(rootDir, WenJian))
        wordApp.Selection.Find.Execute(" ", False, False, False, False, False, True, 1, True, "", 2)
        alltext=0
        for i in range(1, myDoc02.Paragraphs.Count):
            alltext=alltext + len(myDoc02.Paragraphs(i).Range.text)
        myRange = myDoc02.Range(0, alltext)
        myDoc02.Close()
        wordApp.Quit()
        if WenZi in myRange.text:
            return True

if __name__ == "__main__":
    f=codecs.open(tgdirname + '/查重结果.txt', 'a', encoding='utf-8')
    f.write(time.asctime(time.localtime(time.time())) + '\n')
    f.close()
    for root, dirs, files in os.walk(rootDir):
        for name in files:
            if name[0] == r"~" or name[0] == r".":
                continue
            else:
                print(name)
                dLs=HuiDuanLuoShu(name)
                for duanluoHao in range(1, dLs):
                    BJWenZi = HuiDuanLuoWenZi(name, duanluoHao)
                    if len(BJWenZi)>300:
                        for name02 in files:
                            if name02 != name:
                                BJchang=int(len(BJWenZi)/3)
                                BiJiaoZhi=BJWenZi[BJchang:BJchang*2]
                                if BiJiaoLa(BiJiaoZhi, name02):
                                    f=codecs.open(tgdirname + '/查重结果.txt', 'a', encoding='utf-8')
                                    f.write("section %s of %s can be found in  %s \n Content:%s \n ==  == == == == == == == == == =\n" 
                                            % (duanluoHao, name.center(20), name02.center(20), BiJiaoZhi[:30]))
                                    f.close()
    print("Great Job is Done!")
