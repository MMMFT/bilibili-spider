# !/usr/bin/python3
# -*- coding:utf-8 -*-

import os
import sys
import requests
import csv
from threading import Thread
import json
import time
from openpyxl import Workbook

Threads = []
infos = []
headers = {
    'User-Agent':'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/59.0.3071.104 Safari/537.36',
}
proxies = {
    'http': 'http://120.26.110.59:8080',
    'http': 'http://120.52.32.46:80',
    'http': 'http://218.85.133.62:80',
}

def CheckArgs():
    if(len(sys.argv) < 2):
        print('Usage: python video.py [startAid] [endAid]')
        return False
    return True

def GetInfos():
    startAid = sys.argv[1]
    endAid = sys.argv[2]
    for i in range(int(startAid), int(endAid)):
        Request(i)

def Request(aid):
    arr = []
    arr.append(aid)
    t = Thread(target=SendRequest, args=(arr,))
    Threads.append(t)

def SendRequest(n):
    url = 'https://api.bilibili.com/x/web-interface/view?aid=' + str(n[0])
    print('正在获取av' + str(n[0]) + '的信息')
    r = s.get(url, headers=headers, proxies=proxies)
    text = r.text
    info_json = json.loads(text)
    if info_json['code'] != 0:
        return
    infos.append(info_json)

def writeXls(fileName, sheetName, data):
    wb = Workbook()
    filename = fileName + '.xlsx'
    ws1 = wb.active
    ws1.title = sheetName
    ws1['A1'] = "av号"
    ws1['B1'] = "标题"
    ws1['C1'] = "描述"
    ws1['D1'] = "UP主Id"
    ws1['E1'] = "UP主名称"
    ws1['F1'] = "播放数"
    ws1['G1'] = "弹幕数"
    ws1['H1'] = "收藏数"
    ws1['I1'] = "硬币数"
    ws1['J1'] = "点赞数"
    ws1['K1'] = "创建时间"
    for idx, val in enumerate(data):
        col_A = 'A%s' % (idx + 2)
        col_B = 'B%s' % (idx + 2)
        col_C = 'C%s' % (idx + 2)
        col_D = 'D%s' % (idx + 2)
        col_E = 'E%s' % (idx + 2)
        col_F = 'F%s' % (idx + 2)
        col_G = 'G%s' % (idx + 2)
        col_H = 'H%s' % (idx + 2)
        col_I = 'I%s' % (idx + 2)
        col_J = 'J%s' % (idx + 2)
        col_K = 'K%s' % (idx + 2)
        ws1[col_A] = val['data']['aid']
        ws1[col_B] = val['data']['title']
        ws1[col_C] = val['data']['desc']
        ws1[col_D] = val['data']['owner']['mid']
        ws1[col_E] = val['data']['owner']['name']
        ws1[col_F] = val['data']['stat']['view']
        ws1[col_G] = val['data']['stat']['danmaku']
        ws1[col_H] = val['data']['stat']['favorite']
        ws1[col_I] = val['data']['stat']['coin']
        ws1[col_J] = val['data']['stat']['like']
        timeStamp = val['data']['pubdate']
        timeArray = time.localtime(timeStamp)
        otherStyleTime = time.strftime("%Y-%m-%d %H:%M:%S", timeArray)
        ws1[col_K] = otherStyleTime
    wb.save(filename=filename)


if __name__ == '__main__':
    if(CheckArgs() == False):
        sys.exit(-1)
    s = requests.Session()
    GetInfos()
    for t in range(len(Threads)):
        time.sleep(1)
        Threads[t].start()
    for t in range(len(Threads)):
        Threads[t].join()
    print('Info have been saved ')
    startAid = sys.argv[1]
    endAid = sys.argv[2]
    sname = startAid + '--' + endAid
    writeXls('bili视频信息统计', sname, infos)