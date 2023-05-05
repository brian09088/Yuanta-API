import json
import os
import re
import sys
import numpy

import asciichartpy

# 停用 loguru 的 stdout 輸出
os.environ['LOGURU_AUTOINIT'] = 'False'

import ctypes
import comtypes
import comtypes.client
import datetime
import dateutil.relativedelta
import loguru
import wcwidth
import wx

class YuantaQuoteAXCtrl:
    def __init__(self, parent):
        self.parent = parent

        container = ctypes.POINTER(comtypes.IUnknown)()
        control = ctypes.POINTER(comtypes.IUnknown)()
        guid = comtypes.GUID()
        sink = ctypes.POINTER(comtypes.IUnknown)()

        ctypes.windll.atl.AtlAxCreateControlEx(
            'YUANTAQUOTE.YuantaQuoteCtrl.1',
            self.parent.Handle,
            None,
            ctypes.byref(container),
            ctypes.byref(control),
            ctypes.byref(guid),
            sink
        )
        self.ctrl = comtypes.client.GetBestInterface(control)
        self.sink = comtypes.client.GetEvents(self.ctrl, self)

        self.Host = None
        self.Port = None
        self.Username = None
        self.Password = None

        # 儲存繪圖點位
        self.Series = []

    def Config(self, host, port, username, password):
        self.Host = host
        self.Port = port
        self.Username = username
        self.Password = password

    def XXF(self):
        today = datetime.date.today()
        day = datetime.date.today().replace(day=1)
        while day.weekday() != 2:
            day = day + datetime.timedelta(days=1)
        day = day + dateutil.relativedelta.relativedelta(days=14)
        if day < today:
            day = day + dateutil.relativedelta.relativedelta(months=1)
        codes = [
            'A',
            'B',
            'C',
            'D',
            'E',
            'F',
            'G',
            'H',
            'I',
            'J',
            'K',
            'L'
        ]
        y = day.year % 10
        m = codes[day.month - 1]
        return f'{m}{y}'

    def TXF(self):
        return f'TXF{self.XXF()}'

    def Logon(self):
        self.ctrl.SetMktLogon(
            self.Username,
            self.Password,
            self.Host,
            self.Port,
            1,
            0
        )

    # 省略未使用事件函數

    def OnGetMktAll(self,
        symbol,
        refPri,
        openPri,
        highPri,
        lowPri,
        upPri,
        dnPri,
        matchTime,
        matchPri,
        matchQty,
        tolMatchQty,
        bestBuyQty,
        bestBuyPri,
        bestSellQty,
        bestSellPri,
        fdbPri,
        fdbQty,
        fdsPri,
        fdsQty,
        reqType):
        # 因畫面寬度關係，僅保存最後 165 個點位
        self.Series.append(float(matchPri))
        if len(self.Series) > 165:
            self.Series = self.Series[1:]

        now = datetime.datetime.now()
        clX = f' {now:%H:%M:%S.%f} '
        cl01 = f' {symbol: >6} '
        cl02 = f' {reqType: >4} '
        cl03 = f' {refPri: >10} '
        cl04 = f' {openPri: >10} '
        cl05 = f' {highPri: >10} '
        cl06 = f' {lowPri: >10} '
        cl07 = f' {upPri: >10} '
        cl08 = f' {dnPri: >10} '
        cl09 = f' {matchTime: >15} '
        if matchTime:
            cl09 = f' {matchTime[:2]}:{matchTime[2:4]}:{matchTime[4:6]}.{matchTime[6:]} '
        cl10 = f' {matchPri: >10} '
        cl11 = f' {matchQty: >8} '
        cl12 = f' {tolMatchQty: >8} '

        # 產生標題行
        date = f'{datetime.date.today():%Y-%m-%d}'
        title = '                                                                                        MatchPrice                                                                              '
        # 如果有二個點以上，則產生折線圖
        # 原因是 asciichartpy 套件的 Y 軸需要二個點相減來產生，不然會發生例外
        chart = ''
        if len(set(self.Series)) > 1:
            chart = asciichartpy.plot(self.Series, {
                'width': 116,
                'height': 30
            }) 
        sys.stdout.write(
            '\033[2J' +
            f'          元大期貨                                                                                                                                                    {  date  }\n' +
            f'          ----------------------------------------------------------------------------------------------------------------------------------------------------------------------\n' +
            f'          |        接收時間 |   代碼 | 盤別 |     參考價 |     開盤價 |     最高價 |     最低價 |     漲停價 |     跌停價 |        成交時間 |   成交價位 | 成交數量 | 總成交量 |\n' +
            f'          ----------------------------------------------------------------------------------------------------------------------------------------------------------------------\n' +
            f'          |{           clX}|{  cl01}|{cl02}|{      cl03}|{      cl04}|{      cl05}|{      cl06}|{      cl07}|{      cl08}|{           cl09}|{      cl10}|{    cl11}|{    cl12}|\n' +
            f'          ----------------------------------------------------------------------------------------------------------------------------------------------------------------------\n' +
            f'\n' +
            f'{title}\n' +
            f'\n' +
            f'{chart}\n'
        )

    def OnMktStatusChange(self,
        status,
        msg,
        reqType):
        code = ''
        find = re.match(r'^(?P<code>\d).+$', msg)
        if find:
            code = msg[0]
            msg = msg[1:]
        if msg.endswith('!'):
            msg = msg[:-1]
        msg = msg + ' ' * (50 - wcwidth.wcswidth(msg))
        clX = f' {datetime.datetime.now():%H:%M:%S.%f} '
        cl01 = f' {reqType: >7} '
        cl02 = f' {status: >6} '
        cl03 = f' {code: >4} '
        cl04 = f' {msg} '

        loguru.logger.info(
            f'OnMktStatusChange\n' +
            f'--------------------------------------------------------------------------------------------------\n' +
            f'|                 | reqType | status | code | msg                                                |\n' +
            f'--------------------------------------------------------------------------------------------------\n' +
            f'|{            clX}|{   cl01}|{  cl02}|{cl03}|{cl04                                              }|\n' +
            f'--------------------------------------------------------------------------------------------------\n'
        )

        if status != 2:
            return

        code = self.TXF()
        loguru.logger.info(f'OnMktStatusChange: AddMktReg(code={code}, 4, reqType={reqType}, 0)\n')
        result = self.ctrl.AddMktReg(code, 4, reqType, 0)
        loguru.logger.info(f'OnMktStatusChange: AddMktReg(code={code}, 4, reqType={reqType}, 0) => {result}\n')

    def OnGetTimePack(self,
        strTradeType,
        strTime,
        reqType):
        clX = f' {datetime.datetime.now():%H:%M:%S.%f} '
        cl01 = f' {reqType: >7} '
        cl02 = f' {strTime[:2]}:{strTime[2:4]}:{strTime[4:6]}.{strTime[6:]} '
        cl03 = f' {strTradeType: >12} '
        loguru.logger.info(
            f'OnGetTimePack\n' +
            f'--------------------------------------------------------------\n' +
            f'|                 | reqType |         strTime | strTradeType |\n' +
            f'--------------------------------------------------------------\n' +
            f'|{            clX}|{   cl01}|{           cl02}|{        cl03}|\n' +
            f'--------------------------------------------------------------\n'
        )

    # 省略未使用事件函數

def main():
    app = wx.App()

    frame = wx.Frame(
        parent=None,
        id=wx.ID_ANY,
        title='Yuanta.Quote'
    )
    frame.Hide()

    quote = YuantaQuoteAXCtrl(frame)
    quote.Config(
        host='apiquote.yuantafutures.com.tw',
        port='80',
        username='M123143296',
        password='yaunda123'
    )
    quote.Logon()

    app.MainLoop()

if __name__ == '__main__':
    loguru.logger.add(
        f'{datetime.date.today():%Y%m%d}.log',
        rotation='1 day',
        retention='7 days',
        level='DEBUG'
    )
    main()