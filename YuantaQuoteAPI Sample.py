import re

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
        # 父層元件必須是頂層視窗，才能正常接收事件
        self.parent = parent

        # 準備呼叫 ATL API 時所需參數
        container = ctypes.POINTER(comtypes.IUnknown)()
        control = ctypes.POINTER(comtypes.IUnknown)()
        guid = comtypes.GUID()
        sink = ctypes.POINTER(comtypes.IUnknown)()

        # 建立 ActiveX 控制元件實體
        ctypes.windll.atl.AtlAxCreateControlEx(
            'YUANTAQUOTE.YuantaQuoteCtrl.1',
            self.parent.Handle,
            None,
            ctypes.byref(container),
            ctypes.byref(control),
            ctypes.byref(guid),
            sink
        )
        # 取得 ActiveX 控制元件實體
        self.ctrl = comtypes.client.GetBestInterface(control)
        # 綁定 ActiveX 控制元件事件
        self.sink = comtypes.client.GetEvents(self.ctrl, self)

        # 連線資訊
        self.Host = None
        self.Port = None
        self.Username = None
        self.Password = None

    # 設定連線資訊
    def Config(self, host, port, username, password):
        self.Host = host
        self.Port = port
        self.Username = username
        self.Password = password

    # 登入元大服務
    def Logon(self):
        # 登入日盤主機
        self.ctrl.SetMktLogon(
            self.Username,
            self.Password,
            self.Host,
            self.Port,
            1,
            0
        )

    # ActiveX 控制元件事件
    def OnGetBreakResume(self,
        symbol,
        breakTime,
        resumeTime,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetDelayClose(self,
        symbol,
        delayClose,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetDelayOpen(self,
        symbol,
        delayOpen,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetFutStatus(self,
        symbol,
        functionCode,
        breakTime,
        startTime,
        reopenTime,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetLimitChange(self,
        symbol,
        functionCode,
        statusTime,
        level,
        expandType,
        reqType):
        pass

    # ActiveX 控制元件事件
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
        pass

    # ActiveX 控制元件事件
    def OnGetMktData(self,
        priType,
        symbol,
        qty,
        pri,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnGetMktQuote(self,
        symbol,
        disClosure,
        duration,
        reqType):
        pass

    # ActiveX 控制元件事件
    # 連線行為造成的狀態改變會從這個事件通知
    def OnMktStatusChange(self,
        status,
        msg,
        reqType):
        # 取出訊息開頭可能存在的連線狀態代碼
        code = ''
        find = re.match(r'^(?P<code>\d).+$', msg)
        if find:
            code = msg[0]
            msg = msg[1:]
        # 去除訊息結尾可能存在的驚嘆號
        if msg.endswith('!'):
            msg = msg[:-1]
        # 計算文字寬度並補上結尾空白
        msg = msg + ' ' * (50 - wcwidth.wcswidth(msg))
        # 產出欄位值
        clX = f' {datetime.now():%H:%M:%S.%f} '
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

    # ActiveX 控制元件事件
    # 登入後不定期會接收到元大服務主機發送的時戳，供客戶端進行校時使用
    def OnGetTimePack(self,
        strTradeType,
        strTime,
        reqType):
        clX = f' {datetime.now():%H:%M:%S.%f} '
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

    # ActiveX 控制元件事件
    def OnGetTradeStatus(self,
        symbol,
        tradeStatus,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnRegError(self,
        symbol,
        updMode,
        errCode,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnTickRangeDataError(self,
        symbol,
        errCode,
        reqType):
        pass

    # ActiveX 控制元件事件
    def OnTickRegError(self,
        symbol,
        mode,
        errCode,
        reqType):
        pass

def main():
    app = wx.App()

    frame = wx.Frame(
        parent=None,
        id=wx.ID_ANY,
        title='Yuanta.Quote'
    )
    frame.Hide()

    # 加入封裝後的元大期貨 API 報價元件
    quote = YuantaQuoteAXCtrl(frame)
    # 設定連線資訊
    quote.Config(
        host='<Host>',
        port='<Port>',
        username='<Username>',
        password='<Password>'
    )
    # 登入元大服務
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