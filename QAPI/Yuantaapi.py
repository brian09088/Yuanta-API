#!/usr/bin/env python
# coding: utf-8

# In[2]:


# -*- coding: utf-8 -*-
import wx, time
import wx.grid
import  wx.lib.anchors as anchors
from ctypes import byref, POINTER, windll
from comtypes import IUnknown, GUID
from comtypes.client import GetModule,  GetBestInterface, GetEvents
import queue as queue

user32 = windll.user32
atl = windll.atl

q = queue.Queue()


# In[3]:




class AppFrame(wx.Frame):
    """
    A Frame that says AppFrame
    """
    def __init__(self, *args, **kw):
        # ensure the parent's __init__ is called
        super(AppFrame, self).__init__(*args, **kw)

        # create a panel in the frame
        pnl = wx.Panel(self)        

        ############################################################################
        #   API連線資訊
        wx.StaticBox(pnl, label='API連線資訊', pos=(1, 1), size=(200, 148))
        wx.StaticText(pnl, label='帳號', pos=(11, 30))
        wx.StaticText(pnl, label='密碼', pos=(11, 60))
        self.acc = wx.TextCtrl(pnl, pos=(45, 26), size=(100, 25))
        self.pwd = wx.TextCtrl(pnl, pos=(45, 55), size=(100, 25),style=wx.TE_PASSWORD )

        logon = wx.Button(pnl, wx.ID_ANY, label='登入', pos=(11, 85), size=(40, 30))
        logon.Bind(wx.EVT_BUTTON, self.OnLogonBtn)

        self.connect_status = wx.TextCtrl(pnl, pos=(11, 118), size=(150, 25))
        self.connect_status.Enable(False)
        
        ############################################################################                        

        ############################################################################
        #   API商品資訊
        wx.StaticBox(pnl, label='API商品資訊', pos=(220, 1), size=(210, 148))
        wx.StaticText(pnl, label='商品代碼', pos=(230, 30))

        self.symbol = wx.TextCtrl(pnl, pos=(290, 26), size=(120, 25))
        self.rbAm = wx.RadioButton(pnl, 1, label='T', pos=(230, 60), style=wx.RB_GROUP)
        self.rbPm = wx.RadioButton(pnl, 2, label='T+1', pos=(260, 60))

        register = wx.Button(pnl,wx.ID_ANY,label='註冊', pos=(230, 85), size=(40, 30))
        register.Bind(wx.EVT_BUTTON, self.OnRegisterBtn)

        unregister = wx.Button(pnl, wx.ID_ANY, label='取消註冊', pos=(290, 85), size=(70, 30))
        unregister.Bind(wx.EVT_BUTTON, self.OnUnRegisterBtn)

        self.quote_status = wx.TextCtrl(pnl, pos=(230, 118), size=(150, 25))
        self.quote_status.Enable(False)

        UpdateMode = ["1-Snapshot","2-Update","4-SnapshotUpd"]        
        self.modle = wx.Choice(pnl,choices = UpdateMode,pos=(310, 55),size=(100, 10))         
        self.modle.SetSelection(2)
        ###################################################################################

        ###################################################################################
        #   五檔
        self.BestBuySell = wx.TextCtrl(pnl, pos=(443, 8), size=(200, 140), style = wx.TE_MULTILINE)        
        self.BestBuySell.Enable(False)        
        ###################################################################################

        ###################################################################################
        #   報價攔
        # Create a wxGrid object
        self.grid = wx.grid.Grid(pnl,pos=(1,150),size=(643, 185))

        # Then we call CreateGrid to set the dimensions of the grid
        # (10 rows and 10 columns in this example)
        self.grid.CreateGrid(8, 7)
        for index in range(self.grid.GetNumberRows()):
            self.grid.SetRowLabelValue(index, "-----")
                
        self.grid.SetColLabelValue(0,"參考價")
        self.grid.SetColLabelValue(1, "開盤價")
        self.grid.SetColLabelValue(2, "最高價")
        self.grid.SetColLabelValue(3, "最低價")        
        self.grid.SetColLabelValue(4, "成交時間")        
        self.grid.SetColLabelValue(5, "成交價")
        self.grid.SetColLabelValue(6, "成交量")
        self.grid.SetDefaultCellBackgroundColour('BLACK')        
        self.grid.EnableEditing(False)
        #attr = wx.grid.GridCellAttr()
        #attr.SetReadOnly(True)
        #grid.SetRowAttr(0, attr)
        ###################################################################################

    def OnLogonBtn(self,event):
        LogonJob(Job.LOGON, self.acc.GetValue(),self.pwd.GetValue())        

    def OnRegisterBtn(self,event):
        updatemodle = self.modle.GetString(self.modle.GetSelection())
        if self.rbAm.GetValue() == True:
            RegisterJob(Job.REGISTER,self.symbol.GetValue(),1,updatemodle)
        else:
            RegisterJob(Job.REGISTER, self.symbol.GetValue(), 2,updatemodle)

    def OnUnRegisterBtn(self,event):
        if self.rbAm.GetValue() == True:
            UnRegisterJob(Job.UNREGISTER,self.symbol.GetValue(),1)
        else:
            UnRegisterJob(Job.UNREGISTER, self.symbol.GetValue(), 2)

    def UpdateSymbol(self,symbol, RefPri, OpenPri, HighPri, LowPri, MatchTim, MatchPri ,MatchQty ,BestBuyPri ,BestBuyQty ,BestSellPri ,BestSellQty):
        for index in range(self.grid.GetNumberRows()):
            # cell color define
            if self.grid.GetRowLabelValue(index) == symbol:
                #參考價
                self.grid.SetCellValue(index, 0,RefPri)
                self.grid.SetCellTextColour(index, 0,wx.WHITE)
                #self.grid.SetAttr(index, 0, self.unchanged_attr)

                # 開盤價
                self.grid.SetCellValue(index, 1, OpenPri)
                if OpenPri == RefPri:
                    self.grid.SetCellTextColour(index, 1, wx.WHITE)
                elif OpenPri > RefPri:
                    self.grid.SetCellTextColour(index, 1, wx.RED)
                else:
                    self.grid.SetCellTextColour(index, 1, wx.GREEN)

                #最高價
                self.grid.SetCellValue(index, 2, HighPri)
                if HighPri == RefPri:
                    self.grid.SetCellTextColour(index, 2, wx.WHITE)
                elif HighPri > RefPri:
                    self.grid.SetCellTextColour(index, 2, wx.RED)
                else:
                    self.grid.SetCellTextColour(index, 2, wx.GREEN)

                #最低價
                self.grid.SetCellValue(index, 3, LowPri)
                if LowPri == RefPri:
                    self.grid.SetCellTextColour(index, 3, wx.WHITE)
                elif LowPri > RefPri:
                    self.grid.SetCellTextColour(index, 3, wx.RED)
                else:
                    self.grid.SetCellTextColour(index, 3, wx.GREEN)

                #成交時間
                self.grid.SetCellValue(index, 4, MatchTim)                
                self.grid.SetCellTextColour(index, 4, wx.WHITE)

                #成交價
                self.grid.SetCellValue(index, 5, MatchPri)
                if MatchPri == RefPri:
                    self.grid.SetCellTextColour(index, 5, wx.WHITE)
                elif MatchPri > RefPri:
                    self.grid.SetCellTextColour(index, 5, wx.RED)
                else:
                    self.grid.SetCellTextColour(index, 5, wx.GREEN)

                #成交量
                self.grid.SetCellValue(index, 6, MatchQty)         
                if MatchPri == RefPri:
                    self.grid.SetCellTextColour(index, 6, wx.WHITE)
                elif MatchPri > RefPri:
                    self.grid.SetCellTextColour(index, 6, wx.RED)
                else:
                    self.grid.SetCellTextColour(index, 6, wx.GREEN)
                
                #買賣 價/五檔
                self.BestBuySell.SetValue('symbol : {}\nbuy : {} \nbuyQty : {} \nsell : {} \nsellQty : {} \n'.format (symbol,BestBuyPri,BestBuyQty,BestSellPri,BestSellQty))

    def InsertSymbol(self,symbol):
        for index in range(self.grid.GetNumberRows()):
            if self.grid.GetRowLabelValue(index) == symbol:
                #已經有了
                return
            elif self.grid.GetRowLabelValue(index) == "-----":
                # insert
                self.grid.SetRowLabelValue(index,symbol)
                return

    def DeleteSymbol(self,symbol):
        for index in range(self.grid.GetNumberRows()):
            if self.grid.GetRowLabelValue(index) == symbol:
                self.grid.SetRowLabelValue(index, "-----")
                self.grid.SetCellValue(index, 0, "")
                self.grid.SetCellValue(index, 1, "")
                self.grid.SetCellValue(index, 2, "")
                self.grid.SetCellValue(index, 3, "")
                self.grid.SetCellValue(index, 4, "")
                self.grid.SetCellValue(index, 5, "")
                self.grid.SetCellValue(index, 6, "")

                #self.grid.DeleteRows(index,1,True)
                return

    def SetConnectStatusValue(self,Value):
        self.connect_status.SetValue(Value)

    def SetQuoteStatusValue(self,Value):
        self.quote_status.SetValue(Value)



# In[4]:


class switch(object):
    def __init__(self, value):
        self.value = value
        self.fall = False

    def __iter__(self):
        """Return the match method once, then stop"""
        yield self.match
        yield self.match
        raise StopIteration

    def match(self, *args):
        """Indicate whether or not to enter a case suite"""
        if self.fall or not args:
            return True
        elif self.value in args: # changed for v1.5, see below
            self.fall = True
            return True
        else:
            return False


class Job:
    LOGON = 1
    REGISTER = 2
    UNREGISTER = 3
    QUOTE = 4
    INSERTSYMBOL = 5
    DELETESYMBOL = 6

    def __init__ (self, job_type):
        self.job_type = job_type
        #q.put (self)

class LogonJob(Job):
    def __init__(self, job_type,account,password):
        super(LogonJob, self).__init__(job_type)
        self.account = account
        self.password = password
        q.put(self)

class RegisterJob(Job):
    def __init__(self, job_type,regSymbol,AmPm,modle):
        super(RegisterJob, self).__init__(job_type)
        self.regSymbol = regSymbol
        self.AmPm = AmPm
        self.modle = modle
        q.put(self)

class UnRegisterJob(Job):
    def __init__(self, job_type,unSymbol,AmPm):
        super(UnRegisterJob, self).__init__(job_type)
        self.unSymbol = unSymbol
        self.AmPm = AmPm
        q.put(self)


class QuoteJob(Job):
    def __init__(self, job_type, symbol, RefPri, OpenPri, HighPri, LowPri,  MatchTime, MatchPri, MatchQty, BestBuyPri, BestBuyQty, BestSellPri, BestSellQty):
        super(QuoteJob, self).__init__(job_type)
        self.symbol = symbol
        self.RefPri = RefPri
        self.OpenPri = OpenPri
        self.HighPri = HighPri
        self.LowPri = LowPri
        self.MatchTime = MatchTime
        self.MatchPri = MatchPri
        self.MatchQty = MatchQty
        self.BestBuyPri = BestBuyPri
        self.BestBuyQty = BestBuyQty
        self.BestSellPri = BestSellPri
        self.BestSellQty = BestSellQty
        q.put(self)

class InsertSymbol(Job):
    def __init__(self, job_type, insertSymbol):
        super(InsertSymbol, self).__init__(job_type)
        self.symbol = insertSymbol
        q.put(self)

class DeleteSymbol(Job):
    def __init__(self, job_type, deleteSymbol):
        super(DeleteSymbol, self).__init__(job_type)
        self.symbol = deleteSymbol
        q.put(self)

def DoJob(Bot, x):
    for case in switch(x.job_type):
        if case(Job.LOGON):
            Bot.login (x.account,x.password)
            break
        if case(Job.REGISTER):
            print(x.job_type,x.regSymbol,x.modle[0])            
            Bot.RegisterQuoteSymbol(x.regSymbol,x.modle,x.AmPm)
            break
        if case(Job.UNREGISTER):
            print(x.job_type,x.unSymbol)
            Bot.UnRegisterQuoteSymbol(x.unSymbol,x.AmPm)
            break
        if case(Job.QUOTE):
            frame.UpdateSymbol(x.symbol,x.RefPri, x.OpenPri, x.HighPri, x.LowPri, x.MatchTime, x.MatchPri, x.MatchQty, x.BestBuyPri, x.BestBuyQty, x.BestSellPri, x.BestSellQty)
            break
        if case(Job.INSERTSYMBOL):
            frame.InsertSymbol(x.symbol)
            break
        if case(Job.DELETESYMBOL):
            frame.DeleteSymbol(x.symbol)
            break


class YuantaQuoteEvents(object):
    def __init__(self, parent):
        self.parent = parent
    def OnMktStatusChange (self, this, Status, Msg, ReqType):
        print ('OnMktStatusChange {},{},{}'.format (ReqType, Msg, Status))
        frame.SetConnectStatusValue(Msg)

    def OnRegError(self, this, symbol, updmode, ErrCode, ReqType):
        print ('OnRegError {},{},{},{}'.format (ReqType, ErrCode, symbol, updmode))

    def OnGetMktData(self, this, PriType, symbol, Qty, Pri, ReqType):
        print ('OnGetMktData')

    def OnGetMktQuote(self, this, symbol, DisClosure, Duration, ReqType):
        print ('OnGetMktQuote')

    def OnGetMktAll(self, this, symbol, RefPri, OpenPri, HighPri, LowPri, UpPri, DnPri, MatchTime, MatchPri, MatchQty, TolMatchQty,
        BestBuyQty, BestBuyPri, BestSellQty,BestSellPri, FDBPri, FDBQty, FDSPri, FDSQty, ReqType):
        print ('OnGetMktAll')
        print ('buy{} buyQty{} sell{} sellQty{}'.format (BestBuyPri,BestBuyQty,BestSellPri,BestSellQty))
        QuoteJob(Job.QUOTE,symbol, RefPri, OpenPri, HighPri, LowPri, MatchTime, MatchPri, MatchQty ,BestBuyPri, BestBuyQty, BestSellPri, BestSellQty)

    def OnGetDelayClose(self, this, symbol, DelayClose, ReqType):
        print ('OnGetDelayClose')

    def OnGetBreakResume(self, this, symbol, BreakTime, ResumeTime, ReqType):
        print ('OnGetBreakResume')

    def OnGetTradeStatus(self, this, symbol, TradeStatus, ReqType):
        print ('OnGetTradeStatus')

    def OnTickRegError(self, this, strSymbol, lMode, lErrCode, ReqType):
        print ('OnTickRegError')

    def OnGetTickData(self, this, strSymbol, strTickSn, strMatchTime, strBuyPri, strSellPri, strMatchPri, strMatchQty, strTolMatQty,
        strMatchAmt, strTolMatAmt, ReqType):
        print ('OnGetTickData')

    def OnTickRangeDataError(self, this, strSymbol, lErrCode, ReqType):
        print ('OnTickRangeDataError')

    def OnGetTickRangeData(self, this, strSymbol, strStartTime, strEndTime, strTolMatQty, strTolMatAmt, ReqType):
        print ('OnGetTickRangeData')

    def OnGetTimePack(self, this, strTradeType, strTime, ReqType):
        print ('OnGetTimePack {},{}'.format (strTradeType, strTime))

    def OnGetDelayOpen(self, this, symbol, DelayOpen, ReqType):
        print ('OnGetDelayOpen')

    def OnGetFutStatus(self, this, symbol, FunctionCode, BreakTime, StartTime, ReopenTime, ReqType):
        print ('OnGetFutStatus')

    def OnGetLimitChange(self, this, symbol, FunctionCode, StatusTime, Level, ExpandType, ReqType):
        print ('OnGetLimitChange')



class YuantaQuoteWapper:
    def __init__(self, handle, bot):
        self.bot = bot

        Iwindow = POINTER(IUnknown)()
        Icontrol = POINTER(IUnknown)()
        Ievent = POINTER(IUnknown)()

        res = atl.AtlAxCreateControlEx("YUANTAQUOTE.YuantaQuoteCtrl.1", handle, None,
                                    byref(Iwindow),
                                    byref(Icontrol),
                                    byref(GUID()),
                                    Ievent)

        self.YuantaQuote = GetBestInterface(Icontrol)
        self.YuantaQuoteEvents = YuantaQuoteEvents(self)
        self.YuantaQuoteEventsConnect = GetEvents(self.YuantaQuote, self.YuantaQuoteEvents)


class StockBot:
    def __init__(self, botuid):
        self.Yuanta = YuantaQuoteWapper (botuid, self)

    def login(self,account,password):
        # T port 80/443 , T+1 port 82/442 ,  reqType=1 T盤 , reqType=2  T+1盤        
        self.Yuanta.YuantaQuote.SetMktLogon(account, password, 'apiquote.yuantafutures.com.tw', '80', 1, 0) # 正式環境
        self.Yuanta.YuantaQuote.SetMktLogon(account, password, 'apiquote.yuantafutures.com.tw', '82', 2, 0)

        print ('login')

    def RegisterQuoteSymbol (self, QuoteSymbol,modle,ret_type):            
        ret = self.Yuanta.YuantaQuote.AddMktReg (QuoteSymbol, modle[0], ret_type, 0)
        print ("AddMktReg {}".format (ret))     #正常ret = 0        
        frame.SetQuoteStatusValue(str(ret))
        if ret == 0:
            InsertSymbol(Job.INSERTSYMBOL,QuoteSymbol)

    def UnRegisterQuoteSymbol (self, QuoteSymbol,ret_type):
        ret = self.Yuanta.YuantaQuote.DelMktReg (QuoteSymbol, ret_type)
        print ("DelMktReg {}".format (ret))     #正常ret = 0
        frame.SetQuoteStatusValue(str(ret))
        #if ret == 0:
        #    DeleteSymbol(Job.DELETESYMBOL,QuoteSymbol)


class MyApp(wx.App):
    def MainLoop(self, run_func):

        # Create an event loop and make it active.  If you are
        # only going to temporarily have a nested event loop then
        # you should get a reference to the old one and set it as
        # the active event loop when you are done with this one...
        evtloop = wx.GUIEventLoop()
        old = wx.EventLoop.GetActive()
        wx.EventLoop.SetActive(evtloop)

        # This outer loop determines when to exit the application,
        # for this example we let the main frame reset this flag
        # when it closes.
        while self.keepGoing:
            # At this point in the outer loop you could do
            # whatever you implemented your own MainLoop for.  It
            # should be quick and non-blocking, otherwise your GUI
            # will freeze.

            # call_your_code_here()
            run_func ()
            while not q.empty():
                next_job = q.get()
                DoJob (Bot, next_job)

            # This inner loop will process any GUI events
            # until there are no more waiting.
            while evtloop.Pending():
                evtloop.Dispatch()

            # Send idle events to idle handlers.  You may want to
            # throttle this back a bit somehow so there is not too
            # much CPU time spent in the idle handlers.  For this
            # example, I'll just snooze a little...
            time.sleep(0.10)
            evtloop.ProcessIdle()

        wx.EventLoop.SetActive(old)

    def OnInit(self):
        self.keepGoing = True
        return True


def run_job():
    while not q.empty():
        next_job = q.get()
        DoJob(Bot, next_job)



# In[5]:


if __name__ == "__main__":
    app=MyApp()    
    frame = AppFrame(None, title='YuantaQuoteAPI Sample',size = (670,370))
    frame.Show(True)
    Bot = StockBot(frame.Handle)    
    app.MainLoop(run_job)


# In[ ]:




