#-*- coding: utf-8 -*-
'''

1. import list
2. tuning values
3. classes
    - Stock
        ucode(self, code)
        uname(self, name)
        uprice(self, price)
        uquant(self, quantity)
        ustren(self,strength)
        uressq(self,residual_sq)
        uresbq(self,residual_bq)
        ubtime(self,b_time)
        ubpric(self,b_price)
        uontra(self,on_trade)
        ubuyfl(self,buy_flag)
        update(self, dic)
        buy()
        sell()
    - XASessionEvents
    - XAQueryEvents
4. function list
    - login_process(demo = False)
    - stock_quotation(codes)
    - get_top_trade_cost(field = 1, day = 0)
    - get_top_trade_volume(field = 1, day = 0)
    - get_server_time()
    - parse(addr, ident, with_a = 1)
    - get_current_time()
    - starter(start_hour = 8, start_minute = 55)
    - checker(stock_object, end_hour = 15, end_minute = 35)
    - finisher()
    - signal_handler(signal, frame)
    - system_init()
    - pre_filter(i_temp, max_length = 50, min_price = 1000, max_price = 10000)
    - group_update(code_pointer, class_pointer, init = False, time_interval = None)
    - group_display_and_print(code_pointer, class_pointer, length = 3, with_print = True)
    - main()
5. main function

'''

import win32com.client
import pythoncom
import time
from datetime import datetime
import requests as rs
import bs4
import re
import sys
import signal
import numpy as np


# Tuning values are listed.
tun_val_01 = 1
tun_val_02 = 340
tun_val_03 = 30
tun_val_04 = 30
tun_val_05 = [35,100]
tun_val_06 = 30
tun_val_07 = 2
tun_val_08 = [97,100]
tun_val_09 = 400
tun_val_10 = 408
tun_val_11 = 884
tun_val_12 = 306
tun_val_13 = 0

min_leng = 1200

# For certification
certi_pass = ''
identification = ''
password = ''

class Stock:
    def __init__(self, code):
        self.code = code
        self.price = [None] * min_leng
        self.quantity = [None] * min_leng
        self.strength = [None] * min_leng
        self.name = None
        self.residual_sq = None
        self.residual_bq = None
        self.b_time = None
        self.b_price = None
        self.on_trade = False
        self.buy_flag = False
        self.maxOSC = -10000
    def ucode(self, code):
        self.code = code
    def uname(self, name):
        self.name = name
    def uprice(self, price):
        self.price.append(int(price))
        if self.price[0] is None:
            self.price.pop(0)
    def uquant(self, quantity):
        self.quantity.append(int(quantity))
        if self.quantity[0] is None:
            self.quantity.pop(0)
    def ustren(self,strength):
        self.strength.append(float(strength))
        if self.strength[0] is None:
            self.strength.pop(0)
    def uressq(self,residual_sq):
        self.residual_sq = int(residual_sq)
    def uresbq(self,residual_bq):
        self.residual_bq = int(residual_bq)
    def ubtime(self,b_time):
        self.b_time = b_time
    def ubpric(self,b_price):
        self.b_price = b_price
    def uontra(self,on_trade):
        self.on_trade = on_trade
    def ubuyfl(self,buy_flag):
        self.buy_flag = buy_flag
    def uMaxOSC(self,maxOSC):
        self.maxOSC = maxOSC
    funcmap = {
        1:ucode,2:uname,3:uprice,4:uquant,5:ustren,
        6:uressq,7:uresbq,8:ubtime,9:ubpric,10:uontra,
        11:ubuyfl,12:uMaxOSC
    }
    def update(self, dic):
        for key in dic.keys():
            self.funcmap[key](self,dic[key])
            
    def buy():
        pass
    def sell():
        pass
    
    def MACD(self, d):
        if not(self.price[0] is None):
            if d == 0:
                return np.mean(self.price[-tun_val_10:]) - np.mean(self.price[-tun_val_11:])
            if d > 0:
                return np.mean(self.price[-tun_val_10-d:-d]) - np.mean(self.price[-tun_val_11-d:-d])
        else:
            return None
    @property
    def signal(self):
        if not(self.price[0] is None):
            return np.mean([self.MACD(d = l) for l in range(tun_val_12)])
        else:
            return None
    @property
    def OSC(self):
        if not(self.price[0] is None):
            return self.MACD(d=0) - self.signal
        else:
            return None
    
class XASessionEvents:
    logInState = 0
    def OnLogin(self, code, msg):
        print("OnLogin method is called")
        print(str(code))
        print(msg)
        if str(code) == '0000':
            XASessionEvents.logInState = 1

    def OnLogout(self):
        print("OnLogout method is called")

    def OnDisconnect(self):
        print("OnDisconnect method is called")

class XAQueryEvents:
    queryState = 0
    def OnReceiveData(self, szTrCode):
        print("ReceiveData")
        XAQueryEvents.queryState = 1
    def OnReceiveMessage(self, systemError, mesageCode, message):
        print("ReceiveMessage")
        
def login_process(demo = False):
    if demo:
        server_addr = "demo.ebestsec.co.kr"
        user_certificate_pass = None
    else:
        server_addr = "hts.ebestsec.co.kr"
        user_certificate_pass = certi_pass
    
    server_port = 20001
    server_type = 0
    user_id = identification
    user_pass = password

    inXASession = win32com.client.DispatchWithEvents("XA_Session.XASession", XASessionEvents)
    inXASession.ConnectServer(server_addr, server_port)
    inXASession.Login(user_id, user_pass, user_certificate_pass, server_type, 0)

    while XASessionEvents.logInState == 0:
        time.sleep(0.01)
        pythoncom.PumpWaitingMessages()

def stock_quotation(codes):
    number = len(codes)
    concat_list = ''.join(codes)
    inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    inXAQuery.LoadFromResFile("C:\\eBest\\xingAPI\\Res\\t8407.res")
    inXAQuery.SetFieldData('t8407InBlock', 'nrec', 0, number)
    inXAQuery.SetFieldData('t8407InBlock', 'shcode', 0, concat_list)
    inXAQuery.Request(0)

    while XAQueryEvents.queryState == 0:
        time.sleep(0.01)
        pythoncom.PumpWaitingMessages()

    result0, result1 = [], []
    nCount =inXAQuery.GetBlockCount('t8407OutBlock1')
    for i in range(nCount):
        op01 = inXAQuery.GetFieldData('t8407OutBlock1', 'shcode', i)
        op02 = inXAQuery.GetFieldData('t8407OutBlock1', 'hname', i)
        op03 = inXAQuery.GetFieldData('t8407OutBlock1', 'price', i)
        op04 = inXAQuery.GetFieldData('t8407OutBlock1', 'sign', i)
        op05 = inXAQuery.GetFieldData('t8407OutBlock1', 'change', i)
        op06 = inXAQuery.GetFieldData('t8407OutBlock1', 'diff', i)
        op07 = inXAQuery.GetFieldData('t8407OutBlock1', 'volume', i)
        op08 = inXAQuery.GetFieldData('t8407OutBlock1', 'offerho', i)
        op09 = inXAQuery.GetFieldData('t8407OutBlock1', 'bidho', i)
        op10 = inXAQuery.GetFieldData('t8407OutBlock1', 'cvolume', i)
        op11 = inXAQuery.GetFieldData('t8407OutBlock1', 'chdegree', i)
        op12 = inXAQuery.GetFieldData('t8407OutBlock1', 'open', i)
        op13 = inXAQuery.GetFieldData('t8407OutBlock1', 'high', i)
        op14 = inXAQuery.GetFieldData('t8407OutBlock1', 'low', i)
        op15 = inXAQuery.GetFieldData('t8407OutBlock1', 'value', i)
        op16 = inXAQuery.GetFieldData('t8407OutBlock1', 'offerrem', i)
        op17 = inXAQuery.GetFieldData('t8407OutBlock1', 'bidrem', i)
        op18 = inXAQuery.GetFieldData('t8407OutBlock1', 'totofferrem', i)
        op19 = inXAQuery.GetFieldData('t8407OutBlock1', 'totbidrem', i)
        op20 = inXAQuery.GetFieldData('t8407OutBlock1', 'jnilclose', i)
        op21 = inXAQuery.GetFieldData('t8407OutBlock1', 'uplmtprice', i)
        op22 = inXAQuery.GetFieldData('t8407OutBlock1', 'dnlmtprice', i)
        result0.append([op01,op02])
        result1.append([
            op03,op04,op05,op06,op07,
            op08,op09,op10,op11,op12,
            op13,op14,op15,op16,op17,
            op18,op19,op20,op21,op22])
    XAQueryEvents.queryState = 0
    return result0, result1

def get_top_trade_cost(field = 1, day = 0):
    time.sleep(1)
    inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    inXAQuery.LoadFromResFile("C:\\eBest\\xingAPI\\Res\\t1463.res")
    inXAQuery.SetFieldData('t1463InBlock', 'gubun', 0, field)
    inXAQuery.SetFieldData('t1463InBlock', 'jnilgubun', 0, day)
    inXAQuery.Request(0)

    while XAQueryEvents.queryState == 0:
        time.sleep(0.01)
        pythoncom.PumpWaitingMessages()

    nCount =inXAQuery.GetBlockCount('t1463OutBlock1')
    result = []
    for i in range(nCount):
        result.append(inXAQuery.GetFieldData('t1463OutBlock1', 'shcode', i))
    XAQueryEvents.queryState = 0
    return result

def get_top_trade_volume(field = 1, day = 0):
    time.sleep(1)
    inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    inXAQuery.LoadFromResFile("C:\\eBest\\xingAPI\\Res\\t1452.res")
    inXAQuery.SetFieldData('t1452InBlock', 'gubun', 0, field)
    inXAQuery.SetFieldData('t1452InBlock', 'jnilgubun', 0, day)
    inXAQuery.Request(0)

    while XAQueryEvents.queryState == 0:
        time.sleep(0.01)
        pythoncom.PumpWaitingMessages()

    nCount =inXAQuery.GetBlockCount('t1452OutBlock1')
    result = []
    for i in range(nCount):
        result.append(inXAQuery.GetFieldData('t1452OutBlock1', 'shcode', i))
    XAQueryEvents.queryState = 0
    return result
    
def get_server_time():
    inXAQuery = win32com.client.DispatchWithEvents("XA_DataSet.XAQuery", XAQueryEvents)
    inXAQuery.LoadFromResFile("C:\\eBest\\xingAPI\\Res\\t0167.res")
    inXAQuery.Request(0)

    while XAQueryEvents.queryState == 0:
        time.sleep(0.01)
        pythoncom.PumpWaitingMessages()
        
    dt = inXAQuery.GetFieldData('t0167OutBlock', 'dt', 0)
    tt = inXAQuery.GetFieldData('t0167OutBlock', 'time', 0)
    XAQueryEvents.queryState = 0
    return dt, tt

def parse(addr, ident, with_a = 1):
    response = rs.get(addr)
    html_content = response.text.encode(response.encoding)
    navigator = bs4.BeautifulSoup(html_content, 'lxml')
    html = navigator.find_all(href=re.compile(ident))
    return_list = []
    for line in html:
        r = re.search('code\=(.+?)\"', str(line))
        if r:
            if with_a:
                return_list.append("A"+r.group(1))
            else:
                return_list.append(r.group(1))
    return return_list

def get_current_time():
    return datetime.now()

def starter(start_hour = 8, start_minute = 55):
    while 1:
        current_time = datetime.now()
        h, m = current_time.hour, current_time.minute
        if (((h == start_hour)&(m >= start_minute))|(h > start_hour)):
            break
        time.sleep(5)
        
    #create file with "current time" name.
    current_time = datetime.now()
    s = "%04d_%02d_%02d_%02d_%02d_%02d" % (current_time.year,
                                           current_time.month,
                                           current_time.day,
                                           current_time.hour,
                                           current_time.minute,
                                           current_time.microsecond)
    
    global f, f_trade
    f = open(s+'.txt','w')
    f_trade = open(s+'_trade.txt','w')

def checker(stock_object, end_hour = 15, end_minute = 35):
    current_time = datetime.now()
    h, m = current_time.hour, current_time.minute
    
        #selling logic
    for stock in stock_object:
        if ((stock.buy_flag)&(not(stock.on_trade))):
            det1 = tun_val_10[0] * stock.b_price >= tun_val_10[1] * stock.price[-1]
            det2 = (400 * (max(1) - 1) >= 1)
            det3 = (1 <= 0)
            if(any([det1,det2,det3])):
                pass

            #buying logic
    for stock in stock_object:
        if not(stock.price[0] is None):
            if (not(stock.buy_flag)&(not(stock.on_trade))):
                det1 = 100 * (
                    (stock.quantity[-1] - stock.quantity[-tun_val_03]) / tun_val_03 
                    - (stock.quantity[-tun_val_04] - stock.quantity[-tun_val_04-tun_val_03]) / tun_val_03
                ) > tun_val_01 * (stock.quantity[-1] - stock.quantity[-tun_val_02])
                det2 = tun_val_05[1] * (
                    np.mean(stock.strength[-tun_val_06:]) - np.mean(stock.strength[-2*tun_val_06:-tun_val_06])
                ) > tun_val_05[0]
                det3 = stock.residual_sq[-1] > tun_val_09 * stock.residual_bq[-1]
                if(all([det1,det2,det3])):
                    pass
                
    return not(((h == end_hour)&(m >= end_minute))|(h > end_hour))

def finisher():
    f.close()
    f_trade.close()
    
def signal_handler(signal, frame):
    finisher()
    sys.exit(0)
    
def system_init():
    signal.signal(signal.SIGINT, signal_handler)
    
def pre_filter(i_temp, max_length = 50, min_price = 1000, max_price = 10000):
    iteration = int(len(i_temp)/max_length) + (len(i_temp)%max_length > 0)
    result = []
    for j in range(iteration):
        if j == (iteration - 1):
            nfrom = j * max_length
            
            temp0, temp1 = stock_quotation(i_temp[nfrom:])
        else:
            nfrom = j * max_length
            nto = (j+1) * max_length
            temp0, temp1 = stock_quotation(i_temp[nfrom:nto])
        temp0 = np.array(temp0)
        temp1 = np.asarray(temp1, dtype=np.float32)
        i = 0
        temp_result = []
        for p in temp1[:,0]:
            det1 = (int(p) > min_price)
            det2 = (int(p) < max_price)
            if det1 & det2:
                temp_result.append(i)
            i = i + 1
        for i in temp_result:
            result.append(temp0[i,0])
    result = np.array(result).tolist()
    if (len(result) > 50):
        result = result[:50]
    return result

def group_update(code_pointer, class_pointer, init = False, time_interval = None):
    if time_interval is None:
        pass
    else:
        time.sleep(time_interval)
    r0, r1 = stock_quotation(code_pointer)
    i = 0
    if init:
        while i < len(code_pointer):
            dic = {2:r0[i][1], 3:r1[i][0], 4:r1[i][4], 5:r1[i][8], 6:r1[i][15], 7:r1[i][16]}
            class_pointer[i].update(dic)
            f.write(str(r0[i][1].encode('euc-kr')) + '|')
            i = i + 1
        f.write('\n')
    else:
        while i < len(code_pointer):
            dic = {3:r1[i][0], 4:r1[i][4], 5:r1[i][8], 6:r1[i][15], 7:r1[i][16]}
            class_pointer[i].update(dic)
            i = i + 1

def group_display_and_print(code_pointer, class_pointer, length = 3, with_print = True):
    i = 0
    for code in code_pointer:
        print i+1
        print 'code: ' + class_pointer[i].code
        print 'name: ' + class_pointer[i].name
        print 'price: ', 
        print class_pointer[i].price[-length:]
        print 'quantity: ',
        print class_pointer[i].quantity[-length:]
        print 'trade_strength: ',
        print class_pointer[i].strength[-length:]
        print 'residual_sell_qunt: ',
        print class_pointer[i].residual_sq
        print 'residual_buy_qunt: ',
        print class_pointer[i].residual_bq
        print '\n'
        i = i + 1
    print get_current_time()
    i = 0
    if with_print:
        for code in code_pointer:
            s1 = str(class_pointer[i].price[-1])
            s2 = str(class_pointer[i].quantity[-1])
            s3 = str(class_pointer[i].strength[-1])
            s4 = str(class_pointer[i].residual_sq)
            s5 = str(class_pointer[i].residual_bq)
            f.write(s1+'|'+s2+'|'+s4+'|'+s5+'|'+s3+'|')
            i = i + 1
        s = str(get_current_time())
        f.write(s+'\n')
    
def main():
    #system initiate
    system_init()

    #trade login
    login_process()

    #start process
    starter()

    #code listing & filtering
    l = [None] * 9
    l[0] = parse(
        'http://finance.naver.com/sise/lastsearch2.nhn',
        '/item/main.nhn\?code=',
        with_a = 0
    )
    l[1] = get_top_trade_cost(field = 1,day = 0)
    l[2] = get_top_trade_cost(field = 2,day = 0)
    l[3] = get_top_trade_volume(field = 1,day = 0)
    l[4] = get_top_trade_volume(field = 2,day = 0)
    l[5] = get_top_trade_cost(field = 1,day = 1)
    l[6] = get_top_trade_cost(field = 2,day = 1)
    l[7] = get_top_trade_volume(field = 1,day = 1)
    l[8] = get_top_trade_volume(field = 2,day = 1)
    l_sum = []
    for i in range(9):
        l_sum = l_sum + l[i]

    codes = list(set(l_sum))
    codes = pre_filter(codes)

    #generate object
    stocks = [None] * len(codes)
    i = 0
    for code in codes:
        stocks[i] = Stock(code)
        i = i + 1

    #object initialize
    group_update(codes, stocks, init = True, time_interval = 1)
    group_display_and_print(codes, stocks)

    #trade and archive
    while checker(stock_object = stocks):
        group_update(codes, stocks, time_interval = 1)
        group_display_and_print(codes, stocks)

    finisher()
    
if __name__=='__main__':
    main()
