#-*- coding: utf-8 -*-

import numpy as np
import pandas as pd
import win32com.client
import time
import signal
import sys


def signal_handler(signal, frame):
    finisher()
    sys.exit(0)
signal.signal(signal.SIGINT, signal_handler)

inCpTdUtil = win32com.client.Dispatch("CpTrade.CpTdUtil")
inCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
inCpStockCode = win32com.client.Dispatch("CpUtil.CpStockCode")
inCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
inStockChart = win32com.client.Dispatch("CpSysdib.StockChart")

def parser(addr, ident):
    import requests as rs
    import bs4
    import re
    response = rs.get(addr)
    html_content = response.text.encode(response.encoding)
    navigator = bs4.BeautifulSoup(html_content)
    html = navigator.find_all(href=re.compile(ident))
    return_list = []
    for line in html:
        r = re.search('code\=(.+?)\"', str(line))
        if r:
            return_list.append("A"+r.group(1))
    return return_list

def all_code_list():
    inCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
    ind_code_number = np.array(list(inCpCodeMgr.GetIndustryList()))
    raw_codes = []
    for i in ind_code_number:
        raw_codes.append(list(inCpCodeMgr.GetGroupCodeList(i)))
    result = []
    for codes in raw_codes:
        for code in codes:
            result.append(code)
    return result

def preprocessing(code_list):
    num_of_lines = int(round(len(code_list)/110+0.5))
    new_code_list = list(code_list)
    new_code_list = np.array(new_code_list)
    new_code_list.resize(num_of_lines,110)
    temp = []
    result = []
    for i in new_code_list:
        temp.append(stock_quotation(i, quotation_options = [0,3]))
    for i in np.array(temp).reshape(-1,2):
        price = int(i[1])
        if (price > 1000) & (price < 10000):
            result.append(i)
    result = np.array(result)
    if len(result) > 110:
        result = list(result[:110].transpose()[0])
    else:
        result = list(result.transpose()[0])
    return result

def cybos_check_account(account_number):
    inCpTd6033 = win32com.client.Dispatch("CpTrade.CpTd6033")
    retArr = {}
    inCpTd6033.SetInputValue(0, account_number)
    retval = inCpTd6033.BlockRequest()
    received_code_number = inCpTd6033.GetHeaderValue(7)
    for i in range(received_code_number):
        code = inCpTd6033.GetDataValue(12,i)
        quantity = inCpTd6033.GetDataValue(3,i)
        price = inCpTd6033.GetDataValue(4,i)
        bh, bm = 0, 0
        retArr[code] = [quantity, price, bh, bm]
    return retArr

def cybos_buy(code, quantity):
    inCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311")
    inCpTd0311.SetInputValue(0, '2') #sell:1, buy:2
    inCpTd0311.SetInputValue(1, "335249152")
    inCpTd0311.SetInputValue(2, '10')
    inCpTd0311.SetInputValue(3, code) #code
    inCpTd0311.SetInputValue(4, quantity) #quantity
    inCpTd0311.SetInputValue(5, 0)
    inCpTd0311.SetInputValue(8, "03") #market price
    inCpTd0311.BlockRequest()

def cybos_sell(code, quantity):
    inCpTd0311 = win32com.client.Dispatch("CpTrade.CpTd0311")
    inCpTd0311.SetInputValue(0, '1') #sell:1, buy:2
    inCpTd0311.SetInputValue(1, "335249152")
    inCpTd0311.SetInputValue(2, '10')
    inCpTd0311.SetInputValue(3, code) #code
    inCpTd0311.SetInputValue(4, quantity) #quantity
    inCpTd0311.SetInputValue(5, 0)
    inCpTd0311.SetInputValue(8, "03") #market price
    inCpTd0311.BlockRequest()

def stock_quotation(codes, quotation_options = [0,3,11,13,14,21,22]):
    inStockMst2 = win32com.client.Dispatch("dscbo1.StockMst2")
    '''
    0 - (string) 종목 코드
    1 - (string) 종목명
    2 - (long) 시간(HHMM)
    3 - (long) 현재가
    4 - (long) 전일대비
    5 - (char) 상태구분 
    6 - (long) 시가
    7 - (long) 고가
    8 - (long) 저가
    9 - (long) 매도호가
    10 - (long) 매수호가
    11 - (unsigned long) 거래량 [주의] 단위 1주
    12 - (long) 거래대금 [주의] 단위 천원
    13 - (long) 총매도잔량
    14 - (long) 총매수잔량
    15 - (long) 매도잔량
    16 - (long) 매수잔량
    17 - (unsigned long) 상장주식수
    18 - (long) 외국인보유비율(%)
    19 - (long) 전일종가
    20 - (unsigned long) 전일거래량
    21 - (long) 체결강도
    22 - (unsigned long) 순간체결량
    23 - (char) 체결가비교 Flag
    24 - (char) 호가비교 Flag
    25- (char) 동시호가구분
    26 - (long) 예상체결가
    27 - (long) 예상체결가 전일대비
    28 - (long) 예상체결가 상태구분
    29- (unsigned long) 예상체결가 거래량
    '''
    retArr = []
    str_concat = ""
    for strs in codes:
        str_concat += strs
        str_concat +=","
    str_concat = str_concat[:-1]
    inStockMst2.SetInputValue(0, str_concat)
    retval = inStockMst2.BlockRequest()
    for i in range(len(codes)):
        result_options = []
        for j in quotation_options:
            result_options.append(inStockMst2.GetDataValue(j,i))
        retArr.append(result_options)
    return retArr

def transform_list(retArr):
    code_list = [list(i) for i in zip(*retArr)]
    return code_list

def createTableIndex(code_list):
    first = [x for pair in zip(code_list[0],code_list[0],code_list[0],code_list[0],code_list[0],code_list[0]) for x in pair]
    second = ['price','quantity','op1','op2','op3','op4']*len(code_list[0])
    tuples = list(zip(first,second)) + [('time','time')]
    index = pd.MultiIndex.from_tuples(tuples, names=['Code', 'Class'])
    return index

def makeElement(code_list):
    now = time.localtime()
    s = "%04d-%02d-%02d %02d:%02d:%02d" % (now.tm_year,now.tm_mon,now.tm_mday,now.tm_hour,now.tm_min,now.tm_sec)
    element = [x for pair in zip(code_list[1],code_list[2],code_list[3],code_list[4],code_list[5],code_list[6]) for x in pair] + [s]
    return element
    
def verifyCode(codes):
    s = "\n<CODE VALIDATING>"
    print s
    f.write(s)
    i = 0
    kill_list = []
    for code in codes:
        #GetStockSupervisionKind
        #GetStockStatusKind
        s = str(i) + " | " + str(inCpCodeMgr.GetStockStatusKind(code)) + " | " + inCpStockCode.CodeToName(code)
        s = s.encode('utf-8')
        print s.decode('utf-8')
        j()
        f.write(s)
        if (inCpStockCode.CodeToName(code) == '')|(inCpCodeMgr.GetStockStatusKind(code)):
            kill_list.append(i)
        i += 1
    for x in kill_list[::-1]:
        codes.pop(x)

def printCode(codes):
    s = "\n<CODE PRINTING>"
    print s
    f.write(s)
    i = 0
    for code in codes:
        s = str(i) + " | " + code + " | " + inCpStockCode.CodeToName(code)
        s = s.encode('utf-8')
        print s.decode('utf-8')
        j()
        f.write(s)
        i += 1

def j():
    s = '\n'
    f.write(s)

def starter(set_hour = 8, set_minute = 55):
    while 1:
        now = time.localtime()
        h, m = now.tm_hour, now.tm_min
        if (((h == set_hour)&(m >= set_minute))|(h > set_hour)):
            break
        time.sleep(5)

def checker(data_frame, wallet,rh = 15, rm = 5, bsh = 9, bsm = 30, beh = 14, bem = 30):
    now = time.localtime()
    h, m = now.tm_hour, now.tm_min
    print wallet
    if (((h == bsh)&(m >= bsm))|(h > bsh)) & (not(((h == beh)&(m >= bem))|(h > beh))):
        code_list = []
        number_of_codes = int(data_frame.shape[0]/6)
        index = list(data_frame.transpose())
        for i in range(number_of_codes):
            code_list.append(index[6*i][0])
        # selling logic
        for i in wallet.keys():
            for j in range(number_of_codes):
                if (i == code_list[j]):
                    ratio = wallet[i][1] / data_frame.iloc[6*j,-1:].astype(np.float).iloc[0]
                    print 'ratio: ' + str(ratio)
                    if (ratio > 1.02) | (ratio < 0.97): 
                        cybos_sell(i,1)
                        del wallet[i]
        # buying logic
        print data_frame.shape[1]
        if (data_frame.shape[1] > 350):
            for i in range(number_of_codes):
                price = data_frame.iloc[6*i,-1:].astype(np.float).iloc[0]
                if (price > 1000) & (price < 10000): 
                    t_0 = data_frame.iloc[1+6*i,-1:].astype(np.float).iloc[0]
                    t_30 = data_frame.iloc[1+6*i,-30:].astype(np.float).iloc[0]
                    t_60 = data_frame.iloc[1+6*i,-60:].astype(np.float).iloc[0]
                    t_340 = data_frame.iloc[1+6*i,-340:].astype(np.float).iloc[0]
                    det = (t_0 - 2 * t_30 + t_60) / (0.3 * (t_0 - t_340))
                    print det
                    if (det > 1.5) & (not(code_list[i] in wallet)):
                        cybos_buy(code_list[i],1)
                        wallet[code_list[i]] = [1, price, h, m]
                        #s = code + 'stock bought'
                        #print(s)
                        #f.write(s)
    else:
        dump_all(wallet)
    if (((h == rh)&(m >= rm))|(h > rh)):
        finisher()
    return not(((h == rh)&(m >= rm))|(h > rh))

def finisher():
    print('END PROCESS')
    f.close()
    
def dump_all(wallet):
    for k in wallet.keys():
        cybos_sell(k, wallet[k][0])
    wallet.clear()
    
def main():
        
    #init
    inCpTdUtil.TradeInit(0)

    starter()
            
    my_wallet = cybos_check_account("335249152")
    dump_all(my_wallet)

    now = time.localtime()
    s = "%04d_%02d_%02d_%02d_%02d_%02d" % (now.tm_year,now.tm_mon,now.tm_mday,now.tm_hour,now.tm_min,now.tm_sec)
    f = open(s+'.txt','w')

    l1 = parser('http://finance.naver.com/sise/lastsearch2.nhn','/item/main.nhn\?code=')
    l2 = parser('http://finance.naver.com/sise/sise_quant.nhn?sosok=1','/item/main.nhn\?code=')
    l3 = parser('http://finance.naver.com/sise/sise_quant.nhn?sosok=0','/item/main.nhn\?code=')
    tl = list(set(l1+l2+l3))
    #tl = all_code_list()
    sc = preprocessing(tl)

    printCode(sc)
    verifyCode(sc)
    printCode(sc)

    cl = transform_list(stock_quotation(sc))
    ind = createTableIndex(cl)
    e = makeElement(cl)
    df = pd.DataFrame(e,index=ind)

    j()
    str_list = []
    for i in range(len(df[0])):
        str_list.append(str(df[0][i]))
        str_list.append('|')

    s = ''.join(str_list)
    f.write(s)
    print df

    i = 0
    while checker(data_frame = df, wallet = my_wallet):
        time.sleep(1.5)
        i += 1
        cl = transform_list(stock_quotation(sc))
        e = makeElement(cl)
        df[i] = e
        print df.iloc[:,-2:]
    
        j()
        str_list = []
        for x in range(len(df[i])):
            str_list.append(str(df[i][x]))
            str_list.append('|')

        s = ''.join(str_list)
        f.write(s)

if __name__ == "__main__":
    main()