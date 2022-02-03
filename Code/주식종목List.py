import win32com.client
from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Stocks"


# 연결 여부 체크
objCpCybos = win32com.client.Dispatch("CpUtil.CpCybos")
bConnect = objCpCybos.IsConnect
if (bConnect == 0):
    ws.append(["PLUS가 정상적으로 연결되지 않음. "])
    exit()

# 종목코드 리스트 구하기
objCpCodeMgr = win32com.client.Dispatch("CpUtil.CpCodeMgr")
codeList = objCpCodeMgr.GetStockListByMarket(1)  # 거래소
codeList2 = objCpCodeMgr.GetStockListByMarket(2)  # 코스닥


ws.append(["거래소 종목코드", len(codeList)])
for i, code in enumerate(codeList):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    ws.append([i, code, secondCode, stdPrice, name])

ws.append(["코스닥 종목코드", len(codeList2)])
for i, code in enumerate(codeList2):
    secondCode = objCpCodeMgr.GetStockSectionKind(code)
    name = objCpCodeMgr.CodeToName(code)
    stdPrice = objCpCodeMgr.GetStockStdPrice(code)
    ws.append([i, code, secondCode, stdPrice, name])

ws.append(["거래소 + 코스닥 종목코드 ", len(codeList) + len(codeList2)])

wb.save("주식종목List.xlsx")
