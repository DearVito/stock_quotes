import win32com.client
import yfinance as yf
import pandas as pd
tickers_list = ['RASP.ME', 'PLZL.ME', 'GMKN.ME', 'POLY.ME', 'ALNU.ME', 'ALRS.ME', 'NMTP.ME', 'SNGS.ME', 'NVTK.ME', 'SIBN.ME', 'GAZP.ME', 'ROSN.ME', 'RNFT.ME', 'JNOS.ME', 'TATN.ME', 'LKOH.ME', 'NLMK.ME', 'CHMF.ME', 'MGTS.ME', 'RTKM.ME', 'TTLK.ME', 'MTSS.ME', 'IRAO.ME', 'RSTI.ME', 'FEES.ME', 'UPRO.ME', 'OGKB.ME', 'IRGZ.ME', 'LSNG.ME', 'TGKA.ME', 'MRKP.ME', 'KUBE.ME', 'MRKU.ME', 'MRKZ.ME', 'PHOR.ME', 'AKRN.ME', 'OBUV.ME', 'GCHE.ME', 'KOGK.ME', 'LSRG.ME', 'TRMK.ME', 'CHEP.ME', 'QIWI.ME', 'GLTR.ME', 'MVID.ME', 'TCSG.ME', 'BELU.ME', 'PIKK.ME', 'MDMG.ME', 'SBER.ME', 'CBOM.ME', 'AVAN.ME', 'BSPB.ME', 'VTBR.ME', 'SFIN.ME', 'AQUA.ME']


Excel = win32com.client.Dispatch("Excel.Application")
wb = Excel.Workbooks.Open(r'C:\Users\Виталий\Desktop\STOCKS.xlsx')
sheet = wb.ActiveSheet

# Import pandas
data = pd.DataFrame(columns=tickers_list)

# Fetch the data
i = 1
X = input('Введите дату в формате: ГГГГ-ММ-ДД')
for ticker in tickers_list:
    data[ticker] = yf.download(ticker, X)['Adj Close']
    sheet.Cells(i,5).value = data[ticker]
    i += 1
# Print first 5 rows of the data
data.head()