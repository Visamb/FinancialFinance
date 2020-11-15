import pandas as pd
from numpy import*
import matplotlib.pyplot as plt
from astropy.table import Table

#Import data
Stocks = pd.read_excel(r'C:\Users\Viktor\Desktop\Finansiell ekonomi\OMXS30.xlsx', sheet_name = 0)

Stockprices = Stocks['X / SEK']
OMXS30 = Stocks['OMXS30']

def R(prices):
    length = len(prices)-1
    Rvector = []
    for i in range(length):
        R = (prices.iloc[i+1]-prices.iloc[i])/prices.iloc[i]
        Rvector.append(R)
    return Rvector

RStocks = R(Stockprices)
ROMX = R(OMXS30)
Dates = Stocks['Datum'].loc[1:]
df = pd.DataFrame()
df['Dates'] = Dates
df['R_OMX'] = ROMX
df['R_Stocks'] = RStocks

#PART2_PLOTS
"""df.plot(x = 'Dates',y = 'R_OMX')
df.plot(x = 'Dates',y = 'R_Stocks')
df.hist(column = 'R_OMX')
df.hist(column = 'R_Stocks')
Stocks.plot(x = 'Datum', y = 'OMXS30' )
Stocks.plot(x = 'Datum', y = 'X / SEK')"""
#plt.show()

#PART3_DESCRIPTIVE_STATISTICS
MAX_OMX = max(df['R_OMX'])
MAX_Stocks = max(df['R_Stocks'])
MIN_OMX = min(df['R_OMX'])
MIN_Stocks = min(df['R_Stocks'])
MED_OMX = float(median(df['R_OMX']))
MED_Stocks = float(median(df['R_Stocks']))
MEAN_OMX = mean(df['R_OMX'])
MEAN_Stocks = mean(df['R_Stocks'])
STD_OMX = std(df['R_OMX'])
STD_Stocks = std(df['R_Stocks'])
SKEWNESS_OMX = mean(((df['R_OMX']-MEAN_OMX)**3)/(STD_OMX**3))
SKEWNESS_Stocks = mean(((df['R_Stocks']-MEAN_Stocks)**3)/(STD_Stocks**3))
KURTOSIS_OMX = mean(((df['R_OMX']-MEAN_OMX)**4)/(STD_OMX**4))
KURTOSIS_Stocks = mean(((df['R_Stocks']-MEAN_Stocks)**4)/(STD_Stocks**4))

st = Table()
st['DATA'] = ['OMX', 'Stock X']
st['MEAN'] = [MEAN_OMX,MEAN_Stocks]
st['STD'] = [STD_OMX,STD_Stocks]
st['SKEWNESS'] = [SKEWNESS_OMX,SKEWNESS_Stocks]
st['KURTOSIS'] = [KURTOSIS_OMX,KURTOSIS_Stocks]
#OMXS30 is obs more like normal distribution. Higher risk in stock X
print(st)

#PART4 obligations
def ytm(FV,C,n,y,marketprice):
    n = n
    y = linspace(-0.1,0.1,10000000)
    P = (C / y) * (1 - (power((1 + y), -n))) + (FV) / (power((1 + y), n))
    Thisval = where((abs(marketprice - P) == min(abs(marketprice - P))))
    print(min(abs(marketprice - P)))
    #print(Thisval)
    yie = (y[Thisval])
    return yie

Bonds = pd.read_excel(r'C:\Users\Viktor\Desktop\Finansiell ekonomi\OMXS30.xlsx', sheet_name = 1, skiprows=4)
dfb = pd.DataFrame()
dfb['Market val'] = Bonds['Marknadsvärde']
dfb['FV'] = Bonds['Nominellt belopp']
dfb['n'] = Bonds['Löptid']
dfb['Start yield'] = Bonds['Kupongränta']/100
dfb['C'] = dfb['FV']*dfb['Start yield']

ytmlist = []
for num in range(8):
    a = (ytm(dfb['FV'].iloc[num],dfb['C'].iloc[num],dfb['n'].iloc[num],dfb['Start yield'].iloc[num], dfb['Market val'].iloc[num]))
    ytmlist.append(a)
print(ytmlist)

nlist = [dfb['n'].iloc[i] for i in range(8)]
plt.scatter(nlist,ytmlist)
plt.show()



