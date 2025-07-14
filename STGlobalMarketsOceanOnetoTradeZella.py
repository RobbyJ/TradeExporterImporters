# This script will process the ST Global Markets (Ocean One) "Detailed" Excel (xls)
# report (created using stgmarkets.propreports.com).
# This will produce a CSV in the format TradeZella expects with the filename including
# the date from and to.
#
# Input: download.xls 
# Output: TradeZella-STGlobal-[STARTDATE]-[ENDDATE].csv
#
# Limitations: Currently works and is tested on Stocks only (not Options)
import os
import xlrd
from dateutil import parser

filename="download.xls"
fixedfilename="downloadcorrectedendian.xls"
expfilenameprefix = "TradeZella-STGlobal-"
filezelladefaultCSVheader = "Date,Time,Symbol,Buy/Sell,Quantity,Price,Spread,Expiration,Strike,Call/Put,Commission,Fees\n"

def FixLittleEndianMarker(infilename, outfilename):
    infile = open(infilename, "rb")
    content = infile.read()
    infile.close()
    content = content[:28]+b'\xFE\xFF'+content[30:]
    outfile = open(outfilename, 'wb')
    outfile.write(content)
    outfile.close()

def ExceltoFileZellaCSV(processFile):
    wb = xlrd.open_workbook(processFile)
    sh = wb.sheet_by_name('Sheet1')
    outputfile = open(expfilenameprefix, 'w', encoding='utf8')
    outputfile.write(filezelladefaultCSVheader)
    tradedate = '12/31/01'
    tradefirstdate = '12/31/01'
    tradetime = '01:01:01'
    tradesymbol = 'UNKNOWN'
    for rownum in range(sh.nrows):
        # Debug output
        # print(sh.row_values(rownum))

        # Checking for the new trading day date entry report row:
        if (sh.row_values(rownum)[0].find('/')!=-1):
            tradedatetime = parser.parse(sh.row_values(rownum)[0])
            tradedate = tradedatetime.strftime("%m/%d/%y")
            if (tradefirstdate == '12/31/01'):
                tradefirstdate = tradedate
        else:
            # Checking for a new trading symbol report row:
            if (sh.row_values(rownum)[0].find(' - ')!=-1):
                tradesymbol = sh.row_values(rownum)[0][0:sh.row_values(rownum)[0].find(' - ')]
            else:
                # Checking for a trade entry report row:
                if ((sh.row_values(rownum)[0].count(':')==2) and (len(sh.row_values(rownum)[0])==8)):
                    tradedatetime = parser.parse(sh.row_values(rownum)[0])
                    tradetime = tradedatetime.strftime("%H:%M:%S")
                    tradecommision = str(sh.row_values(rownum)[10])
                    tradefees = sh.row_values(rownum)[11]+sh.row_values(rownum)[12]+sh.row_values(rownum)[13]+\
                                sh.row_values(rownum)[14]+sh.row_values(rownum)[15]+sh.row_values(rownum)[16]+\
                                sh.row_values(rownum)[17]+sh.row_values(rownum)[18]+sh.row_values(rownum)[19]+\
                                sh.row_values(rownum)[20]
                    if (sh.row_values(rownum)[5]=='B'):
                        outputfile.write(tradedate+','+tradetime+','+tradesymbol+',Buy,' +\
                                         str(round(sh.row_values(rownum)[6]))+','+str(sh.row_values(rownum)[7])+\
                                         ',Stock,,,,'+tradecommision+','+str(tradefees)+'\n')
                    else:
                        outputfile.write(tradedate+','+tradetime+','+tradesymbol+',Sell,'+\
                                         str(round(sh.row_values(rownum)[6]))+','+str(sh.row_values(rownum)[7])+\
                                        ',Stock,,,,'+tradecommision+','+str(tradefees)+'\n')                        
    outputfile.close()
    os.rename(expfilenameprefix, expfilenameprefix+tradefirstdate.replace('/','')+'-'+tradedate.replace('/','')+'.csv' )

FixLittleEndianMarker(filename, fixedfilename)
ExceltoFileZellaCSV(fixedfilename)