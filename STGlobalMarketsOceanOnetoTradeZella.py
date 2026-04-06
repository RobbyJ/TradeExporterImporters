# This script will process the ST Global Markets (Ocean One) "Detailed" Excel (xls)
# report created using stgmarkets.propreports.com (make sure you select Detailed from the drop down).
#
# IMPORTANT: in TradeZella you must set up Ocean One as a LightSpeed Broker type being importing a file.
#
# This tool will produce a CSV in the format TradeZella expects using Lightspeed as the Broker type.
#
# The filename includes the date from and to that trades were found for.
#
# Use the command as follows running from the commmand-line in the folder with this file and exported xls file:
# Command: python .\STGlobalMarketsOceanOnetoTradeZella.py 123456-2026-04-01-to-2026-04-04-detailed.xls
# Input: 123456-2026-04-01-to-2026-04-04-detailed.xls
# Output: TradeZella-STGlobal-[STARTDATE]-[ENDDATE].csv
#
# Limitations: Currently works and is tested on Stocks only (not Options), and only tested on Windows 11
import os
import sys
import xlrd
from dateutil import parser

accountnumber = "1234546"
filename="testfile.xls"
fixedfilename="downloadcorrectedendian.xls"
expfilenameprefix = "TradeZella-STGlobal-"
filezelladefaultCSVheader = '"Account Number","Account Type","Side","Symbol","CUSIP","Currency Code","Security Type","Buy/Sell","Trade Date","Settlement Date","Process Date","Price","Qty","Trade Number","Principal Amount","NET Amount","Commission Amount","Execution Time","Raw Exec. Time","Market Code","Trailer","FeeSEC","FeeMF","Fee1","Fee2","Fee3","FeeStamp","FeeTAF","Fee4","Sequence Number","Side Seq Code","Capacity Code","Office Code","Rep Code","Special Code","Instructions Trade Legend Code","Factor Type2","Trade Interest","Original TradeNumber","Entry Time","Entered By","YieldToMature","YieldToCall","Mutual Fund Sales Charge Rate","Mutual Fund Load Indicator","Transtype"\n'

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
    checkfile = open("checkfile.csv", 'w', encoding='utf8')
    tradedate = '12/31/1901'
    tradefirstdate = '12/31/1901'
    tradetime = '01:01:01'
    tradesymbol = 'UNKNOWN'
    for rownum in range(sh.nrows):
        # Debug output
        # print(sh.row_values(rownum))
        checkfile.write(str(sh.row_values(rownum))+'\n')

        # Checking for the new trading day date entry report row:
        if (sh.row_values(rownum)[0].count('/')==2):
            tradedatetime = parser.parse(sh.row_values(rownum)[0])
            tradedate = tradedatetime.strftime("%m/%d/%Y")
            if (tradefirstdate == '12/31/1901'):
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
                    tradequantity = str(round(sh.row_values(rownum)[6]))
                    tradeprice = str(sh.row_values(rownum)[7])
                    tradeprincipleamount = str(sh.row_values(rownum)[6] * sh.row_values(rownum)[7])
                    tradenetamount = str((sh.row_values(rownum)[6] * sh.row_values(rownum)[7]) - sh.row_values(rownum)[21])                                                                                 
                    tradecommission = str(sh.row_values(rownum)[10])
                    # tradeecnfee = str(sh.row_values(rownum)[11])
                    # tradeorffee = str(sh.row_values(rownum)[13])
                    tradensccfee = str(sh.row_values(rownum)[11] + sh.row_values(rownum)[13] + sh.row_values(rownum)[17])
                    tradesecfee = str(sh.row_values(rownum)[12])
                    tradecatfee = str(sh.row_values(rownum)[14])
                    tradetaffee = str(sh.row_values(rownum)[15])
                    tradeoccfee = str(sh.row_values(rownum)[16])
                    tradeaccfee = str(sh.row_values(rownum)[18])
                    tradeclrfee = str(sh.row_values(rownum)[19])
                    trademiscfee = str(sh.row_values(rownum)[20])
                    tradenetfee = str(sh.row_values(rownum)[21])

                    outputfile.write("\"" + accountnumber + "\",\"1\",") # Account Number, Account Type
                    if (sh.row_values(rownum)[5]=='B'):
                        outputfile.write('"B",') # Side
                        outputfile.write('"'+tradesymbol+'",') #Symbol
                        outputfile.write('"A111111",')   # CUSIP
                        outputfile.write('"USD",')   #Currency Code
                        outputfile.write('"equity",') # Security Type
                        outputfile.write('"Long Buy",') # Buy/Sell
                        outputfile.write('"'+tradedate+'",') # Trade Date
                        outputfile.write('"'+tradedate+'",') # Settlement Date
                        outputfile.write('"'+tradedate+'",') # Process Date
                        outputfile.write('"'+tradeprice+'",') # Price
                        outputfile.write('"'+tradequantity+'",') # Qty
                        outputfile.write('"AaaaA",') # Trade Number
                        outputfile.write('"'+str(tradeprincipleamount)+'",') # Principal Amount
                        outputfile.write('"'+str(tradenetamount)+'",') # NET Amount
                        outputfile.write('"'+tradecommission+'",') # Commission Amount
                        outputfile.write('"'+tradetime+'",') # Execution Time
                        outputfile.write('"'+tradedate+' '+tradetime+'",') # Raw Exec. Time
                        outputfile.write('"N",') # Market Code
                        outputfile.write('"11111111",') # Trailer
                        outputfile.write('"'+tradesecfee+'",') # FeeSEC
                        outputfile.write('"'+trademiscfee+'",') # FeeMF
                        outputfile.write('"'+tradecatfee+'",') # Fee1
                        outputfile.write('"'+tradeoccfee+'",') # Fee2
                        outputfile.write('"'+tradensccfee+'",') # Fee3
                        outputfile.write('"'+tradeaccfee+'",') # FeeStamp
                        outputfile.write('"'+tradetaffee+'",') # FeeTAF
                        outputfile.write('"'+tradeclrfee+'",') # Fee4
                        outputfile.write('"",') # Sequence Number
                        outputfile.write('"",') # Side Seq Code
                        outputfile.write('"",') # Capacity Code
                        outputfile.write('"",') # Office Code
                        outputfile.write('"",') # Rep Code
                        outputfile.write('"",') # Special Code
                        outputfile.write('"",') # Instructions Trade Legend Code
                        outputfile.write('"",') # Factor Type2
                        outputfile.write('"",') # Trade Interest
                        outputfile.write('"",') # Original TradeNumber
                        outputfile.write('"'+tradetime+'",') # Entry Time
                        outputfile.write('"",') # Entered By
                        outputfile.write('"FIX",') # YieldToMature
                        outputfile.write('"",') # YieldToCall
                        outputfile.write('"",') # Mutual Fund Sales Charge Rate
                        outputfile.write('"",') # Mutual Fund Load Indicator
                        outputfile.write('"Trade"\n') # Transtype
                    else:
                        outputfile.write('"S",') # Side
                        outputfile.write('"'+tradesymbol+'",') #Symbol
                        outputfile.write('"A111111",')   # CUSIP
                        outputfile.write('"USD",')   #Currency Code
                        outputfile.write('"equity",') # Security Type
                        outputfile.write('"Long Sell",') # Buy/Sell
                        outputfile.write('"'+tradedate+'",') # Trade Date
                        outputfile.write('"'+tradedate+'",') # Settlement Date
                        outputfile.write('"'+tradedate+'",') # Process Date
                        outputfile.write('"'+tradeprice+'",') # Price
                        outputfile.write('"-'+tradequantity+'",') # Qty
                        outputfile.write('"AaaaA",') # Trade Number
                        outputfile.write('"-'+str(tradeprincipleamount)+'",') # Principal Amount
                        outputfile.write('"-'+str(tradenetamount)+'",') # NET Amount
                        outputfile.write('"'+tradecommission+'",') # Commission Amount
                        outputfile.write('"'+tradetime+'",') # Execution Time
                        outputfile.write('"'+tradedate+' '+tradetime+'",') # Raw Exec. Time
                        outputfile.write('"N",') # Market Code
                        outputfile.write('"11111111",') # Trailer
                        outputfile.write('"'+tradesecfee+'",') # FeeSEC
                        outputfile.write('"'+trademiscfee+'",') # FeeMF
                        outputfile.write('"'+tradecatfee+'",') # Fee1
                        outputfile.write('"'+tradeoccfee+'",') # Fee2
                        outputfile.write('"'+tradensccfee+'",') # Fee3
                        outputfile.write('"'+tradeaccfee+'",') # FeeStamp
                        outputfile.write('"'+tradetaffee+'",') # FeeTAF
                        outputfile.write('"'+tradeclrfee+'",') # Fee4
                        outputfile.write('"",') # Sequence Number
                        outputfile.write('"",') # Side Seq Code
                        outputfile.write('"",') # Capacity Code
                        outputfile.write('"",') # Office Code
                        outputfile.write('"",') # Rep Code
                        outputfile.write('"",') # Special Code
                        outputfile.write('"",') # Instructions Trade Legend Code
                        outputfile.write('"",') # Factor Type2
                        outputfile.write('"",') # Trade Interest
                        outputfile.write('"",') # Original TradeNumber
                        outputfile.write('"'+tradetime+'",') # Entry Time
                        outputfile.write('"",') # Entered By
                        outputfile.write('"FIX",') # YieldToMature
                        outputfile.write('"",') # YieldToCall
                        outputfile.write('"",') # Mutual Fund Sales Charge Rate
                        outputfile.write('"",') # Mutual Fund Load Indicator
                        outputfile.write('"Trade"\n') # Transtype
    outputfile.close()
    checkfile.close()
    os.rename(expfilenameprefix, expfilenameprefix+tradefirstdate.replace('/','')+'-'+tradedate.replace('/','')+'.csv' )


print(sys.argv)
if (len(sys.argv)==2):
    filename = sys.argv[1]
print("Processing: " + filename)
FixLittleEndianMarker(filename, fixedfilename)
ExceltoFileZellaCSV(fixedfilename)
