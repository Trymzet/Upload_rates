# Author - Michal Zawadzki, michalmzawadzki@gmail.com. Updates/modifications highly encouraged (infoanarchism!). :)

import openpyxl, os, zipfile, urllib.request, pandas as pd, sys, numpy as np, datetime
pd.options.mode.chained_assignment = None

# clean old file, download the raw rates file
try:
    os.remove("VATSPOTR.txt")
except FileNotFoundError:
    pass
try:
    print("Downloading rates...")
    urllib.request.urlretrieve("http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip", "VATSPOTR.zip")
except:
    print(r"Oops! Cannot retrieve MA rates from http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip")
    sys.exit(1)

myzip = zipfile.ZipFile("VATSPOTR.zip", "r")
myzip.extractall()
myzip.close()
os.remove("VATSPOTR.zip")

# create a DataFrame
rates = pd.read_csv("VATSPOTR.txt", sep="\t", header=1, index_col=False, parse_dates=[4])
cur_list = ["AED", "CAD", "CHF", "DZD", "EUR", "GBP", "LYD", "SAR", "SEK", "TND", "USD"] # our scope
rates_MA = rates[(rates.iloc[:,0] == "CBSEL") & (rates.iloc[:,2] == "MAD") & (rates.iloc[:,3].isin(cur_list))]

# note that rates are normalized -- divide by the normalizer in order to get the actual rate
rates_MA.iloc[:,7] = rates_MA.iloc[:,7].div(rates_MA.iloc[:,8], axis=0)

# get rid of useless columns and reset index
useless_cols = [x for x in range(rates_MA.shape[1]) if x not in [2, 3, 4, 7]]
rates_MA.drop(rates_MA[useless_cols], axis=1, inplace=True) #rates_MA.drop(rates_MA.columns[useless_cols], axis=1, inplace=True)

# extract the rates' effective date for output file and the file's name -- must use Excel's number format
date = rates_MA.iloc[0,2]
excel_date_format = (date - datetime.datetime(1899, 12, 31)).days + 1
#rates_MA.iloc[:,2] = rates_MA.iloc[:,2].dt.strftime("%d-%m-%Y") #"%m/%d/%Y"
rates_MA.iloc[:,2] = np.array(excel_date_format)
print(rates_MA)

# file path + name of the file
title = r"..\Upload_rates\Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_" + str(date)[:-9] + ".xlsx"

# create the header
header = pd.DataFrame([["CURRENCY_RATES", "COMPANY_ID=HP", "SOURCE=BOM-MAD", ""],
                       ["BASE_CURRENCY", "FOREIGN_CURRENCY", "EFFECTIVE_DATE", "RATE"]])

# create the final xlsx
with pd.ExcelWriter(title, engine="openpyxl") as writer:
    header.to_excel(writer, index=False, header=False)
    rates_MA.to_excel(writer, index=False, header=False, startrow=2)

# cleanup
os.remove("VATSPOTR.txt")
