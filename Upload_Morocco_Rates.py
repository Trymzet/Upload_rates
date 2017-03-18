# Author - Michal Zawadzki, michalmzawadzki@gmail.com. Updates/modifications highly encouraged (infoanarchism!). :)

import openpyxl, os, zipfile, urllib.request, pandas as pd, sys, numpy as np
pd.options.mode.chained_assignment = None

# cleaning files, downloading the rates
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
rates = pd.read_csv("VATSPOTR.txt", sep="\t", header=1, index_col=False)
cur_list = ["AED", "CAD", "CHF", "DZD", "EUR", "GBP", "LYD", "SAR", "SEK", "TND", "USD"] # our scope
rates_MA = rates[(rates.iloc[:,0] == "CBSEL") & (rates.iloc[:,2] == "MAD") & (rates.iloc[:,3].isin(cur_list))]

# get rid of useless columns and reset index
useless_cols = [x for x in range(rates_MA.shape[1]) if x not in [2, 3, 4, 7]]
rates_MA.drop(rates_MA[useless_cols], axis=1, inplace=True) #rates_MA.drop(rates_MA.columns[useless_cols], axis=1, inplace=True)
rates_MA.reset_index(drop=True, inplace=True)
rates_MA = pd.DataFrame(rates_MA, index=[x for x in range(rates_MA.shape[0])])

#add the header
rates_MA.columns = [["CURRENCY_RATES", "COMPANY_ID=HP", "SOURCE=BOM-MAD", " "], ["BASE_CURRENCY", "FOREIGN_CURRENCY", "EFFECTIVE_DATE", "RATE"]]

print(rates_MA)

# extract the rates' effective date for output file and the file's name
date = rates_MA["SOURCE=BOM-MAD"]["EFFECTIVE_DATE"][0]
date_formatted = str(pd.to_datetime(date, format="%Y%m%d"))[:-9] # slicing the hours out
print(date_formatted)
""""
# infer_datetime_format=True
#dates = rates_MA.iloc[2:,2]
#dates = pd.to_datetime(dates.astype(str), format="%Y%m%d")
#dates = dates.dt.strftime("%m%d%Y")
#rates_MA.iloc[2:,2] = dates
#print(dates)
"""

# file path + name of the file
title = r"..\Upload_rates\Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_" + date_formatted + ".xls"

# convert the DataFrame to the final xls
rates_MA.to_excel(title, "ExchangeRates", index=[x for x in range[1, 14]])

# cleanup
os.remove("VATSPOTR.txt")
