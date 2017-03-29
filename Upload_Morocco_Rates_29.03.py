# Author - Michal Zawadzki, michalmzawadzki@gmail.com. Updates/modifications highly encouraged (infoanarchism!). :)

import openpyxl, zipfile, urllib.request, pandas as pd, datetime, xml.etree.ElementTree as ET
from os import remove
from numpy import array
from time import sleep
from sys import exit
pd.options.mode.chained_assignment = None


# courtesy of Austin Taylor, http://www.austintaylor.io/ -- adapted for our use
def xml2df(root):
    all_records = []
    headers = []
    for i, child in enumerate(root):
        record = []
        for subchild in child:
            record.append(subchild.text)
            if subchild.tag not in headers:
                headers.append(subchild.tag)
        all_records.append(record)
    return pd.DataFrame(all_records, columns=headers)


# format the date as the bare int format is treated as General, and we need it to be an Excel Date type
# use openpyxl's builtin number formats for date_format
def format_date_to_excel(excel_file_location, date_format="mm-dd-yy"):
    wb = openpyxl.load_workbook(excel_file_location)
    ws = wb.active
    for row in ws:
        if "A" not in str((row[2]).value):  # skip header rows, picked "A" because column C headers have it :)
            row[2].number_format = date_format
    wb.save(excel_file_location)


def prepare_morocco():
    # clean old file, download the raw rates file
    try:
        remove("VATSPOTR.txt")
    except FileNotFoundError:
        pass
    try:
        print("Downloading MA rates...")
        urllib.request.urlretrieve("http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip", "VATSPOTR.zip")
    except:
        print(r"Oops! Cannot retrieve MA rates from http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip")

    myzip = zipfile.ZipFile("VATSPOTR.zip", "r")
    myzip.extractall()
    myzip.close()
    remove("VATSPOTR.zip")

prepare_morocco()

# create the header as a separate DF; could use one DataFrame once MultiIndex columns are better supported
header = pd.DataFrame([["CURRENCY_RATES", "COMPANY_ID=HP", "SOURCE=BOM-MAD", ""],
                       ["BASE_CURRENCY", "FOREIGN_CURRENCY", "EFFECTIVE_DATE", "RATE"]])

# read the txt to a DataFrame and leave only the currencies in scope
MA_csv = pd.read_csv("VATSPOTR.txt", sep="\t", header=1, index_col=False, parse_dates=[4])
MA_cur_in_scope = ["AED", "CAD", "CHF", "DZD", "EUR", "GBP", "LYD", "SAR", "SEK", "TND", "USD"]
rates_MA = MA_csv[(MA_csv.iloc[:,0] == "CBSEL") & (MA_csv.iloc[:,2] == "MAD") & (MA_csv.iloc[:,3].isin(MA_cur_in_scope))]

# note that rates in the raw file are normalized -- divide by the normalizer in order to get the actual rate
rates_MA.iloc[:,7] = rates_MA.iloc[:,7].div(rates_MA.iloc[:,8])

# get rid of useless columns
output_columns = [2, 3, 4, 7]
useless_columns = rates_MA[[x for x in range(rates_MA.shape[1]) if x not in output_columns]]
rates_MA.drop(useless_columns, axis=1, inplace=True)

# extract the rates' effective date for output file and the file's name -- must use Excel's number format
MA_effective_date = rates_MA.iloc[0,2]
MA_excel_date_format = (MA_effective_date - datetime.datetime(1899, 12, 31)).days + 1
rates_MA.iloc[:,2] = array(MA_excel_date_format)

# file path + name of the file
# use an elastic one in the future -- settings file?
MA_output_path = r"..\Upload_rates\Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_" + str(MA_effective_date)[:-9] + ".xlsx"

# create the final xlsx
try:
    with pd.ExcelWriter(MA_output_path, engine="openpyxl") as writer:
        header.to_excel(writer, index=False, header=False)
        rates_MA.to_excel(writer, index=False, header=False, startrow=2)
except:
    print("Unable to generate MA rates. :(")

format_date_to_excel(MA_output_path)

print("MA rates generated :)")

# parse the XML file and store it as a string
# genrealize and refactor later on :p
try:
    print("Downloading TR rates...")
    TR_rates_XML = urllib.request.urlopen("http://www.tcmb.gov.tr/kurlar/today.xml")
except:
    print(r"Oops! Cannot retrieve TR rates from http://www.tcmb.gov.tr/kurlar/today.xml")
    sleep(5)
    exit(1)
TR_rates_string = TR_rates_XML.read()

# create an ElementTree to easily access CurrencyCodes
TR_etree = ET.fromstring(TR_rates_string)

# retrieve a list of Currency codes (scope: the first 12)
TR_number_of_rates = 12
TR_cur_in_scope = pd.Series([child.attrib["CurrencyCode"] for child in TR_etree[:TR_number_of_rates]])

# convert the ElementTree to a DataFrame for easy manipulation and Excel conversion
TR_xml_df = xml2df(TR_etree)

# retrieve and format the dates -- TODO: check whether Tarih == Date
TR_effective_dates = pd.to_datetime(pd.Series([TR_etree.attrib["Date"] for i in range(TR_number_of_rates)]))
TR_effective_date = TR_effective_dates[0]
TR_excel_date_format = (TR_effective_date - datetime.datetime(1899, 12, 31)).days + 1

TR_base_cur = pd.Series(["TRY" for _ in range(TR_number_of_rates)])
TR_rates = TR_xml_df.iloc[:TR_number_of_rates,4].astype(float)

# use the real values of the rates
normalizers = TR_xml_df.iloc[:TR_number_of_rates, 0].astype(int)
TR_rates_denormalized = TR_rates.div(normalizers)

TR_data = pd.concat([TR_base_cur, TR_cur_in_scope, TR_effective_dates, TR_rates_denormalized], axis=1)

# convert dates to Excel's numeric date format
TR_data.iloc[:,2] = array(TR_excel_date_format)

# name of the TR excel file
TR_output_path = r"..\Upload_rates\Other Rates\TURKEY_RATES\TURKEY_RATES_" + str(TR_effective_date)[:-9] + ".xlsx"

# update the header
header.iloc[0, 2] = "SOURCE=TNB-TRY"

# create the final xlsx
try:
    with pd.ExcelWriter(TR_output_path, engine="openpyxl") as writer:
        header.to_excel(writer, index=False, header=False)
        TR_data.to_excel(writer, index=False, header=False, startrow=2)
except:
    print("Unable to generate TR rates. :(")

format_date_to_excel(TR_output_path)

print("TR rates generated :)")

# TODO: create a settings file with the destination folder for the output file
# if the directory does not exist - create it
# beautify the final date converting if statement - maybe isinstance(row[2].value, basestring)?
# refactor

# cleanup
remove("VATSPOTR.txt")
