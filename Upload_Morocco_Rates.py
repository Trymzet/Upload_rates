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
        print("Downloading rates...")
        urllib.request.urlretrieve("http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip", "VATSPOTR.zip")
    except:
        print(r"Oops! Cannot retrieve MA rates from http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip")
        sleep(5)
        exit(1)

    myzip = zipfile.ZipFile("VATSPOTR.zip", "r")
    myzip.extractall()
    myzip.close()
    remove("VATSPOTR.zip")

# parse the XML file and store it as a string
# genrealize and refactor later on :p
TR_rates_XML = urllib.request.urlopen("http://www.tcmb.gov.tr/kurlar/today.xml")
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

print(TR_excel_date_format)

TR_base_cur = pd.Series(["TRY" for _ in range(TR_number_of_rates)])
TR_rates = TR_xml_df.iloc[:TR_number_of_rates,4].astype(float)

# use the real values of the rates
normalizers = TR_xml_df.iloc[:TR_number_of_rates, 0].astype(int)
TR_rates_denormalized = TR_rates.div(normalizers)

TR_data = pd.concat([TR_base_cur, TR_cur_in_scope, TR_effective_dates, TR_rates_denormalized], axis=1)

print(TR_data)

# name of the TR excel file
TR_title = r"..\Upload_rates\Other Rates\TURKEY_RATES_" + str(TR_effective_date)[:-9] + ".xlsx"
print(TR_title)

# create the final xlsx
with pd.ExcelWriter(TR_title, engine="openpyxl") as writer:
    header.to_excel(writer, index=False, header=False)
    TR_data.to_excel(writer, index=False, header=False, startrow=2)

#TODO: find date format for TR dates
format_date_to_excel(TR_title, date_format="")
