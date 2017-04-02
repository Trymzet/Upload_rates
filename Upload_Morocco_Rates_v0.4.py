# Author - Michal Zawadzki, michalmzawadzki@gmail.com. Updates/modifications highly encouraged (infoanarchism!). :)

import openpyxl
import pandas as pd
import xml.etree.ElementTree as ElementTree
from zipfile import ZipFile
from urllib.request import urlopen, urlretrieve
from datetime import datetime
from os import remove
from numpy import array
pd.options.mode.chained_assignment = None

# for testing
try:
    remove("Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_2017-04-01.xlsx")
    remove("Other rates\TURKEY_RATES\TURKEY_RATES_2017-03-31.xlsx")
    remove("Other rates\SLOVAKIA_RATES\SLOVAKIA_RATES_2017-03-31.xlsx")
except:
    pass


# download xml, convert to string format
def load_xml(rates_url, country_abbreviation):
    try:
        print("Downloading {} rates...".format(country_abbreviation))
        rates_xml = urlopen(rates_url)
        rates_string = rates_xml.read()
        return rates_string
    except:
        print(r"Oops! Cannot retrieve {} rates from {}".format(country_abbreviation, rates_url))
        return


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
        urlretrieve("http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip", "VATSPOTR.zip")
    except:
        print(r"Oops! Cannot retrieve MA rates from http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip")

    myzip = ZipFile("VATSPOTR.zip", "r")
    myzip.extractall()
    myzip.close()
    remove("VATSPOTR.zip")


def generate_header(country_abbreviation):
    header = pd.DataFrame([["CURRENCY_RATES", "COMPANY_ID=HP", "", ""],
                       ["BASE_CURRENCY", "FOREIGN_CURRENCY", "EFFECTIVE_DATE", "RATE"]])
    sources = {"MA": "SOURCE=BOM-MAD", "TR": "SOURCE=TNB-TRY", "SK": "SOURCE=ECB-EUR",
               "RU": "SOURCE=NBR-RUB", "PL": "SOURCE=PNB-PLN"}
    header.iloc[0][2] = sources[country_abbreviation]
    return header


def generate_excel_output(header, data, output_path, country_abbreviation):
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            header.to_excel(writer, index=False, header=False)
            data.to_excel(writer, index=False, header=False, startrow=2)
            print("{} rates generated :)".format(country_abbreviation))
    except:
        print("Unable to generate {} rates. :(".format(country_abbreviation))

    format_date_to_excel(output_path)


prepare_morocco()

# create the header as a separate DF; could use one DataFrame once MultiIndex columns are better supported
MA_header = generate_header("MA")

# read the txt to a DataFrame and leave only the currencies in scope
MA_csv = pd.read_csv("VATSPOTR.txt", sep="\t", header=1, index_col=False, parse_dates=[4])
MA_cur_in_scope = ["AED", "CAD", "CHF", "DZD", "EUR", "GBP", "LYD", "SAR", "SEK", "TND", "USD"]
MA_data = MA_csv[(MA_csv.iloc[:, 0] == "CBSEL") & (MA_csv.iloc[:, 2] == "MAD") & (MA_csv.iloc[:, 3].isin(MA_cur_in_scope))]

# note that rates in the raw file are normalized -- divide by the normalizer in order to get the actual rate
MA_normalizer = MA_data.iloc[:, 8]
MA_data.iloc[:, 7] = MA_data.iloc[:, 7].div(MA_normalizer)

# get rid of useless columns
output_columns = [2, 3, 4, 7]
useless_columns = MA_data[[x for x in range(MA_data.shape[1]) if x not in output_columns]]
MA_data.drop(useless_columns, axis=1, inplace=True)

# extract the rates' effective date for output file and the file's name -- must use Excel's number format
MA_effective_date = MA_data.iloc[0, 2]
MA_excel_date_format = (MA_effective_date - datetime(1899, 12, 31)).days + 1
MA_data.iloc[:, 2] = array(MA_excel_date_format)

# file path + name of the file
# use an elastic one in the future -- settings file?
MA_output_path = r"..\Upload_rates\Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_" + str(MA_effective_date)[:-9] + ".xlsx"

# create the final xlsx
generate_excel_output(MA_header, MA_data, MA_output_path, "MA")

# cleanup
remove("VATSPOTR.txt")

TR_rates_string = load_xml("http://www.tcmb.gov.tr/kurlar/today.xml", "TR")

# create an ElementTree to easily access CurrencyCodes
TR_etree = ElementTree.fromstring(TR_rates_string)

# retrieve a list of Currency codes (scope: the first 12)
TR_number_of_rates = 12
TR_cur_in_scope = pd.Series([child.attrib["CurrencyCode"] for child in TR_etree[:TR_number_of_rates]])

# convert the ElementTree to a DataFrame for easy manipulation and Excel conversion
TR_xml_df = xml2df(TR_etree)

# retrieve and format the dates -- TODO: check whether Tarih == Date
TR_effective_dates = pd.to_datetime(pd.Series([TR_etree.attrib["Date"] for _rate in range(TR_number_of_rates)]))
TR_effective_date = TR_effective_dates[0]
TR_excel_date_format = (TR_effective_date - datetime(1899, 12, 31)).days + 1

TR_base_cur = pd.Series(["TRY" for _rate in range(TR_number_of_rates)])
TR_rates = TR_xml_df.iloc[:TR_number_of_rates, 4].astype(float)

# use the real values of the rates
normalizers = TR_xml_df.iloc[:TR_number_of_rates, 0].astype(int)
TR_rates_denormalized = TR_rates.div(normalizers)

TR_data = pd.concat([TR_base_cur, TR_cur_in_scope, TR_effective_dates, TR_rates_denormalized], axis=1)

# convert dates to Excel's numeric date format
TR_data.iloc[:, 2] = array(TR_excel_date_format)

# name of the TR excel file TODO: define a function for creating this path; use a settings file
TR_output_path = r"..\Upload_rates\Other Rates\TURKEY_RATES\TURKEY_RATES_" + str(TR_effective_date)[:-9] + ".xlsx"

# update the header
TR_header = generate_header("TR")

# create the final xlsx
generate_excel_output(TR_header, TR_data, TR_output_path, "TR")

SK_rates_string = load_xml("http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml", "SK")
SK_etree = ElementTree.fromstring(SK_rates_string)

SK_number_of_rates = len(SK_etree[2][0].getchildren())
SK_base_cur = pd.Series(["EUR" for _rate in range(SK_number_of_rates)])
SK_cur_in_scope = pd.Series([SK_etree[2][0][i].attrib["currency"] for i in range(SK_number_of_rates)])
SK_rates = pd.Series([SK_etree[2][0][i].attrib["rate"] for i in range(SK_number_of_rates)]).astype(float)

# dates magic
SK_effective_dates = pd.to_datetime(pd.Series([SK_etree[2][0].attrib["time"] for _rate in range(SK_number_of_rates)]))
SK_effective_date = SK_effective_dates[0]
SK_excel_date_format = (SK_effective_date - datetime(1899, 12, 31)).days + 1

SK_data = pd.concat([SK_base_cur, SK_cur_in_scope, SK_effective_dates, SK_rates], axis=1)

# reverse rate values, so that it's e.g. USD/EUR and not EUR/USD
SK_data.iloc[:, -1] = 1 / SK_data.iloc[:, -1]

# paste excelt date int format
SK_data.iloc[:, 2] = array(SK_excel_date_format)

# name of the TR excel file
SK_output_path = r"..\Upload_rates\Other Rates\SLOVAKIA_RATES\SLOVAKIA_RATES_" + str(TR_effective_date)[:-9] + ".xlsx"

SK_header = generate_header("SK")

# create the final xlsx
generate_excel_output(SK_header, SK_data, SK_output_path, "SK")



# TODO
# TODO  -> fix rates for SK - normalizer?
# TODO

# TODO: create a settings file with the destination folder for the output file
# if the directory does not exist - create it
# beautify the final date converting if statement - maybe isinstance(row[2].value, basestring)?
# refactor
