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


# download xml and parse to an ElementTree Element object
def xml_to_element_tree(rates_url, country_abbreviation):
    try:
        print("Downloading {} rates...\n".format(country_abbreviation))
        rates_xml = urlopen(rates_url)
        rates_string = rates_xml.read()
        rates_element_tree = ElementTree.fromstring(rates_string)
        return rates_element_tree
    except:
        print(r"Oops! Cannot retrieve {} rates from {}\n".format(country_abbreviation, rates_url))
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
        print("Downloading MA rates...\n")
        urlretrieve("http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip", "VATSPOTR.zip")
    except:
        print(r"Oops! Cannot retrieve MA rates from http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip\n")

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


def generate_excel_output(header, data, output_path, country_abbreviation, output_date_format="mm-dd-yy"):
    try:
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            header.to_excel(writer, index=False, header=False)
            data.to_excel(writer, index=False, header=False, startrow=2)
            print("{} rates generated :)\n".format(country_abbreviation))
    except:
        print("Unable to generate {} rates. :(\n".format(country_abbreviation))

    format_date_to_excel(output_path, date_format=output_date_format)


##########################################
################ MOROCCO #################
##########################################


prepare_morocco()

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
#MA_data.iloc[:, 2] = array(MA_excel_date_format) unnecessary? -> delete MA_excel_date_format too

# create the final xlsx
MA_output_path = r"..\Upload_rates\Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_" + str(MA_effective_date)[:-9] + ".xlsx"
MA_header = generate_header("MA")
generate_excel_output(MA_header, MA_data, MA_output_path, "MA")

# cleanup
remove("VATSPOTR.txt")


##########################################
################# TURKEY #################
##########################################


TR_etree = xml_to_element_tree("http://www.tcmb.gov.tr/kurlar/today.xml", "TR")

TR_number_of_rates = 12 # first twelve currencies
TR_cur_in_scope = pd.Series([child.attrib["CurrencyCode"] for child in TR_etree[:TR_number_of_rates]])

# convert the ElementTree to a DataFrame for easy manipulation and Excel conversion
TR_xml_df = xml2df(TR_etree) # TODO just use ET

# retrieve and format the dates
TR_effective_dates = pd.to_datetime(pd.Series([TR_etree.attrib["Date"] for _rate in range(TR_number_of_rates)]))
TR_effective_date = TR_effective_dates[0]
TR_excel_date_format = (TR_effective_date - datetime(1899, 12, 31)).days + 1

TR_base_cur = pd.Series(["TRY" for _rate in range(TR_number_of_rates)])
TR_rates = TR_xml_df.iloc[:TR_number_of_rates, 4].astype(float)

# use the real values of the rates
TR_normalizers = TR_xml_df.iloc[:TR_number_of_rates, 0].astype(int)
TR_rates_denormalized = TR_rates.div(TR_normalizers)

TR_data = pd.concat([TR_base_cur, TR_cur_in_scope, TR_effective_dates, TR_rates_denormalized], axis=1)

# convert dates to Excel's numeric date format
#TR_data.iloc[:, 2] = array(TR_excel_date_format) unnecessary? + delete excel_format

# create the final xlsx TODO: define a function for creating this path; use a settings file
TR_output_path = r"..\Upload_rates\Other Rates\TURKEY_RATES\TURKEY_RATES_" + str(TR_effective_date)[:-9] + ".xlsx"
TR_header = generate_header("TR")
generate_excel_output(TR_header, TR_data, TR_output_path, "TR")


##########################################
################ SLOVAKIA ################
##########################################


SK_etree = xml_to_element_tree("http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml", "SK")

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

# paste excel date int format
#SK_data.iloc[:, 2] = array(SK_excel_date_format) unnecessary?

# generate the final xlsx
SK_output_path = r"..\Upload_rates\Other Rates\SLOVAKIA_RATES\SLOVAKIA_RATES_" + str(SK_effective_date)[:-9] + ".xlsx"
SK_header = generate_header("SK")
generate_excel_output(SK_header, SK_data, SK_output_path, "SK")


##########################################
################# RUSSIA #################
##########################################


RU_etree = xml_to_element_tree("http://www.cbr.ru/scripts/XML_daily_eng.asp?", "RU")

RU_number_of_rates = len(RU_etree.getchildren())
RU_base_cur = pd.Series(["RUB" for _rate in range(RU_number_of_rates)])
RU_cur_in_scope = pd.Series([RU_etree[i][1].text for i in range(RU_number_of_rates)])

# retrieve the rates in string format and convert to float
RU_rates_txt = pd.Series([RU_etree[i][-1].text for i in range(RU_number_of_rates)])
RU_rates = RU_rates_txt.str.replace(",", ".").apply(lambda x: float(x))

# Dates magic. Replace symbols for easy conversion. Pandas infers the date format incorrectly; adjust it manually.
RU_effective_dates_str = pd.Series([RU_etree.attrib["Date"].replace(".", "/") for _rate in range(RU_number_of_rates)])
RU_effective_dates = pd.to_datetime(RU_effective_dates_str)
RU_effective_dates = pd.to_datetime(RU_effective_dates.dt.strftime("%d/%m/%Y"))
RU_effective_date = RU_effective_dates[0]

RU_excel_date_format = (RU_effective_date - datetime(1899, 12, 31)).days + 1

# normalize the rates
RU_normalizers = [int(RU_etree[i][2].text) for i in range(RU_number_of_rates)]
RU_rates  = RU_rates.div(RU_normalizers)

RU_data = pd.concat([RU_base_cur, RU_cur_in_scope, RU_effective_dates, RU_rates], axis=1)
#RU_data.iloc[:, 2] = array(RU_excel_date_format) unnecessary?

# delete out_of_scope rates, change TMT to TMM
RU_out_of_scope_rates = ["XDR", "XAU"]
RU_data = RU_data[~(RU_data.iloc[:, 1].isin(RU_out_of_scope_rates))]
RU_data.replace("TMT", "TMM", inplace=True)

# generate the final xlsx
RU_output_path = r"..\Upload_rates\Other Rates\RUSSIA_RATES\RUSSIA_RATES_" + str(RU_effective_date)[:-9] + ".xlsx"
RU_header = generate_header("RU")
generate_excel_output(RU_header, RU_data, RU_output_path, "RU")


##########################################
################ POLAND A ################
##########################################


PL_A_etree = xml_to_element_tree("http://www.nbp.pl/kursy/xml/LastA.xml", "PL A")

PL_A_number_of_rates = len(PL_A_etree.getchildren()) - 2
PL_A_base_cur = pd.Series(["PLN" for _rate in range(PL_A_number_of_rates)])
PL_A_cur_in_scope = pd.Series([PL_A_etree[i][2].text for i in range(2, PL_A_number_of_rates + 2)])

# retrieve the rates in string format and convert to float
PL_A_rates_txt = pd.Series([PL_A_etree[i][-1].text for i in range(2, PL_A_number_of_rates + 2)])
PL_A_rates = PL_A_rates_txt.str.replace(",", ".").apply(lambda x: float(x))

# denormalize
PL_A_normalizers = [int(PL_A_etree[i][1].text) for i in range(2, PL_A_number_of_rates + 2)]
PL_A_rates  = PL_A_rates.div(PL_A_normalizers)

# dates magic
PL_A_effective_dates_str = pd.Series([PL_A_etree[1].text for _rate in range(PL_A_number_of_rates)])
PL_A_effective_dates = pd.to_datetime(PL_A_effective_dates_str)
PL_A_effective_date = PL_A_effective_dates[0]
PL_A_effective_dates = pd.to_datetime(PL_A_effective_dates.dt.strftime("%m/%d/%Y"))

PL_A_data = pd.concat([PL_A_base_cur, PL_A_cur_in_scope, PL_A_effective_dates, PL_A_rates], axis=1)
#PL_A_data.iloc[:, 2] = array(PL_A_excel_date_format) unnecessary?

PL_A_out_of_scope_rate = "XDR"
PL_A_replacement_rates = {"AFN": "AFA", "GHS": "GHC", "MGA": "MGF", "MZN": "MZM", "SDG": "SDD", "SRD": "SRG", "ZWL": "ZWD"}
PL_A_data = PL_A_data[PL_A_data.iloc[:, 1] != PL_A_out_of_scope_rate]
PL_A_data.replace(PL_A_replacement_rates, inplace=True)

# generate the final xlsx
PL_A_output_path = r"..\Upload_rates\Other Rates\POLAND_A_RATES\POLAND_A_RATES_" + str(PL_A_effective_date)[:-9] + ".xlsx"
PL_A_header = generate_header("PL")
generate_excel_output(PL_A_header, PL_A_data, PL_A_output_path, "PL A")


##########################################
################ POLAND B ################
##########################################


PL_B_etree = xml_to_element_tree("http://www.nbp.pl/kursy/xml/LastB.xml", "PL B")

PL_B_number_of_rates = len(PL_B_etree.getchildren()) - 2
PL_B_base_cur = pd.Series(["PLN" for _rate in range(PL_B_number_of_rates)])
PL_B_cur_in_scope = pd.Series([PL_B_etree[i][2].text for i in range(2, PL_B_number_of_rates + 2)])

# retrieve the rates in string format and convert to float
PL_B_rates_txt = pd.Series([PL_B_etree[i][-1].text for i in range(2, PL_B_number_of_rates + 2)])
PL_B_rates = PL_B_rates_txt.str.replace(",", ".").apply(lambda x: float(x))

# denormalize
PL_B_normalizers = [int(PL_B_etree[i][1].text) for i in range(2, PL_B_number_of_rates + 2)]
PL_B_rates  = PL_B_rates.div(PL_B_normalizers)

# dates magic
PL_B_effective_dates_str = pd.Series([PL_B_etree[1].text for _rate in range(PL_B_number_of_rates)])
PL_B_effective_dates = pd.to_datetime(PL_B_effective_dates_str)
PL_B_effective_date = PL_B_effective_dates[0]
PL_B_effective_dates = pd.to_datetime(PL_B_effective_dates.dt.strftime("%m/%d/%Y"))

PL_B_data = pd.concat([PL_B_base_cur, PL_B_cur_in_scope, PL_B_effective_dates, PL_B_rates], axis=1)
#PL_B_data.iloc[:, 2] = array(PL_B_excel_date_format) unnecessary?

PL_B_replacement_rates = {"AFN": "AFA", "GHS": "GHC", "MGA": "MGF", "MZN": "MZM", "SDG": "SDD", "SRD": "SRG",
                          "ZWL": "ZWD", "ZMW": "ZMK"}
PL_B_data.replace(PL_B_replacement_rates, inplace=True)

# generate the final xlsx
PL_B_output_path = r"..\Upload_rates\Other Rates\POLAND_B_RATES\POLAND_B_RATES_" + str(PL_B_effective_date)[:-9] + ".xlsx"
PL_B_header = generate_header("PL")
generate_excel_output(PL_B_header, PL_B_data, PL_B_output_path, "PL B")

# TODO
# TODO create a settings file with the destination folder for the output file
# TODO

# if the directory does not exist - create it
# beautify the final date converting if statement - maybe isinstance(row[2].value, basestring)?
# refactor
# add a log file?
