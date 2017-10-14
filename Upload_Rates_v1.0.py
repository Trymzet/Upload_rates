# Author - Michal Zawadzki, michalmzawadzki@gmail.com. Updates/modifications highly encouraged (infoanarchism!). :)

import xml.etree.ElementTree as ElementTree
from openpyxl import load_workbook
from zipfile import ZipFile
from urllib.request import urlopen
from urllib.request import urlretrieve
from datetime import datetime
from os import remove
from os.path import exists
from os import makedirs
from time import sleep
from pandas import DataFrame
from pandas import Series
from pandas import concat
from pandas import read_csv
from pandas import ExcelWriter
from pandas import to_datetime
from pandas import options
options.mode.chained_assignment = None


def generate_output_folder():
    today_date = datetime.now().strftime("%Y-%m-%d")
    destination_folder = "Rates\\" + today_date + "_rates"  # e.g. Rates\2017-04-04_rates
    if not exists(destination_folder):
        makedirs(destination_folder)
    return destination_folder


def prepare_morocco():
    # delete the old file, download the raw rates file
    try:
        remove("VATSPOTR.txt")
    except FileNotFoundError:
        pass
    try:
        print("Downloading MA rates...")
        urlretrieve("http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip", "VATSPOTR.zip")
    except:
        print("Oops! Cannot retrieve MA rates from http://polaris-pro-ent.houston.hpe.com:8080/VATSPOTR.zip \n"
              "Please ensure your corporate connectivity is working.\n")
        sleep(2)
        return 1
    myzip = ZipFile("VATSPOTR.zip", "r")
    myzip.extractall()
    myzip.close()
    remove("VATSPOTR.zip")


# download xml and parse to an ElementTree Element object
def xml_to_element_tree(rates_url, country_abbreviation):
    try:
        print("Downloading {} rates...".format(country_abbreviation))
        rates_xml = urlopen(rates_url)
        rates_string = rates_xml.read()
        rates_element_tree = ElementTree.fromstring(rates_string)
        return rates_element_tree
    except:
        print("Oops! Cannot retrieve {} rates from {}\n".format(country_abbreviation, rates_url))
        print("Self-detonating in: {}".format("3"))
        sleep(1)
        print("Self-detonating in: {}".format("2"))
        sleep(1)
        print("Self-detonating in: {}".format("1"))
        sleep(1)
        print("Self-detonating in: {}".format("0"))
        sleep(1)
        print("Robot joke! But seriously, please fix me ¥[*.*]¥\n")
        sleep(2)
        return 1


# format the date as the bare int format is treated as General, and we need it to be an Excel Date type
# use openpyxl's builtin number formats for date_format
def format_date_to_excel(excel_file_location, date_format="mm-dd-yy"):
    wb = load_workbook(excel_file_location)
    ws = wb.active
    for row in ws:
        if isinstance(row[2].value, datetime): # the dates are pandas datetime Series, so they are easily recognized
            row[2].number_format = date_format
    wb.save(excel_file_location)


def generate_header(country_abbreviation):
    header = DataFrame([["CURRENCY_RATES", "COMPANY_ID=HP", "", ""],
                       ["BASE_CURRENCY", "FOREIGN_CURRENCY", "EFFECTIVE_DATE", "RATE"]])
    sources = {"MA": "SOURCE=BOM-MAD", "TR": "SOURCE=TNB-TRY", "SK": "SOURCE=ECB-EUR",
               "RU": "SOURCE=NBR-RUB", "PL": "SOURCE=PNB-PLN"}
    header.iloc[0][2] = sources[country_abbreviation]
    return header


def generate_excel_output(header, data, output_path, country_abbreviation, output_date_format="mm-dd-yy"):
    try:
        with ExcelWriter(output_path, engine="openpyxl") as writer:
            header.to_excel(writer, index=False, header=False)
            data.to_excel(writer, index=False, header=False, startrow=2)
            print("{} rates generated :)\n".format(country_abbreviation))
    except:
        print("Unable to generate {} rates. :(\n".format(country_abbreviation))

    format_date_to_excel(output_path, date_format=output_date_format)


##########################################
################ MOROCCO #################
##########################################


def generate_MA():
    cleanup = prepare_morocco()
    if cleanup == 1:
        return

    # read the txt to a DataFrame and leave only the currencies in scope
    MA_csv = read_csv("VATSPOTR.txt", sep="\t", header=1, index_col=False, parse_dates=[4])
    MA_cur_in_scope = ["AED", "CAD", "CHF", "DZD", "EUR", "GBP", "LYD", "SAR", "SEK", "TND", "USD"]
    MA_data = MA_csv[(MA_csv.iloc[:, 0] == "CBSEL") & (MA_csv.iloc[:, 2] == "MAD") & (MA_csv.iloc[:, 3].isin(MA_cur_in_scope))]

    # note that rates in the raw file are normalized -- divide by the normalizer in order to get the actual rate
    MA_normalizer = MA_data.iloc[:, 8]
    MA_data.iloc[:, 7] = MA_data.iloc[:, 7].div(MA_normalizer)

    # get rid of useless columns
    output_columns = [2, 3, 4, 7]
    useless_columns = MA_data[[x for x in range(MA_data.shape[1]) if x not in output_columns]]
    MA_data.drop(useless_columns, axis=1, inplace=True)

    # extract the rates' effective date for output file and the file's name
    MA_effective_date = MA_data.iloc[0, 2]

    MA_output_path = output_folder + r"\MOROCCO_RATES_" + str(MA_effective_date)[:-9] + ".xlsx"
    MA_header = generate_header("MA")
    generate_excel_output(MA_header, MA_data, MA_output_path, "MA")

    try:
        remove("VATSPOTR.txt")
    except FileNotFoundError:
        pass


##########################################
################# TURKEY #################
##########################################


def generate_TR():

    TR_etree = xml_to_element_tree("http://www.tcmb.gov.tr/kurlar/today.xml", "TR")
    if TR_etree == 1:
        return

    TR_number_of_rates = 12 # first twelve currencies
    TR_cur_in_scope = Series([child.attrib["CurrencyCode"] for child in TR_etree[:TR_number_of_rates]])

    TR_effective_dates = to_datetime(Series([TR_etree.attrib["Date"] for _rate in range(TR_number_of_rates)]))
    TR_effective_date = TR_effective_dates[0]

    TR_base_cur = Series(["TRY" for _rate in range(TR_number_of_rates)])
    TR_rates = Series([float(TR_etree[i][3].text) for i in range(TR_number_of_rates)])

    # use the real values of the rates
    TR_normalizers = [int(TR_etree[i][0].text) for i in range(TR_number_of_rates)]
    TR_rates_denormalized = TR_rates.div(TR_normalizers)

    TR_data = concat([TR_base_cur, TR_cur_in_scope, TR_effective_dates, TR_rates_denormalized], axis=1)

    TR_output_path = output_folder + r"\TURKEY_RATES_" + str(TR_effective_date)[:-9] + ".xlsx"
    TR_header = generate_header("TR")
    generate_excel_output(TR_header, TR_data, TR_output_path, "TR")


##########################################
################ SLOVAKIA ################
##########################################


def generate_SK():
    SK_etree = xml_to_element_tree("http://www.ecb.europa.eu/stats/eurofxref/eurofxref-daily.xml", "SK")
    if SK_etree == 1:
        return

    SK_number_of_rates = len(SK_etree[2][0].getchildren())
    SK_base_cur = Series(["EUR" for _rate in range(SK_number_of_rates)])
    SK_cur_in_scope = Series([SK_etree[2][0][i].attrib["currency"] for i in range(SK_number_of_rates)])
    SK_rates = Series([SK_etree[2][0][i].attrib["rate"] for i in range(SK_number_of_rates)]).astype(float)

    SK_effective_dates = to_datetime(Series([SK_etree[2][0].attrib["time"] for _rate in range(SK_number_of_rates)]))
    SK_effective_date = SK_effective_dates[0]

    SK_data = concat([SK_base_cur, SK_cur_in_scope, SK_effective_dates, SK_rates], axis=1)

    # reverse rate values, so that it's e.g. USD/EUR and not EUR/USD
    SK_data.iloc[:, -1] = 1 / SK_data.iloc[:, -1]

    SK_output_path = output_folder + r"\SLOVAKIA_RATES_" + str(SK_effective_date)[:-9] + ".xlsx"
    SK_header = generate_header("SK")
    generate_excel_output(SK_header, SK_data, SK_output_path, "SK")


##########################################
################# RUSSIA #################
##########################################


def generate_RU():

    today_date = datetime.now().strftime("%d/%m/%Y")
    try:
        RU_etree = xml_to_element_tree("http://www.cbr.ru/scripts/XML_daily_eng.asp?date_req=" + str(today_date), "RU")
    except:
        RU_etree = xml_to_element_tree("http://www.cbr.ru/scripts/XML_daily_eng.asp?", "RU")
    if RU_etree == 1:
        return

    RU_number_of_rates = len(RU_etree.getchildren())
    RU_base_cur = Series(["RUB" for _rate in range(RU_number_of_rates)])
    RU_cur_in_scope = Series([RU_etree[i][1].text for i in range(RU_number_of_rates)])

    # retrieve the rates in string format and convert to float
    RU_rates_txt = Series([RU_etree[i][-1].text for i in range(RU_number_of_rates)])
    RU_rates = RU_rates_txt.str.replace(",", ".").apply(lambda x: float(x))

    # Replace symbols for easy conversion. Pandas infers the date format incorrectly; adjust it manually.
    RU_effective_dates_str = Series([RU_etree.attrib["Date"].replace(".", "/") for _rate in range(RU_number_of_rates)])
    RU_effective_dates = to_datetime(RU_effective_dates_str)
    RU_effective_dates = to_datetime(RU_effective_dates.dt.strftime("%d/%m/%Y"))
    RU_effective_date = RU_effective_dates[0]

    # normalize the rates
    RU_normalizers = [int(RU_etree[i][2].text) for i in range(RU_number_of_rates)]
    RU_rates  = RU_rates.div(RU_normalizers)

    RU_data = concat([RU_base_cur, RU_cur_in_scope, RU_effective_dates, RU_rates], axis=1)

    # as in the original VBA script
    RU_out_of_scope_rates = ["XDR", "XAU"]
    RU_data = RU_data[~(RU_data.iloc[:, 1].isin(RU_out_of_scope_rates))]
    RU_data.replace("TMT", "TMM", inplace=True)

    RU_output_path = output_folder + r"\RUSSIA_RATES_" + str(RU_effective_date)[:-9] + ".xlsx"
    RU_header = generate_header("RU")
    generate_excel_output(RU_header, RU_data, RU_output_path, "RU")


##########################################
################ POLAND A ################
##########################################


def generate_PL_A():

    PL_A_etree = xml_to_element_tree("http://www.nbp.pl/kursy/xml/LastA.xml", "PL A")
    if PL_A_etree == 1:
        return

    PL_A_number_of_rates = len(PL_A_etree.getchildren()) - 2
    PL_A_base_cur = Series(["PLN" for _rate in range(PL_A_number_of_rates)])
    PL_A_cur_in_scope = Series([PL_A_etree[i][2].text for i in range(2, PL_A_number_of_rates + 2)])

    # retrieve the rates in string format and convert to float
    PL_A_rates_txt = Series([PL_A_etree[i][-1].text for i in range(2, PL_A_number_of_rates + 2)])
    PL_A_rates = PL_A_rates_txt.str.replace(",", ".").apply(lambda x: float(x))

    # denormalize
    PL_A_normalizers = [int(PL_A_etree[i][1].text) for i in range(2, PL_A_number_of_rates + 2)]
    PL_A_rates  = PL_A_rates.div(PL_A_normalizers)

    PL_A_effective_dates_str = Series([PL_A_etree[1].text for _rate in range(PL_A_number_of_rates)])
    PL_A_effective_dates = to_datetime(PL_A_effective_dates_str)
    PL_A_effective_date = PL_A_effective_dates[0]
    PL_A_effective_dates = to_datetime(PL_A_effective_dates.dt.strftime("%m/%d/%Y"))

    PL_A_data = concat([PL_A_base_cur, PL_A_cur_in_scope, PL_A_effective_dates, PL_A_rates], axis=1)

    PL_A_out_of_scope_rate = "XDR"
    PL_A_replacement_rates = {"AFN": "AFA", "GHS": "GHC", "MGA": "MGF", "MZN": "MZM", "SDG": "SDD", "SRD": "SRG", "ZWL": "ZWD"}
    PL_A_data = PL_A_data[PL_A_data.iloc[:, 1] != PL_A_out_of_scope_rate]
    PL_A_data.replace(PL_A_replacement_rates, inplace=True)

    PL_A_output_path = output_folder + r"\POLAND_A_RATES_" + str(PL_A_effective_date)[:-9] + ".xlsx"
    PL_A_header = generate_header("PL")
    generate_excel_output(PL_A_header, PL_A_data, PL_A_output_path, "PL A")


##########################################
################ POLAND B ################
##########################################


def generate_PL_B():
    PL_B_etree = xml_to_element_tree("http://www.nbp.pl/kursy/xml/LastB.xml", "PL B")
    if PL_B_etree == 1:
        return

    PL_B_number_of_rates = len(PL_B_etree.getchildren()) - 2
    PL_B_base_cur = Series(["PLN" for _rate in range(PL_B_number_of_rates)])
    PL_B_cur_in_scope = Series([PL_B_etree[i][2].text for i in range(2, PL_B_number_of_rates + 2)])

    # retrieve the rates in string format and convert to float
    PL_B_rates_txt = Series([PL_B_etree[i][-1].text for i in range(2, PL_B_number_of_rates + 2)])
    PL_B_rates = PL_B_rates_txt.str.replace(",", ".").apply(lambda x: float(x))

    # denormalize
    PL_B_normalizers = [int(PL_B_etree[i][1].text) for i in range(2, PL_B_number_of_rates + 2)]
    PL_B_rates  = PL_B_rates.div(PL_B_normalizers)

    PL_B_effective_dates_str = Series([PL_B_etree[1].text for _rate in range(PL_B_number_of_rates)])
    PL_B_effective_dates = to_datetime(PL_B_effective_dates_str)
    PL_B_effective_date = PL_B_effective_dates[0]
    PL_B_effective_dates = to_datetime(PL_B_effective_dates.dt.strftime("%m/%d/%Y"))

    PL_B_data = concat([PL_B_base_cur, PL_B_cur_in_scope, PL_B_effective_dates, PL_B_rates], axis=1)

    PL_B_replacement_rates = {"AFN": "AFA", "GHS": "GHC", "MGA": "MGF", "MZN": "MZM", "SDG": "SDD", "SRD": "SRG",
                              "ZWL": "ZWD", "ZMW": "ZMK"}
    PL_B_data.replace(PL_B_replacement_rates, inplace=True)

    PL_B_output_path = output_folder + r"\POLAND_B_RATES_" + str(PL_B_effective_date)[:-9] + ".xlsx"
    PL_B_header = generate_header("PL")
    generate_excel_output(PL_B_header, PL_B_data, PL_B_output_path, "PL B")


output_folder = generate_output_folder()

if datetime.today().hour > 12:
    generate_MA()
    generate_TR()
    generate_SK()
    generate_PL_A()
else:
    # only generate RU in the morning
    generate_RU()

# only generate PL_B rates on Wednesday
if datetime.today().weekday() == 2:
    generate_PL_B()

input("Press ENTER to enter the Matrix")

# TODO
# TODO check whether there is a need to upload/generate a given rate (maybe the rates were not updated since yesterday)
# -> either implement on files (search directory etc. or just catch an exception if uploading automatically)

# beautify the final date converting if statement - maybe isinstance(row[2].value, basestring)?
# add a log file? -- only if automating the upload to server as well
# add an auto-run
