"""
prepare_morocco()

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
print(rates_MA)

# file path + name of the file
# use an elastic one in the future -- settings file?
MA_title = r"..\Upload_rates\Morocco Rates\MOROCCO_RATES\MOROCCO_RATES_" + str(MA_effective_date)[:-9] + ".xlsx"

# create the header as a separate DF; could use one DataFrame once MultiIndex columns are better supported
header = pd.DataFrame([["CURRENCY_RATES", "COMPANY_ID=HP", "SOURCE=BOM-MAD", ""],
                       ["BASE_CURRENCY", "FOREIGN_CURRENCY", "EFFECTIVE_DATE", "RATE"]])

# create the final xlsx
with pd.ExcelWriter(MA_title, engine="openpyxl") as writer:
    header.to_excel(writer, index=False, header=False)
    rates_MA.to_excel(writer, index=False, header=False, startrow=2)

# format the date as the bare int format is treated as General, and we need it to be an Excel Date type
wb_MA = openpyxl.load_workbook(title_MA)
ws_MA = wb_MA.active
for row in ws_MA:
    if "A" not in str((row[2]).value):  # skip header rows, picked "A" because column C headers have it :)
        row[2].number_format = "mm-dd-yy"
wb_MA.save(title_MA)

# TODO: create a settings file with the destination folder for the output file
# if the directory does not exist - create it
# beautify the final date converting if statement - maybe isinstance(row[2].value, basestring)?
# refactor

# cleanup
remove("VATSPOTR.txt")
"""
