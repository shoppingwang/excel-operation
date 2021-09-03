import xlsxwriter

SHEET_NAME = 'Analysis'
HEADERS = ['Variable', 'Formula', 'What the number stands for']
ROWS = [
    ['ITM_ID', '=COUNTIF(Master!$B$2:$B$18866, "*")', 'These are the new entries created by coders'],
    ['MEASURE_NOTE', '=COUNTIF(Master!$N$2:$N$18866, "*Duplicate entry*")', '# of duplicate entries '],
    ['MEASURE_NOTE', '=COUNTIF(Master!$N$2:$N$18866, "*Not mandatory*")',
     '# of measures are not mandatory (e.g., recommended, include "maybe")'],
    ['DECISION_MAKER', '=COUNTIF(Master!$Y$2:$Y$18866, "*Non-government*")', '# of measures are imposed by non-government decision makers'],
    ['DECISION_MAKER', '=COUNTIF(Master!$Y$2:$Y$18866, "*Government*")', '# of measures are imposed by central/local government'],
    ['DATE_NOTE', '=COUNTA(Master!$AC$2:$AC$18866)',
     '# of entries for which coders found that the start date information provided by original PHSM was wrong'],
    ['OTHER_NOTE', '=COUNTIF(Master!$AB$2:$AB$18866, "*Date*")',
     '# of entrie of which "OTHER_NOTE" column contains "Date" (the majority of them are corrections on the start date)'],
    ['LINK_NOTE', '=COUNTA(Master!$D$2:$D$18866)', '# of entries of which the original source link was not accessible/useful'],
    ['MEASURE_DESCRIPTION', '=COUNTIF(Master!$J$2:$J$18866, "*")',
     '# of entries about which coders found a better or the accurate measure description '],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Not a measure*")',
     '# of entries which are not an international travel measure based on coding protocol'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Providing travel advice or warning*")',
     '# of entries which are "Providing health advice or warning" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Testing*")', '# of entries which are "Testing" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Quarantine*")', '# of entries which are "Quarantine" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Other measures*")', '# of entries which are "Other measures"'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Suspending or restricting flights*")',
     '# of entries which are "Suspending or restricting flights" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Health screening*")', '# of entries which are "Health screening" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Closing borders*")', '# of entries which are "Closing borders" measures'],
    ['MEASURE_TYPE', '=SUM(COUNTIF(Master!$L$2:$L$18866, {"*Country-based measures*","*Country-based travel restriction/ban*"}))',
     '# of entries which are "Country-based travel restriction/ban" measures'],
    ['MEASURE_TYPE', '=SUM(COUNTIF(Master!$L$2:$L$18866, {"*Individual-based measures*","*Individual-based travel restriction/ban*"}))',
     '# of entries which are "Individual-based travel restriction/ban" measures'],
    ['MEASURE_TYPE',
     '=SUM(COUNTIF(Master!$L$2:$L$18866, {"*Suspending or restricting other travel mechanisms*","*Suspending or restricting other transportation modes*"}))',
     '# of entries which are "Suspending other transportation modes" measures (not international flights)'],
    ['MEASURE_TYPE', '=SUM(COUNTIF(Master!$L$2:$L$18866, {"*Route-based measures*","*Route-based travel restriction/ban*"}))',
     '# of entries which are "Route-based travel restriction/ban" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$18866, "*Additional travel documents*")',
     '# of entries which are "Additional travel documents" '],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$18866, {"*Extension*","*Extesnsion*"}))',
     '# of entries which is an extension of a previous entry'],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$18866, {"*Strengthening*","*Strenthening*"}))',
     '# of entries which is a strengthening of a previous entry'],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$18866, {"*Finished*","*Finish*"}))', '# of entries which is an end of a previous entry'],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$18866, {"*Easing*","*Phase-out*"}))',
     '# of entries which is a relaxation of a previous entry'],
    ['UPDATE_TYPE', '=COUNTIF(Master!$P$2:$P$18866, "*New*")', '# of entries which is an original entry']
]


def create_analysis_sheet(file_path: str, row_count: int):
    with xlsxwriter.Workbook(file_path) as workbook:
        worksheet = workbook.add_worksheet(name=SHEET_NAME)
        worksheet.write(0, 0, 1234)  # Writes an int
        worksheet.write(1, 0, 1234.56)  # Writes a float
        worksheet.write(2, 0, 'Hello')  # Writes a string
        worksheet.write(3, 0, None)  # Writes None
        worksheet.write(4, 0, True)  # Writes a bool
