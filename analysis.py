import numpy as np
from pandas import DataFrame

import utils

SHEET_NAME = 'Analysis'
HEADERS = ['Variable', 'Formula', 'What the number stands for']
ROWS = [
    ['ITM_ID', '=COUNTIF(Master!$B$2:$B$%(last_row_number)s, "*")', 'These are the new entries created by coders'],
    ['MEASURE_NOTE', '=COUNTIF(Master!$N$2:$N$%(last_row_number)s, "*Duplicate entry*")', '# of duplicate entries '],
    ['MEASURE_NOTE', '=COUNTIF(Master!$N$2:$N$%(last_row_number)s, "*Not mandatory*")',
     '# of measures are not mandatory (e.g., recommended, include "maybe")'],
    ['DECISION_MAKER', '=COUNTIF(Master!$Y$2:$Y$%(last_row_number)s, "*Non-government*")',
     '# of measures are imposed by non-government decision makers'],
    ['DECISION_MAKER', '=COUNTIF(Master!$Y$2:$Y$%(last_row_number)s, "*Government*")',
     '# of measures are imposed by central/local government'],
    ['DATE_NOTE', '=COUNTA(Master!$AC$2:$AC$%(last_row_number)s)',
     '# of entries for which coders found that the start date information provided by original PHSM was wrong'],
    ['OTHER_NOTE', '=COUNTIF(Master!$AB$2:$AB$%(last_row_number)s, "*Date*")',
     '# of entrie of which "OTHER_NOTE" column contains "Date" (the majority of them are corrections on the start date)'],
    ['LINK_NOTE', '=COUNTA(Master!$D$2:$D$%(last_row_number)s)',
     '# of entries of which the original source link was not accessible/useful'],
    ['MEASURE_DESCRIPTION', '=COUNTIF(Master!$J$2:$J$%(last_row_number)s, "*")',
     '# of entries about which coders found a better or the accurate measure description '],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Not a measure*")',
     '# of entries which are not an international travel measure based on coding protocol'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Providing travel advice or warning*")',
     '# of entries which are "Providing health advice or warning" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Testing*")', '# of entries which are "Testing" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Quarantine*")', '# of entries which are "Quarantine" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Other measures*")', '# of entries which are "Other measures"'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Suspending or restricting flights*")',
     '# of entries which are "Suspending or restricting flights" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Health screening*")',
     '# of entries which are "Health screening" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Closing borders*")',
     '# of entries which are "Closing borders" measures'],
    ['MEASURE_TYPE',
     '=SUM(COUNTIF(Master!$L$2:$L$%(last_row_number)s, {"*Country-based measures*","*Country-based travel restriction/ban*"}))',
     '# of entries which are "Country-based travel restriction/ban" measures'],
    ['MEASURE_TYPE',
     '=SUM(COUNTIF(Master!$L$2:$L$%(last_row_number)s, {"*Individual-based measures*","*Individual-based travel restriction/ban*"}))',
     '# of entries which are "Individual-based travel restriction/ban" measures'],
    ['MEASURE_TYPE',
     '=SUM(COUNTIF(Master!$L$2:$L$%(last_row_number)s, {"*Suspending or restricting other travel mechanisms*","*Suspending or restricting other transportation modes*"}))',
     '# of entries which are "Suspending other transportation modes" measures (not international flights)'],
    ['MEASURE_TYPE', '=SUM(COUNTIF(Master!$L$2:$L$%(last_row_number)s, {"*Route-based measures*","*Route-based travel restriction/ban*"}))',
     '# of entries which are "Route-based travel restriction/ban" measures'],
    ['MEASURE_TYPE', '=COUNTIF(Master!$L$2:$L$%(last_row_number)s, "*Additional travel documents*")',
     '# of entries which are "Additional travel documents" '],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$%(last_row_number)s, {"*Extension*","*Extesnsion*"}))',
     '# of entries which is an extension of a previous entry'],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$%(last_row_number)s, {"*Strengthening*","*Strenthening*"}))',
     '# of entries which is a strengthening of a previous entry'],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$%(last_row_number)s, {"*Finished*","*Finish*"}))',
     '# of entries which is an end of a previous entry'],
    ['UPDATE_TYPE', '=SUM(COUNTIF(Master!$P$2:$P$%(last_row_number)s, {"*Easing*","*Phase-out*"}))',
     '# of entries which is a relaxation of a previous entry'],
    ['UPDATE_TYPE', '=COUNTIF(Master!$P$2:$P$%(last_row_number)s, "*New*")', '# of entries which is an original entry']
]


def _convert_rows_data(row_count: int):
    rows = []
    for r in ROWS:
        rows.append([r[0], r[1] % {"last_row_number": str(row_count + 1)}, r[2]])
    return rows


def create_analysis_sheet(file_path: str, row_count: int):
    print(f"Write analysis results to {file_path}")
    data = _convert_rows_data(row_count)
    df = DataFrame(np.array(data), columns=HEADERS)
    utils.append_df_to_excel(file_path, df, sheet_name=SHEET_NAME, truncate_sheet=True, index=False)
