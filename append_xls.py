import sys
import os.path
import pandas as pd

def series_to_iso8601(ser):
    """Return a series of ISO 8601 dates from a series of datetimes."""

    def to_iso8601(x):
        return x.strftime('%Y-%m-%d') if (x > pd.datetime(1900,1,1)) else ''

    return ser.apply(to_iso8601)

def append_xls(fname1, fname2, filename):
    """Merges two spreadsheets into one."""

    # Load results into dataframes
    with pd.ExcelFile(fname1) as xls:
        tabname = xls.sheet_names[0]
        df1 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    with pd.ExcelFile(fname2) as xls:
        tabname = xls.sheet_names[0]
        df2 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    # Remember the right order of columns, because 'append' alphasorts:
    col_order = df1.columns

    df = df1.append(df2)

    df = df[col_order] # Fix the order of cols.

    df['Mail Date']  = series_to_iso8601(df['Mail Date'])
    df['FF Date']    = series_to_iso8601(df['FF Date'])
    df['First Date'] = series_to_iso8601(df['First Date'])
    df['Pack Date']  = series_to_iso8601(df['Pack Date'])

    df.to_csv(os.path.splitext(filename)[0] + '.csv', encoding='utf-8', index_label='Mail Code')
    # We could export to XLSX instead but it's much slower:
    # df.to_excel(os.path.splitext(filename)[0] + ".xlsx")

if __name__ == "__main__":
    fname1, fname2, filename = sys.argv[1:4]
    append_xls(fname1, fname2, filename)