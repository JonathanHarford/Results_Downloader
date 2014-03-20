import sys
import os.path
import pandas as pd

def append_xls(us_filename, ny_filename, filename):
    """Merges two single-sheet spreadsheets into one."""

    def to_iso8601(x):
        '''Convert a datetime to ISO8601 string.'''
        return x.strftime('%Y-%m-%d') if (x > pd.datetime(1900,1,1)) else ''

    # Load results into dataframes
    with pd.ExcelFile(us_filename) as xls:
        tabname = xls.sheet_names[0]
        df1 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    with pd.ExcelFile(ny_filename) as xls:
        tabname = xls.sheet_names[0]
        df2 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    # Remember the right order of columns, because 'append' alphasorts:
    col_order = df1.columns

    df = df1.append(df2)

    df = df[col_order] # Fix the order of cols.

    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)
        
    df.to_csv(os.path.splitext(filename)[0] + '.csv', encoding='utf-8', index_label='Mail Code')
    # We could export to XLSX instead but it's much slower:
    # df.to_excel(os.path.splitext(filename)[0] + ".xlsx")

if __name__ == "__main__":
    fname1, fname2, filename = sys.argv[1:4]
    append_xls(fname1, fname2, filename)