import sys
import os
from time import strftime

import pandas as pd

import scrape_effort_list
from config import US_NAV, NY_NAV
from append_xls import append_xls
from download_from_site import download_from_site

def to_iso8601(x):
    '''Convert a datetime to ISO8601 string.'''
    return x.strftime('%Y-%m-%d') if (x > pd.datetime(1900,1,1)) else ''

if __name__ == "__main__":

    scrape_effort_list.main() # Download effort lists
    
    # If no arguments, we're done once we've downloaded and processed the lists!
    try:
        effort_cat = sys.argv[1]
    except IndexError:
        sys.exit() 
    
    # Figure out which effort list matches the type of download we want to do (6m, 12m, or DFLN)
    for filename in os.listdir('.'):
        if filename.startswith(effort_cat):
            break
        else:
            continue
    
    # Create list of which efforts we want results for
    packages_to_download = [line.strip().split("\t") for line in open(filename, 'rb').readlines()]

    us_pkgs_download = [pkg for org, pkg in packages_to_download if org==US_NAV['org']]
    ny_pkgs_download = [pkg for org, pkg in packages_to_download if org==NY_NAV['org']]

    # Download the results for US and NY.
    # We could do the downloads simultaneously (threads), but that causes Firefox troubles: two browsers using the same profile.

    # Download results to files
    download_from_site(US_NAV, us_pkgs_download)
    download_from_site(NY_NAV, ny_pkgs_download)

    # Rename results files, adding timestamp
    now_str = strftime('''%Y%m%d-%H%M%S''')
    us_filename = 'US Results ' + now_str + '.xls'
    ny_filename = 'NY Results ' + now_str + '.xls'
    os.rename(US_NAV['filename'], us_filename)
    os.rename(NY_NAV['filename'], ny_filename)

    # Load results into dataframes
    with pd.ExcelFile(us_filename) as xls:
        tabname = xls.sheet_names[0]
        df1 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    with pd.ExcelFile(ny_filename) as xls:
        tabname = xls.sheet_names[0]
        df2 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    # Remember the right order of columns, because 'append' alphasorts:
    with df1.columns as col_order: 
        df = df1.append(df2)
        df = df[col_order] 
    
    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)
        
    # Save merged results reports
    df.to_csv(os.path.splitext(filename)[0] + '.csv', encoding='utf-8', index_label='Mail Code')
    # We could export to XLSX instead but it's much slower:
    # df.to_excel(os.path.splitext(filename)[0] + ".xlsx")


