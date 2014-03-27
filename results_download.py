"""
Usage: results_download.py [--set=<set>|--efflist=<filename>] [-k | --keeplists] [-l | --keepdl]
       results_download.py (-h | --help)

Options:
  -h --help             Show this screen.
  -k --keeplists        Don't delete scraped lists.
  -l --keepdl           Don't delete original downloads after processing.  
  --set=<set>           Download <set> of results, e.g. "6m", "12m", "DFLN"
  --efflist=<filename>  Use a particular list of efforts instead of a 
                        standard results set.
"""
#  --config=<configfile> Use <configfile> (contains login info) instead 
#                        of config.py.
#  --quiet               print less text

import sys
import os
from time import strftime

import pandas as pd
from docopt import docopt

import scrape_effort_list
from config import US_NAV, NY_NAV
from download_from_site import download_from_site
# import join_cpps 

def to_iso8601(x):
    '''Convert a datetime to ISO8601 string.'''
    return x.strftime('%Y-%m-%d') if (x > pd.datetime(1900,1,1)) else ''

if __name__ == "__main__":
    args = docopt(__doc__)
    # print(args)
    
    efflist = args['--efflist']
    if efflist is None:
        scrapefiles = scrape_effort_list.main() # Download effort lists
    
        # If no set (or efflist), we're done once we've downloaded and processed the lists!
        effset = args['--set']
        if effset is None:
            sys.exit()
        
        # Figure out which effort list matches the type of download we want to do (6m, 12m, or DFLN)
        for filename in os.listdir('.'):
            if filename.startswith(effset):
                efflist = filename
                break

        # Delete scraped lists (unless we shouldn't)
        if not args['--keeplists']:
            for filename in scrapefiles:
                if filename != efflist:
                    os.remove(filename)
        
    print("Working with " + efflist)
    
    # Create list of which efforts we want results for
    packages_to_download = [line.strip().split("\t") for line in open(efflist, 'rb').readlines()]

    us_pkgs_download = [pkg for org, pkg in packages_to_download if org==US_NAV['org']]
    ny_pkgs_download = [pkg for org, pkg in packages_to_download if org==NY_NAV['org']]

    # Download the results for US and NY.
    # We could do the downloads simultaneously (threads), but that causes 
    # Firefox troubles from two browsers using the same profile.

    # Download results to files
    download_from_site(US_NAV, us_pkgs_download)
    download_from_site(NY_NAV, ny_pkgs_download)

    # Rename results files, adding timestamp
    print("Renaming reports...")
    now_str = strftime('''%Y%m%d-%H%M%S''')
    us_filename = 'US Results ' + now_str + '.xls'
    ny_filename = 'NY Results ' + now_str + '.xls'
    os.rename(US_NAV['filename'], us_filename)
    os.rename(NY_NAV['filename'], ny_filename)

    print("Opening reports...")
    with pd.ExcelFile(us_filename) as xls:
        tabname = xls.sheet_names[0]
        df1 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])
    
    with pd.ExcelFile(ny_filename) as xls:
        tabname = xls.sheet_names[0]
        df2 = xls.parse(tabname, index_col=1, parse_dates=['FF Date','Mail Date'])

    # Remember the right order of columns, because 'append' alphasorts:
    col_order = df1.columns
    print("Merging reports...") 
    df = df1.append(df2)
    df = df[col_order] 
    
    # Delete the downloaded files (unless we shouldn't)
    if not args['--keepdl']:
        os.remove(us_filename)
        os.remove(ny_filename)
    
    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)
        
    print('Saving merged results reports...') # CSV is much faster than XLSX
    df.to_csv(os.path.splitext(efflist)[0] + ' RAW.csv', encoding='utf-8', index_label='Mail Code')



