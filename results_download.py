#! python3
"""
Usage: results_download.py [--pkgset=<pkgset>|--pkglist=<filename>] [-k | --keeplists] [-l | --keepdl] [--quiet]
       results_download.py (-h | --help)

Options:
  -h --help             Show this screen.
  -k --keeplists        Don't delete scraped lists.
  -l --keepdl           Don't delete original downloads after processing.  
  --pkgset=<pkgset>     Download <set> of results, e.g. "6m", "12m", "DFLN"
  --pkglist=<filename>  Use a particular list of packages instead of a 
                        standard results set.
  --quiet               print less text              
"""


import sys
import os
from time import strftime
import logging

import pandas as pd
from docopt import docopt

from create_pkg_tables import create_pkg_tables
from scrape_pkgs import scrape_pkgs
from download_from_site import download_from_site
# import join_cpps 

# Load Configuration
from config import US_NAV, NY_NAV, RESULTS_COLS  # @UnresolvedImport

def to_iso8601(x):
    '''Convert a datetime to ISO8601 string.'''
    return x.strftime('%Y-%m-%d') if (x > pd.datetime(1900,1,1)) else ''

if __name__ == "__main__":
    args = docopt(__doc__)

    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')
    
    pkglist = args['--pkglist']
    if pkglist is None:
 
        t, us_ff_str, ny_ff_str = scrape_pkgs()
        pkglists = create_pkg_tables(t, us_ff_str, ny_ff_str)
        
        # If no set (or pkglist), we're done once we've downloaded and processed the lists!
        pkgset = args['--pkgset']
        if pkgset is None:
            sys.exit()
        
        # Figure out which package list matches the type of download we want to do (6m, 12m, or DFLN)
        for filename in os.listdir('.'):
            if filename.startswith(pkgset):
                pkglist = filename
                break

        # Delete scraped lists (unless we shouldn't)
        if not args['--keeplists']:
            for filename in pkglists:
                if filename != pkglist:
                    os.remove(filename)
        
    logging.info("Working with " + pkglist)
    
    # Create list of which packages we want results for
    packages_to_download = [line.strip().split("\t") for line in open(pkglist)]

    us_pkgs_download = [pkg for org, pkg in packages_to_download if org==US_NAV['org']]
    ny_pkgs_download = [pkg for org, pkg in packages_to_download if org==NY_NAV['org']]

    # Download the results for US and NY.
    # We could do the downloads simultaneously (threads), but that causes 
    # Firefox troubles from two browsers using the same profile.

    # Download results to files
    
    download_from_site(US_NAV, us_pkgs_download)
    download_from_site(NY_NAV, ny_pkgs_download)

    # Rename results files, adding timestamp
    logging.info("Renaming reports...")
    now_str = strftime('''%Y%m%d-%H%M%S''')
    us_filename = 'US Results ' + now_str + '.xls'
    ny_filename = 'NY Results ' + now_str + '.xls'
    

    df = pd.DataFrame()
    logging.info("Opening reports...")

    if us_pkgs_download:
        os.rename(US_NAV['filename'], us_filename)
        with pd.ExcelFile(us_filename) as xls:
            tabname = xls.sheet_names[0]
            df = df.append(xls.parse(tabname, index=None, parse_dates=['FF Date','Mail Date']))

        if not args['--keepdl']:
            os.remove(us_filename)
            
    if ny_pkgs_download:
        os.rename(NY_NAV['filename'], ny_filename)
        with pd.ExcelFile(ny_filename) as xls:
            tabname = xls.sheet_names[0]
            df = df.append(xls.parse(tabname, index=None, parse_dates=['FF Date','Mail Date']))
        
        if not args['--keepdl']:
            os.remove(ny_filename)
    
    df = df[RESULTS_COLS].set_index(['Mail Code'])
    
    logging.info('Saving merged results reports...') 
   
    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)

    # CSV is much faster than XLSX        
    df.to_csv(os.path.splitext(pkglist)[0] + ' RAW.csv', encoding='utf-8', index_label='Mail Code')



