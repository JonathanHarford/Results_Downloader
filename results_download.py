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
import logging

import pandas as pd
from docopt import docopt

from create_pkg_tables import create_pkg_tables
from scrape_pkgs import scrape_pkgs
from download_from_site import download_from_site
from merge_raw_reports import merge_raw_reports
# import join_cpps 

# Load Configuration
from config import SITES # @UnresolvedImport

def to_iso8601(x):
    '''Convert a datetime to ISO8601 string.'''
    return x.strftime('%Y-%m-%d') if (x > pd.datetime(1900,1,1)) else ''

def report_to_csv(df, filename):
    """Convert dates to strings, and make a CSV with 'Mail Code' as the index."""
    
    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)
        
    # CSV is much faster than XLSX:
    df.to_csv(filename, encoding='utf-8', index_label='Mail Code')

if __name__ == "__main__":
    args = docopt(__doc__)

    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')
    
    pkglist = args['--pkglist']
    if pkglist is None:
 
        t, ff_strs = scrape_pkgs()
        pkglists = create_pkg_tables(t, ff_strs)
        
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

    # Download the results for US and NY.
    # We could do the downloads simultaneously (threads), but that causes 
    # Firefox troubles from two browsers using the same profile.

    # Download results to files
    
    raw_report_fns = []
    for site in SITES:
        site_pkgs = [pkg for org, pkg in packages_to_download if org==site['org']]
        filename = download_from_site(site, site_pkgs)
        if filename:
            raw_report_fns.append(filename)

    df = merge_raw_reports(raw_report_fns, os.path.splitext(pkglist)[0] + ' RAW.csv', args['--keepdl'])



