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

import os
import logging

import pandas as pd

# Load Configuration
from config import RESULTS_COLS  # @UnresolvedImport

def report_to_csv(df, filename):
    """Convert dates to strings, and make a CSV with 'Mail Code' as the index."""

    def to_iso8601(dt):
        '''Convert a datetime to ISO8601 string.'''
        return dt.strftime('%Y-%m-%d') if (dt > pd.datetime(1900,1,1)) else ''
    
    # Dates are ugly unless we do this:
    for col in ('Mail Date', 'FF Date', 'First Date', 'Pack Date'):
        df[col]  = df[col].apply(to_iso8601)
        
    # CSV is much faster than XLSX:
    df.to_csv(filename, encoding='utf-8', index_label='Mail Code')

def merge_raw_reports(raw_report_fns, savename, keepdl=False):

    df = pd.DataFrame()
    logging.info("Opening reports...")

    for fn in raw_report_fns:
        with pd.ExcelFile(fn) as xls:
            tabname = xls.sheet_names[0]
            df = df.append(xls.parse(tabname, index=None, parse_dates=['FF Date','Mail Date']))

        if not keepdl:
            os.remove(fn)
    
    df = df[RESULTS_COLS].set_index(['Mail Code']) # Set index AND reorder columns
    
    logging.info('Saving merged results reports...')    
    report_to_csv(df, savename)

    return df


