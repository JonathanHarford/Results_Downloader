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

def merge_raw_reports(raw_report_fns, keepdl=False):

    df = pd.DataFrame()
    logging.info("Opening reports...")

    for fn in raw_report_fns:
        with pd.ExcelFile(fn) as xls:
            tabname = xls.sheet_names[0]
            df = df.append(xls.parse(tabname, index=None, parse_dates=['FF Date','Mail Date']))

        if not keepdl:
            os.remove(fn)
    
    return df[RESULTS_COLS].set_index(['Mail Code']) # Set index AND reorder columns
