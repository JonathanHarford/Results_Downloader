#! python3
"""
Usage: merge_raw_reports.py [-l | --keepdl] [--quiet]
       results_download.py (-h | --help)

Options:
  -h --help             Show this screen.
  -l --keepdl           Don't delete original downloads after processing.  
  --quiet               print less text              
"""

import os
import logging

import docopt
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
    
    return df[RESULTS_COLS] # Reorder columns

def main(args):
    

if __name__ == "__main__":
    args = docopt(__doc__)

    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')

    main(args)