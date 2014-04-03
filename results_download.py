#! python3
"""
results_download: Download and combine results spreadsheets from a number of (very similar) websites.

Usage: results_download.py [--pkgset=<pkgset>|--pkglist=<filename>] [-j | --dontjoin] [-k | --keeplists] [-l | --keepdl] [--quiet]
       results_download.py (-h | --help)

Options:
  -h --help             Show this screen.
  -j --dontjoin         Don't join costs and list descriptions to the dataframe
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

from docopt import docopt

from scrape_pkgs import scrape_pkgs
from create_pkg_tables import create_pkg_tables
from download_from_site import download_from_site
from merge_raw_reports import merge_raw_reports
from report_to_csv import report_to_csv
from join_cpps import join_cpps


# Load Configuration
from config import SITES # @UnresolvedImport

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
    outputfilename = os.path.splitext(pkglist)[0]
    
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

    # Merge them into a Dataframe
    df = merge_raw_reports(raw_report_fns, args['--keepdl'])

    logging.info('Saving raw results...')    
    report_to_csv(df.set_index("Mail Code"), outputfilename + ' RAW')

    df = join_cpps(df)

    logging.info('Saving processed results...')
    report_to_csv(df.set_index("Mailcode"), outputfilename)
