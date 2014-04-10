#! python3
"""
download_from_site: Download results from a number of (very similar) websites.

Usage: download_from_site.py <filename> | --pkgset=<pkgset> [-l | --keepdl] [--quiet]  
       download_from_site.py (-h | --help)

Options:
  -h --help             Show this screen.
  <filename>            Download results in filename's list of packages.
  --pkgset=<pkgset>     Download <set> of results, e.g. "6m", "12m", "DFLN"
  -l --keepdl           Don't delete original downloads after processing.
  --quiet               print less text
"""

import os
import time
from time import strftime
import logging
import csv

import pandas as pd
from docopt import docopt
from selenium import webdriver

from find_latest_file import find_latest_file
from report_to_csv import report_to_csv
# Load Configuration
from config import SITES, RESULTS_COLS # @UnresolvedImport
#from excel_ennumerations import *

def download_from_site(nav, packages_to_download, DOWNLOAD_ATTEMPT_DURATION = 180, NUM_DOWNLOAD_ATTEMPTS = 10):

    if not packages_to_download:
        return False

    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)

    # Keep download bubble from appearing
    fp.set_preference("browser.download.panel.shown", True)
    fp.set_preference("browser.download.dir", os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")

    logging.info("Creating browser...")
    b = webdriver.Firefox(firefox_profile=fp) # Get local session of Firefox
    b.implicitly_wait(120)

    logging.info(("Browsing to " + nav['URL']))
    b.get(nav['report_url']) # Load page

    # Page: Login
    assert "Untitled Page" in b.title
    b.find_element_by_id(nav['username_fld']).send_keys(nav['username'])
    b.find_element_by_id(nav['password_fld']).send_keys(nav['password'])
    b.find_element_by_id(nav['login_btn']).click()

    assert "Untitled Page" in b.title

    # Let's tick the checkboxes for EVERY package whose results we'd like to download.
    b.implicitly_wait(0)
    for package_name in packages_to_download:
        my_xpath = "//span[text()='" + package_name + "']/../../td/input"
        try:
            el = b.find_element_by_xpath(my_xpath)
            el.click()
            logging.info(("TICKED:    " + package_name))
        except:
            logging.info(("NOT FOUND: " + package_name))
    b.implicitly_wait(120)

    # Click "Add"
    b.find_element_by_id(nav['add_btn']).click()

    # Click "Export Records"
    logging.info("Exporting...")

    file_is_downloaded = False

    for try_num in range(NUM_DOWNLOAD_ATTEMPTS):

        # Delete partially downloaded files
        try:
            os.remove(nav['filename'] + '.part')
        except WindowsError:
            pass

        # Delete previously downloaded files
        try:
            os.remove(nav['filename'])
        except WindowsError:
            pass

        # Start the timer
        attempt_start_time = time.clock()
        attempt_end_time = attempt_start_time + DOWNLOAD_ATTEMPT_DURATION

        logging.info(('Download attempt {} at {}').format(try_num + 1, strftime('''%H:%M:%S''')))

        b.find_element_by_id(nav['export_btn']).click()

        # Unfortunately, there doesn't seem to be a way to babysit the download of the file with Selenium.

        while time.clock() < attempt_end_time:
            try:
                filesize = os.path.getsize(nav['filename'])
                if filesize > 1:
                    logging.info('Download successful.')
                    file_is_downloaded = True
                    break
            except WindowsError:
                # Keep waiting...
                pass

        if file_is_downloaded:
            break
        else:
            logging.info('Timed out.')

    logging.info("Closing browser...")
    b.quit()
    
    new_filename = nav['org'] + ' Results FF ' + strftime('''%Y-%m-%d %H%M%S''') + '.xls'
    os.rename(nav['filename'], new_filename)
    
    return new_filename

def merge_raw_reports(raw_report_fns, keepdl=False):

    df = pd.DataFrame(columns=RESULTS_COLS)
    logging.info("Opening reports...")

    for fn in raw_report_fns:
        with pd.ExcelFile(fn) as xls:
            tabname = xls.sheet_names[0]
            df = df.append(xls.parse(tabname, index=None, parse_dates=['FF Date','Mail Date']))
        if not keepdl:
            os.remove(fn)
            
    df = df[RESULTS_COLS] # Reorder columns
    
    return df 

def main(args):
    pkglist = args['<filename>']
    if pkglist is None:
 
        # Figure out which package list matches the type of download we want to do (6m, 12m, or DFLN)
        pkgset = args['--pkgset']
        pkglist = find_latest_file(pkgset)
            
        logging.info("Working with " + pkglist)

    # Create list of which packages we want results for
    packages_to_download = [line.strip().split("\t") for line in open(pkglist)]

    # Download the results for each site.
    # We could do the downloads simultaneously (threads), but that causes 
    # Firefox troubles from multiple browsers using the same profile.

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
    df.to_csv(os.path.splitext(pkglist)[0] + ' RAW.csv',
              index=None, 
              quoting=csv.QUOTE_NONNUMERIC,
              date_format='%Y-%m-%d')
    
    report_to_csv(df.set_index("Mail Code"), os.path.splitext(pkglist)[0] + ' RAW')

if __name__ == "__main__":
    args = docopt(__doc__)

    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')

    main(args)
