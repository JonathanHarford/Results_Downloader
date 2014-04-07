#! python3
"""
scrape_pkgs: Scrape a list of packages from a list of websites.

Usage: scrape_pkgs.py [--quiet]
       scrape_pkgs.py (-h | --help)
       

Options:
  -h --help              Show this screen.
  --quiet                print less text
"""

# Standard library
import os
import re
import datetime
import logging

# Lovely packages others have written
from docopt import docopt
from selenium import webdriver

# My own modules
from packages_table import PackagesTable
from config import SITES # @UnresolvedImport

def create_pkg_tables(t, ff_strs):
    
    full_ff_str = ' '.join([org + ' FF ' + ff_str for org, ff_str in ff_strs])
        
    # Let's make a list of all filenames as we create them. 
    filenames = []

    # Save a table of every effort, even the ones we don't care about.
    filename = "All Packages " + full_ff_str + ".csv"

    filenames.append(filename)
    t.writetocsv(filename)

    t.filter_efforts()

    # Save a lovely table of all packages we might deal with.
    filename = "All Relevant Efforts "  + full_ff_str + ".csv"

    filenames.append(filename)
    t.writetocsv(filename)

    # Save a list of all packages 6 months old and younger.
    filename = "6m " + full_ff_str + ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "6m":
                f.write(rec[1] + "\t" + rec[0] + "\n")

    # Save a list of all packages between 6 and 12 months old.
    filename = "12m " + full_ff_str + ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "12m":
                f.write(rec[1] + "\t" + rec[0] + "\n")

    # Save a list of all DFLN packages.
    filename = "DFLN " + full_ff_str + ".txt"
    
    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[6] in ["PIns", "FSI", "NIns"] and (rec[4].year >= 2011) :
                f.write(rec[1] + "\t" + rec[0] + "\n")

    return filenames

def scrape_pkgs():
    
    def create_browser():
        fp = webdriver.FirefoxProfile()
        fp.set_preference("browser.download.dir", os.getcwd())
        fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
        logging.info("Creating browser...")
        # Get local session of Firefox
        b = webdriver.Firefox(firefox_profile=fp)
        b.implicitly_wait(10)
        return b    
    
    def scrape_package_list(b, nav):

        logging.info(("Browsing to " + nav['URL']))
        b.get(nav['URL']) # Load page
    
        # Page: Login
        ## wait_until_element_loads(b, b.title)
        assert "Untitled Page" in b.title
    
        b.find_element_by_id(nav['username_fld']).send_keys(nav['username'])
        b.find_element_by_id(nav['password_fld']).send_keys(nav['password'])
        b.find_element_by_id(nav['login_btn']).click()
    
        # Page: "HOME"
        assert "HOME" in b.title
    
        # Pull FF date out of string like "Date on File: 04-02-2013"
        match = re.search("Date on File: (\d\d)-(\d\d)-(\d\d\d\d)", b.page_source)
        ff_date = datetime.date(int(match.group(3)), int(match.group(1)), int(match.group(2)))
    
        b.get(nav['report_url'])
        assert "Untitled Page" in b.title
    
        findall = re.findall(r"""<span[^>]*>([^<]*)</span>\s*</td><td>([^<]*)</td><td>(\d+)/(\d+)/(\d+)</td>""", b.page_source)
        table = []
        for package, packcode, mailmonth, mailday, mailyear in findall:
            maildate = datetime.date(int(mailyear), int(mailmonth), int(mailday))
            if packcode == '\xa0': packcode = ""     # Why isn't packcode "&nbsp;"?
            table.append((package, nav['org'], "", packcode, maildate, ff_date, "", "", ""))
        return (table, ff_date.isoformat())
    
    b = create_browser()

    t = PackagesTable("Package",
                      "Org",
                      "Eff",
                      "Pack Code",
                      "Mail Date",
                      "FF Date",
                      "EffType",
                      "Age",
                      "Results Group")
    
    ff_strs = []
    
    for site in SITES:
        (tbl, ff_str) = scrape_package_list(b, site)
        ff_strs.append((site['org'], ff_str))
        for row in tbl:
            t.addrecord(row)

    b.quit()
    
    return t, ff_strs

def main(args):
    if not args['--quiet']:
        logging.basicConfig(level=logging.INFO, 
                            format='%(asctime)s %(levelname)-6s %(message)s',
                            datefmt='%H:%M:%S')
    t, ff_strs = scrape_pkgs()
    return create_pkg_tables(t, ff_strs)

if __name__ == "__main__":
    args = docopt(__doc__)
    main(args)




