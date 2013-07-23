'''
Package_List_Download
@author: jonathanharford

TODO:
* Split into modules, classes, functions
* "Attempting results download" "Reattempting results download (2)"
'''

# Standard library
import sys
import os
import argparse
import re
import datetime
import codecs

# Lovely packages others have written
import win32com.client
import cPickle
from selenium import webdriver
from selenium.webdriver.support.ui import WebDriverWait
from packages_table import PackagesTable

# My own modules
from excelEnnumerations import *
from logininfo import *

def create_browser():
    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.dir", os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")
    print("Creating browser...")
    # Get local session of Firefox
    b = webdriver.Firefox(firefox_profile=fp)
    b.implicitly_wait(10)
    return b

def download_package_list(b, nav):

    print("Browsing to " + nav['URL'])
    b.get(nav['URL']) # Load page

    # Page: Login
    ## wait_until_element_loads(b, b.title)
    assert "Untitled Page" in b.title

    b.find_element_by_id(nav['username']).send_keys(USERNAME)
    b.find_element_by_id(nav['password']).send_keys(PASSWORD)
    b.find_element_by_id(nav['login_btn']).click()

    # Page: "HOME"
    assert "HOME" in b.title

    # Pull FF date out of string like "Date on File: 04-02-2013"
    match = re.search("Date on File: (\d\d)-(\d\d)-(\d\d\d\d)", b.page_source)
    ff_date = datetime.date(int(match.group(3)), int(match.group(1)), int(match.group(2)))

    b.find_element_by_id(nav['reports_lnk']).click()

    # Page: "Reports Home"
    assert "Reports Home" in b.title
    b.find_element_by_id(nav['mtre_lnk']).click()

    # Page: "Reports Home"
    assert "Untitled Page" in b.title

    findall = re.findall(r"""<span[^>]*>([^<]*)</span>\s*</td><td>([^<]*)</td><td>(\d+)/(\d+)/(\d+)</td>""", b.page_source)
    table = []
    for package, packcode, mailmonth, mailday, mailyear in findall:
        maildate = datetime.date(int(mailyear), int(mailmonth), int(mailday))
        if packcode == u'\xa0': packcode = ""     # Why isn't packcode "&nbsp;"?
        table.append((package, nav['org'], packcode, maildate, ff_date))
    return (table, ff_date.isoformat())

def download_fake_package_list(b, nav):
    """For testing purposes. The download of data works fine, so instead of
    going to the website every time, we can use this function instead."""

# # # How to pickle these tables:
# with file("us_table.pickle", "wb") as f:  cPickle.dump(us_table, f)
# with file("ny_table.pickle", "wb") as f: cPickle.dump(ny_table, f)

    if   nav==NY_NAV:
        with file("test/us_table.pickle", "rb") as f:
            return (cPickle.load(f), u'2013-04-24')
    elif nav==US_NAV:
        with file("test/ny_table.pickle","rb") as f:
            return (cPickle.load(f), u'2013-04-22')

if __name__ == "__main__":

    fake_browser = "no_browser" in sys.argv
    if fake_browser:
        create_browser = lambda : False
        download_package_list = download_fake_package_list

    b = create_browser()

    t = PackagesTable("Package", "Org", "Pack Code", "Mail Date", "FF Date")
    ( us_table,  us_ff_str) = download_package_list(b, US_NAV)
    for row in us_table: t.addrecord(row)

    (ny_table, ny_ff_str) = download_package_list(b, NY_NAV)
    for row in ny_table: t.addrecord(row)

    # Save a table of every effort, even the ones we don't care about.
    filename = "All Packages " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".csv"

    t.writetocsv(filename)

    t.filter_efforts()

    # Save a lovely table of all efforts we might deal with.
    filename = "All Relevant Efforts " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".csv"

    t.writetocsv(filename)

    # Save a list of all efforts 6 months old and younger.
    filename = "6m " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".txt"

    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "6m":
                f.write(rec[0] + "\n")

    # Save a list of all efforts between 6 and 12 months old.
    filename = "12m " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".txt"

    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "12m":
                f.write(rec[0] + "\n")

    # Save a list of all DFLN efforts.
    filename = "DFLN " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".txt"

    with open(filename, "w") as f:
        for rec in t:
            if rec[6] in ["PI", "FSI", "NI"] and (rec[4].year >= 2011) :
                f.write(rec[0] + "\n")

    if not fake_browser: b.quit()
