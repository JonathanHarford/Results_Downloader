#! python3

# Standard library
import os
import re
import datetime
import logging

# Lovely packages others have written
from selenium import webdriver

# My own modules
from packages_table import PackagesTable
from config import SITES # @UnresolvedImport

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
    