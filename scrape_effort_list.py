# Standard library
import os
import re
import datetime

# Lovely packages others have written
from selenium import webdriver

# My own modules
from packages_table import PackagesTable

# Load Configuration
import json
config = json.load(open('config.json'))
USERNAME = config['USERNAME']
PASSWORD = config['PASSWORD']
US_NAV = config['US_NAV']
NY_NAV = config['NY_NAV']

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

    print(("Browsing to " + nav['URL']))
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

    b.get(nav['report_url'])
    assert "Untitled Page" in b.title

    findall = re.findall(r"""<span[^>]*>([^<]*)</span>\s*</td><td>([^<]*)</td><td>(\d+)/(\d+)/(\d+)</td>""", b.page_source)
    table = []
    for package, packcode, mailmonth, mailday, mailyear in findall:
        maildate = datetime.date(int(mailyear), int(mailmonth), int(mailday))
        if packcode == '\xa0': packcode = ""     # Why isn't packcode "&nbsp;"?
        table.append((package, nav['org'], "", packcode, maildate, ff_date, "", "", ""))
    return (table, ff_date.isoformat())

def main():
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
    
    (us_table, us_ff_str) = download_package_list(b, US_NAV)
    for row in us_table:
        t.addrecord(row)

    (ny_table, ny_ff_str) = download_package_list(b, NY_NAV)
    for row in ny_table:
        t.addrecord(row)

    b.quit()

    # Let's make a list of all filenames as we create them. 
    filenames = []

    # Save a table of every effort, even the ones we don't care about.
    filename = "All Packages " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".csv"

    filenames.append(filename)
    t.writetocsv(filename)

    t.filter_efforts()

    # Save a lovely table of all efforts we might deal with.
    filename = "All Relevant Efforts " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".csv"

    filenames.append(filename)
    t.writetocsv(filename)

    # Save a list of all efforts 6 months old and younger.
    filename = "6m " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "6m":
                f.write(rec[1] + "\t" + rec[0] + "\n")

    # Save a list of all efforts between 6 and 12 months old.
    filename = "12m " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "12m":
                f.write(rec[1] + "\t" + rec[0] + "\n")

    # Save a list of all DFLN efforts.
    filename = "DFLN " + \
               "US FF "  + us_ff_str + \
               " NY FF " + ny_ff_str + \
               ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[6] in ["PIns", "FSI", "NIns"] and (rec[4].year >= 2011) :
                f.write(rec[1] + "\t" + rec[0] + "\n")

    return filenames
    

if __name__ == "__main__":
    main()
    