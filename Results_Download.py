import sys
import os.path
import os
from time import strftime

import pandas as pd
#import win32com.client
#import selenium
from selenium import webdriver

from excelEnnumerations import *
from logininfo import *

def series_to_iso8601(ser):
    """Return a series of ISO 8601 dates from a series of datetimes."""

    def to_iso8601(x):
        if x > pd.datetime(1900,1,1):
            return x.strftime('%Y-%m-%d')
        else:
            return ''

    return ser.apply(to_iso8601)

def download_effort_results(nav, efforts_to_download):
    if len(efforts_to_download) == 0:
        return False

    fp = webdriver.FirefoxProfile()
    fp.set_preference("browser.download.folderList", 2)
    fp.set_preference("browser.download.manager.showWhenStarting", False)

    # Keep download bubble from appearing
    fp.set_preference("browser.download.panel.shown", True)
    fp.set_preference("browser.download.dir", os.getcwd())
    fp.set_preference("browser.helperApps.neverAsk.saveToDisk", "application/vnd.ms-excel")

    print("Creating browser...")
    b = webdriver.Firefox(firefox_profile=fp) # Get local session of Firefox
    b.implicitly_wait(120)

    print("Browsing to " + nav['URL'])
    b.get(nav['URL']) # Load page

    # Page: Login
    assert "Untitled Page" in b.title
    b.find_element_by_id(nav['username']).send_keys(USERNAME)
    b.find_element_by_id(nav['password']).send_keys(PASSWORD)
    b.find_element_by_id(nav['login_btn']).click()

    # Page: "HOME"
    ## sleep(5) # Pages taking too long too load -- script breaking!
    assert "HOME" in b.title
    b.find_element_by_id(nav['reports_lnk']).click()

    # Page: "Reports Home"
    ## sleep(5) # Pages taking too long too load -- script breaking!
    assert "Reports Home" in b.title
    b.find_element_by_id(nav['mtre_lnk']).click()

    # Page: "Reports Home"
    assert "Untitled Page" in b.title

    # Let's tick the checkboxes for EVERY effort whose results we'd like to download.
    b.implicitly_wait(0)
    for effort_name in efforts_to_download:
        my_xpath = "//span[text()='" + effort_name + "']/../../td/input"
        try:
            el = b.find_element_by_xpath(my_xpath)
            el.click()
            print("TICKED:    " + effort_name)
        except:
            print("NOT FOUND: " + effort_name)
    b.implicitly_wait(120)

    # Click "Add"
    b.find_element_by_id(nav['add_btn']).click()

    # Click "Export Records"
    print("Exporting...")

    b.find_element_by_id(nav['export_btn']).click()

    # Unfortunately, there doesn't seem to be a way to babysit the download of the file with Selenium.

    print("Waiting for file to exist in directory...")
    while 1:
        try:
            filesize = os.path.getsize(nav['filename'])
            if filesize > 1:
                ##print(filesize)
                break
        except WindowsError:
            pass

    print("Closing browser...")
    # b.quit()

    df =   pd.read_excel(nav['filename'],
                         nav['tabname'],
                         parse_dates=['FF Date','Mail Date'],
                         index_col=1)
    os.rename(nav['filename'], nav['tabname'] + ' ' + strftime('''%Y%m%d-%H%M%S''') + '.xls')

    return df

if __name__ == "__main__":
    if len(sys.argv) == 1:
        filename = "testdata"
        print __doc__
        actually_download = False
    else: actually_download = True

    if actually_download:
        filename = sys.argv[1]
        # Get list of which efforts we want results for
        packages_to_download = [line.strip().split("\t") for line in open(filename, 'rb').readlines()]

        us_pkgs_download = [pkg for org, pkg in packages_to_download if org==US_NAV['org']]
        ny_pkgs_download = [pkg for org, pkg in packages_to_download if org==NY_NAV['org']]

        # Download the results for US and NY.
        # We could do the downloads simultaneously (threads), but that causes its own troubles.

    df =           download_effort_results(US_NAV, us_pkgs_download)

    # Remember the right order of columns, because 'append' alphasorts:
    col_order = df.columns

    df = df.append(download_effort_results(NY_NAV, ny_pkgs_download))

    df = df[col_order] # Fix the order of cols.

    # df = df.set_index("Mail Code", verify_integrity=True)
    # df = df[['Mail Code','DM Donors','DM Revenue','Descript','FF Date','Mail Date','ND Count','Package','Qty Mail','TM Donors','TM Revenue','Total Donors','Total Revenue','Web Donors','Web Revenue']]

    df['Mail Date']  = series_to_iso8601(df['Mail Date'])
    df['FF Date']    = series_to_iso8601(df['FF Date'])
    df['First Date'] = series_to_iso8601(df['First Date'])
    df['Pack Date'] = series_to_iso8601(df['Pack Date'])

    df.to_csv(os.path.splitext(filename)[0] + '.csv', encoding='utf-8', index_label='Mail Code')
    # XLSX is VERY slow compared to csv:
    #df.to_excel(os.path.splitext(filename)[0] + ".xlsx")


