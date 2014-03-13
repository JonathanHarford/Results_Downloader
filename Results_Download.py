# TODO
# * There's only on sheet in each xls. We shouldn't need to store the tabname.
# * Script should attempt download until success
# * Downloads from both sites shouldn't be necessary
# * Store login in JSON file?

import sys
import os.path
import os
from time import strftime

from selenium import webdriver

from excelEnnumerations import *
from logininfo import *
from append_xls import *

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
    b.quit()

if __name__ == "__main__":

    filename = sys.argv[1]
    # Get list of which efforts we want results for
    packages_to_download = [line.strip().split("\t") for line in open(filename, 'rb').readlines()]

    us_pkgs_download = [pkg for org, pkg in packages_to_download if org==US_NAV['org']]
    ny_pkgs_download = [pkg for org, pkg in packages_to_download if org==NY_NAV['org']]

    # Download the results for US and NY.
    # We could do the downloads simultaneously (threads), but that causes its own troubles. (I wish I could remember what they are.)

    # Download results to files
    download_effort_results(US_NAV, us_pkgs_download)
    download_effort_results(NY_NAV, ny_pkgs_download)

    # Rename results files
    now_str = strftime('''%Y%m%d-%H%M%S''')
    us_filename = 'USO_US Results ' + now_str + '.xls'
    ny_filename = 'USO_NY Results ' + now_str + '.xls'
    os.rename(US_NAV['filename'], us_filename)
    os.rename(NY_NAV['filename'], ny_filename)

    # Merge the two files
    append_xls(us_filename, ny_filename, filename)


