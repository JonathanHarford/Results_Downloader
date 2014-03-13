# TODO
# * Script should attempt download until success
# * What to do if there's already a USO_EXPORT_MAIL_TABLE.xls in the dir?
# * Downloads from both sites shouldn't be necessary
# * Store login in JSON file?

import sys
import os.path
import os
import time
from time import strftime

from selenium import webdriver

from excelEnnumerations import *
from logininfo import *
from append_xls import *

DOWNLOAD_ATTEMPT_DURATION = 120 # Give each download this long (seconds) to complete.
NUM_DOWNLOAD_ATTEMPTS = 10 # Attempt each download this many times.

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
    b.get(nav['report_url']) # Load page

    # Page: Login
    assert "Untitled Page" in b.title
    b.find_element_by_id(nav['username']).send_keys(USERNAME)
    b.find_element_by_id(nav['password']).send_keys(PASSWORD)
    b.find_element_by_id(nav['login_btn']).click()

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

        print('Download attempt {} at {}').format(try_num + 1, strftime('''%H:%M:%S'''))

        b.find_element_by_id(nav['export_btn']).click()

        # Unfortunately, there doesn't seem to be a way to babysit the download of the file with Selenium.

        while time.clock() < attempt_end_time:
            try:
                filesize = os.path.getsize(nav['filename'])
                if filesize > 1:
                    print('Download successful.')
                    file_is_downloaded = True
                    break
            except WindowsError:
                #sys.stdout.write('.')
                pass

        if file_is_downloaded:
            break
        else:
            print('Timed out.')

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


