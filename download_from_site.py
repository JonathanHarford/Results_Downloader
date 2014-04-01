#! python3

import os
import time
from time import strftime
import logging

from selenium import webdriver

# Load Configuration
from config import USERNAME, PASSWORD  # @UnresolvedImport

#from excel_ennumerations import *

def download_from_site(nav, packages_to_download, DOWNLOAD_ATTEMPT_DURATION = 120, NUM_DOWNLOAD_ATTEMPTS = 10):

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
    b.find_element_by_id(nav['username']).send_keys(USERNAME)
    b.find_element_by_id(nav['password']).send_keys(PASSWORD)
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
                #sys.stdout.write('.')
                pass

        if file_is_downloaded:
            break
        else:
            logging.info('Timed out.')

    logging.info("Closing browser...")
    b.quit()