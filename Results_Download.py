'''
Created on Mar 27, 2012

@author: jonathanharford
'''

import sys
import os.path
from time import sleep
import win32com.client
import selenium
from selenium import webdriver

from excelEnnumerations import *
from logininfo import *

def download_effort_results(nav, efforts_to_download):

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
                print(filesize)
                break
        except WindowsError:
            pass

    print("Closing browser...")
    # b.quit()

if __name__ == "__main__":
    if len(sys.argv) == 1:
        print __doc__
        actually_download = False
    else: actually_download = True

    if actually_download:

        # Get list of which efforts we want results for
        packages_to_download = [line.strip() for line in open(sys.argv[1], 'rb').readlines()]

        # Download the results for US and NY.
        # We could do the downloads simultaneously (threads), but that causes its own troubles.

        download_effort_results(US_NAV, packages_to_download)
        download_effort_results(NY_NAV, packages_to_download)

    # Open first the NY spreadsheet.
    xl = win32com.client.Dispatch("Excel.Application")
    xl.Visible = 1
    wb_tmp = xl.Workbooks.Open(os.path.join(os.getcwd(), NY_NAV['filename']))
    ws_tmp = wb_tmp.Sheets(1)

    # We have to insert some columns to make the format match the US results.
    xl.CutCopyMode = False
    ws_tmp.Range("N1").EntireColumn.Insert()
    ws_tmp.Range("T1").EntireColumn.Insert()

    # Copy the NY results.
    # ws_tmp.Range(ws_tmp.Range("Y65536").End(xlUp), ws_tmp.Range("A2")).Copy() # No.
    ws_tmp.Range(ws_tmp.Range("X65536").End(xlUp), ws_tmp.Range("A2")).Copy()

    # Now open the US spreadsheet and append the NY results, closing the NY file afterwards.
    wb = xl.Workbooks.Open(os.path.join(os.getcwd(),US_NAV['filename']))
    ws = wb.Sheets(1)
    ws.Range("A65536").End(xlUp).GetOffset(1, 0).Select()
    ws.Paste()

    # Save merged file
    wb.SaveAs(Filename=os.path.join(os.getcwd(),r"Results.xlsx"), FileFormat=xlOpenXMLWorkbook, CreateBackup=False)

    # Close and remove the old ones.
    xl.CutCopyMode = False # Clear clipboard so we can close spreadsheet without complaint.
    wb_tmp.Close(False)
    wb.Close(False)

