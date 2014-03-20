import sys
import os
from time import strftime

import scrape_effort_list
from config import US_NAV, NY_NAV
from append_xls import append_xls
from download_from_site import download_from_site


if __name__ == "__main__":

    scrape_effort_list.main() # Download effort lists
    
    # If no arguments, we're done once we've downloaded and processed the lists!
    try:
        effort_cat = sys.argv[1]
    except IndexError:
        sys.exit() 
    
    # Figure out which effort list matches the type of download we want to do (6m, 12m, or DFLN)
    for filename in os.listdir('.'):
        if filename.startswith(effort_cat):
            break
        else:
            continue
    
    # Create list of which efforts we want results for
    packages_to_download = [line.strip().split("\t") for line in open(filename, 'rb').readlines()]

    us_pkgs_download = [pkg for org, pkg in packages_to_download if org==US_NAV['org']]
    ny_pkgs_download = [pkg for org, pkg in packages_to_download if org==NY_NAV['org']]

    # Download the results for US and NY.
    # We could do the downloads simultaneously (threads), but that causes Firefox troubles: two browsers using the same profile.

    # Download results to files
    download_from_site(US_NAV, us_pkgs_download)
    download_from_site(NY_NAV, ny_pkgs_download)

    # Rename results files
    now_str = strftime('''%Y%m%d-%H%M%S''')
    us_filename = 'US Results ' + now_str + '.xls'
    ny_filename = 'NY Results ' + now_str + '.xls'
    os.rename(US_NAV['filename'], us_filename)
    os.rename(NY_NAV['filename'], ny_filename)

    # Merge the two files
    append_xls(us_filename, ny_filename, filename)


