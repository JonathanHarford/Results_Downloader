Results_Downloader
==================

Scripts I use weekly to download and combine results spreadsheets from two (very similar) websites.

results_download.py first runs scrape_effort_list.py to drive (via Selenium) a Firefox instance to the websites, scrape a list of possible downloads, then filter the list into three sets: efforts that are 6 months old, efforts that are 12 months old, and DFLN efforts.

If results_download.py was given an argument of "6m", "12m", or "DFNL", it inputs the relevant just-scraped list into download_from_site.py.

Via selenium again, download_from_site.py downloads the results (xls) from the two sites. It then uses pandas to open these results spreadsheets in Excel and merge them into one xlsx file.

TODO:
-----

* Downloads from both sites shouldn't be necessary
* Separate GtoG results out to separate file? 
* Logging
* Join CPP tables
* Use SQL instead of pandas?