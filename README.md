Results_Downloader
==================

A pair of scripts I use weekly to download results spreadsheets from two (very similar) websites.

Package_List_Download drives (via Selenium) a Firefox instance to the websites, scrapes a list of possible downloads, then filters the list into three sets: efforts that are 6 months old, efforts that are 12 months old, and DFLN efforts.

Results_Download takes a list of efforts (such as one of the files the above script outputs) as command-line input and downloads the results (xls) from the two sites. It then uses pandas to open these results spreadsheets in Excel and merge them into one xlsx file.
