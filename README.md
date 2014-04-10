Results_Downloader
==================
```
Download and combine results spreadsheets from a number of (very similar) websites.

Usage: results_download.py [--pkgset=<pkgset>|--pkglist=<filename>] [-k | --keeplists] [-l | --keepdl] [--quiet]
       results_download.py (-h | --help)

Options:
  -h --help             Show this screen.
  -k --keeplists        Don't delete scraped lists.
  -l --keepdl           Don't delete original downloads after processing.  
  --pkgset=<pkgset>     Download <set> of results, e.g. "6m", "12m", "DFLN"
  --pkglist=<filename>  Use a particular list of packages instead of a 
                        standard results set.
  --quiet               print less text
```

(As you don't know these websites (I keep all info about them in an untracked config.json), this surely has no practical use for you beyond seeing what kind of code I write.)

* scrape_pkgs drives (via [Selenium](http://docs.seleniumhq.org/)) a [Firefox](http://www.mozilla.org/en-US/firefox/) instance to the websites and scrapes a list of packages from each, which then get merged together.
* create_pkg_tables creates tables (i.e. CSV and TXT files) from this list.
* download_from_site takes a list of packages (derived from the previous tables) to get results for, and downloads the appropriate results (as XLS).
* merge_raw_reports merges these raw results into one pandas DataFrame


TODO:
-----
* Tests
* Separate GtoG results out to separate file
* Give warnings when list costs, production costs, or list descriptions aren't found
