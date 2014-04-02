def create_pkg_tables(t, ff_strs):
    
    full_ff_str = ' '.join([org + ' FF ' + ff_str for org, ff_str in ff_strs])
        
    # Let's make a list of all filenames as we create them. 
    filenames = []

    # Save a table of every effort, even the ones we don't care about.
    filename = "All Packages " + full_ff_str + ".csv"

    filenames.append(filename)
    t.writetocsv(filename)

    t.filter_efforts()

    # Save a lovely table of all packages we might deal with.
    filename = "All Relevant Efforts "  + full_ff_str + ".csv"

    filenames.append(filename)
    t.writetocsv(filename)

    # Save a list of all packages 6 months old and younger.
    filename = "6m " + full_ff_str + ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "6m":
                f.write(rec[1] + "\t" + rec[0] + "\n")

    # Save a list of all packages between 6 and 12 months old.
    filename = "12m " + full_ff_str + ".txt"

    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[8] == "12m":
                f.write(rec[1] + "\t" + rec[0] + "\n")

    # Save a list of all DFLN packages.
    filename = "DFLN " + full_ff_str + ".txt"
    
    filenames.append(filename)
    with open(filename, "w") as f:
        for rec in t:
            if rec[6] in ["PIns", "FSI", "NIns"] and (rec[4].year >= 2011) :
                f.write(rec[1] + "\t" + rec[0] + "\n")

    return filenames

if __name__ == "__main__":
    import scrape_pkgs
    create_pkg_tables(scrape_pkgs())