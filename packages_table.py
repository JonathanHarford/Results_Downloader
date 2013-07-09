import csv
import codecs
import cStringIO
import re
import datetime

EFFTYPES = {"A":"Prsp",
            "B":"Rnwl",
            "D":"PI",
            "F":"FSI",
            "G":"Rnwl",
            "H":"Rnwl",
            "L":"NI",
            "N":"NI",
            "P":"Prsp",
            "R":"Rnwl",
            "S":"Rnwl"}

class PackagesTable(list):

    class UnicodeWriter:
        """
        A CSV writer which will write rows to CSV file "f",
        which is encoded in the given encoding.
        """

        def __init__(self, f, dialect=csv.excel, encoding="utf-8", **kwds):
            # Redirect output to a queue
            self.queue = cStringIO.StringIO()
            self.writer = csv.writer(self.queue, dialect=dialect, **kwds)
            self.stream = f
            self.encoder = codecs.getincrementalencoder(encoding)()

        def writerow(self, row):
            self.writer.writerow([unicode(s).encode("utf-8") for s in row])
            # Fetch UTF-8 output from the queue ...
            data = self.queue.getvalue()
            data = data.decode("utf-8")
            # ... and reencode it into the target encoding
            data = self.encoder.encode(data)
            # write to the target stream
            self.stream.write(data)
            # empty queue
            self.queue.truncate(0)

        def writerows(self, rows):
            for row in rows:
                self.writerow(row)

    def __init__(self, *args):
        self.fieldnames = list(args)

    def __str__(self):
        s = []
        for rec in self:
            for i in range(len(self.fieldnames)):
                s.append(self.fieldnames[i] + ": " + rec[i])
            s.append("\n")
        return "\n".join(s)

    def addrecord(self, record):
        self.append(list(record))

    def addfield(self, pos, fieldname, defaultval=""):
        self.fieldnames.insert(pos,fieldname)
        for rec in self:
            rec.insert(pos, defaultval)

    def filter_efforts(self):

        # Filter out packages that aren't "our" efforts
        recnum = 0
        while recnum < len(self):
            if re.match(r"[ABDFGHLNPRS]\d\d\d\s", self[recnum][0]):
                recnum = recnum + 1
            else:
                del self[recnum]
                continue


        self.addfield(2,"Eff")
        self.addfield(6,"EffType")
        self.addfield(7,"Age")
        self.addfield(8,"Results Group")
        age_6m  = set()
        age_12m = set()

        for rec in self:
            pkg, org, eff, packcode, maildate, ffdate, efftype, age, resgrp = rec

            # Truncate Eff from Package
            eff = pkg[0:4]
            rec[2] = eff

            # Calculate EffType of package
            efftype = EFFTYPES[eff[0]]
            rec[6] = efftype

            # Calculate age of package
            age = (ffdate-maildate).days
            rec[7] = age

            if age <= 180:
                age_6m.add(eff)
            elif 365 > age > 180:
                age_12m.add(eff)

        for rec in self:
            pkg, org, eff, packcode, maildate, ffdate, efftype, age, resgrp = rec

            if   eff in age_6m  and age < 365:
                resgrp = "6m"
            elif eff in age_12m and age < 600:
                resgrp = "12m"

            rec[8] = resgrp






    def writetocsv(self, filename):
        writer = self.UnicodeWriter(open(filename, 'wb'))
        writer.writerow(self.fieldnames)
        writer.writerows(self)
