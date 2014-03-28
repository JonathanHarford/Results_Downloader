import csv
import re

EFFTYPES = {"A":"Prsp",
            "B":"Rnwl",
            "D":"PIns",
            "F":"FSI",
            "G":"Rnwl",
            "H":"Rnwl",
            "L":"NIns",
            "N":"NIns",
            "P":"Prsp",
            "R":"Rnwl",
            "S":"Rnwl"}

class PackagesTable(list):

    def __init__(self, *args):
        self.fieldnames = list(args)

#     def __str__(self):
#         s = []
#         for rec in self:
#             for i in range(len(self.fieldnames)):
#                 s.append(self.fieldnames[i] + ": " + rec[i])
#             s.append("\n")
#         return "\n".join(s)

    def addrecord(self, record):
        self.append(list(record))

#     def addfield(self, pos, fieldname, defaultval=""):
#         self.fieldnames.insert(pos,fieldname)
#         for rec in self:
#             rec.insert(pos, defaultval)

    def filter_efforts(self):

        # Filter out packages that aren't [ABDFGHLNPRS]\d\d\d\s efforts
        recnum = 0
        while recnum < len(self):
            if re.match(r"[ABDFGHLNPRS]\d\d\d\s", self[recnum][0]): 
                recnum = recnum + 1
            else:
                del self[recnum]
                continue


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

        # Classify _efforts_ by age (not merely packages, which can vary).
        for rec in self:
            eff, age = rec[2], rec[7]
            if   eff in age_6m  and age < 365:
                rec[8] = "6m"
            elif eff in age_12m and age < 600:
                rec[8] = "12m"

    def writetocsv(self, filename):
        with open(filename, 'w') as f:
            writer = csv.writer(f, dialect=csv.excel, lineterminator='\n')
            print(self.fieldnames)
            writer.writerow(self.fieldnames)
            writer.writerows(self)
