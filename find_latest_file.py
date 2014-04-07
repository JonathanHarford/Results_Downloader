#! python3

import os
import re

def datepull(s):
    match = re.search('(\d\d\d\d)-(\d\d)-(\d\d)', s)
    return int(match.group(1) + match.group(2) + match.group(3)) if match else False

def find_latest_file(fnfrag, path='.'):
    fn_tups = [(datepull(fn), fn) for fn in os.listdir(path) if fnfrag in fn and datepull(fn)]
    return sorted(fn_tups).pop()[1]