# Check zip contents
import zipfile
import json
import sys

def list_zip(path):
    with zipfile.ZipFile(path, 'r') as z:
        return z.namelist()

eol_files = list_zip("xbrl_sample/E00872_20250625_S100W4FO.zip")
edinet_files = list_zip("xbrl_sample/Xbrl_Search_20260309_181654.zip")

print("EOL files:", len(eol_files))
print("EDINET files:", len(edinet_files))

for f in eol_files[:10]: print("EOL:", f)
for f in edinet_files[:10]: print("EDI:", f)
