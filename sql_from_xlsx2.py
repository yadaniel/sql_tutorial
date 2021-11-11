#!/cygdrive/c/Python39/python

import re, os, sys, sqlite3
import pandas as pd

def usage(exitcode):
    print("usage: sql_from_xlsx.py <infile>.xlsx <incolnum1>:<outcolname1>[<incolnumN>:<outcolnameN>] <outfile>")
    txt = """
    sql_from_xlsx.py bom.xlsx {1:MATCHED_MAN,2:MATCHED_MPN,3:OUR_MPN, 4:OUR_MAN, 10:STATE,5:ROHS} out.db
    sql_from_xlsx.py bom.xlsx {1:MATCHED_MAN,2:MATCHED_MPN,3:OUR_MPN,4:OUR_MAN, 10:STATE,5:ROHS} out.db
    sql_from_xlsx.py bom.xlsx 1:MATCHED_MAN 2:MATCHED_MPN 3:OUR_MPN 4:OUR_MAN 10:STATE 5:ROHS out.db
        incolnum selects the column number in the input xlsx file, the order may be different, all data is text
        outcolname assignes the column name in the database table, all names will be uppercase
        database table name equals infile without xlsx extension
    """
    print("usage examples:", txt)
    sys.exit(exitcode)

# pattern = re.compile(r'(?P<infile>[_a-z0-9./]+[.]xlsx)\s+(?P<col>(\d+:\w+\s+)+)?(?P<outfile>.+)', re.I)
pattern = re.compile(r'(?P<infile>[_a-z0-9./]+)\s+(?P<col>(\d+:\w+\s+)+)?(?P<outfile>.+)', re.I)
# print(' '.join(sys.argv[1:]))
m = pattern.match(' '.join(sys.argv[1:]))
if not m:
    usage(1)

infile = m.group("infile")
columns = []
colnums = []
colnames = []
if m.group("col"):
    for col in m.group("col").split():
        colnum, colname = col.split(":")
        colnums.append(int(colnum))
        colnames.append(colname.upper())
        columns.append((int(colnum), colname.upper()))
outfile = None
if m.group("outfile"):
    outfile = m.group("outfile")

if not infile.upper().endswith(".XLSX"):
    print(f"{infile} not XLSX file")
    usage(2)

if not os.path.isfile(infile):
    print(f"{infile} does not exist")
    usage(3)

if m := re.match(r"(?P<tablename>.*)[.]XLSX", os.path.basename(infile).upper()):
    tablename = m.group("tablename")
else:
    print("match error")
    usage(4)

if len(set(colnums)) < len(colnums):
    print("colnums not unique")
    usage(5)

if len(set(colnames)) < len(colnames):
    print("colnames not unique")
    usage(6)

# # DEBUG
print(f"infile = {infile}")
print(f"outfile = {outfile}")
print(f"columns = {columns}")
print(f"tablename = {tablename}")
# sys.exit()

con = sqlite3.connect(outfile)
# df = pd.read_excel(infile, header = 0, usecols = ["SAP", "DESC"])
df = pd.read_excel(infile, header = 0, usecols = None if columns == [] else colnums, dtype = str)
df = df.fillna("")
df.columns = [n for _,n in sorted(columns)]
df = df[colnames]
for colname in colnames:
    df[colname] = df[colname].str.strip()

# df.SAP = df.SAP.str.strip()
# df.DESC = df.DESC.str.strip()
# df = df[df.DESC != ""]

# print(df)
# print(type(df))
# print(type(df.SAP))
# print(type(df.DESC))
# print(len(df))
# sys.exit()

cxn = sqlite3.connect(outfile)
df.to_sql(name = tablename, con = cxn, if_exists = 'replace', index = True)
cxn.commit()
cxn.close()

