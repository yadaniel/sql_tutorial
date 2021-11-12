#!/cygdrive/c/Python39/python

import re, os, sys, sqlite3
import pandas as pd

def usage(exitcode):
    print("usage: sql_to_xlsx.py <infile>.db <table> <incolnum1>:<outcolname1>[<incolnumN>:<outcolnameN>] <outfile>")
    txt = """
    sql_to_xlsx.py sqldata.db {1:MAN,2:MPN,3:OUR_MPN, 4:STATE} out.db
    sql_to_xlsx.py sqldata.db {1:MAN,2:MPN,3:OUR_MPN,4:STATE} out.db
    sql_to_xlsx.py sqldata.db 1:MAN 2:MPN 3:OUR_MPN 4:STATE out.db
        incolnum selects the column number in the sqldata table, the order may be different, all data is text
        outcolname assignes the column name in the xlsx sheet, all names will be uppercase
        xlsx filename equals infile without db extension
    """
    print("usage examples:", txt)
    print("hints:")
    print("    sqlite3 data.db '.schema'")
    print("    sqlite3 data.db '.schema <table>'")
    sys.exit(exitcode)

pattern = re.compile(r'(?P<infile>[_a-z./][_a-z0-9./]*)\s+(?P<table>[_a-z][_a-z0-9]*)(?P<col>(\s+\d+:\w+)+)?', re.I)
# print(' '.join(sys.argv[1:]))
m = pattern.match(' '.join(sys.argv[1:]))
if not m:
    usage(1)

infile = m.group("infile")
table = m.group("table")
outfile = table
columns = []
colnums = []
colnames = []
if m.group("col"):
    for col in m.group("col").split():
        colnum, colname = col.split(":")
        colnum = int(colnum) - 1            # dataframe counts from 0, xlsx from 1
        colnums.append(colnum)
        colnames.append(colname.upper())
        columns.append((colnum, colname.upper()))

if not infile.upper().endswith(".DB"):
    print(f"{infile} not DB file")
    usage(2)

if not os.path.isfile(infile):
    print(f"{infile} does not exist")
    usage(3)

if len(set(colnums)) < len(colnums):
    print("colnums not unique")
    usage(4)

if len(set(colnames)) < len(colnames):
    print("colnames not unique")
    usage(5)

# # DEBUG
print(f"infile = {infile}")
print(f"table = outfile = {outfile}")
print(f"columns = {columns}")
# sys.exit()

con = sqlite3.connect(infile)
df = pd.read_sql_query(f"SELECT * from {table}", con)
df = df.iloc[:,colnums]
df.columns = colnames
# print(df.head())
df.to_excel(f"{outfile}.xlsx", index = False, sheet_name = table)
con.close()

