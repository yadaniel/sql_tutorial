#!/cygdrive/c/Python39/python

import re, os, sys, sqlite3
import pandas as pd

infile = None
if len(sys.argv) == 1:
    print("usage: sql_from_xlsx.py <infile>.xlsx")
    sys.exit(1)
elif len(sys.argv) == 2:
    infile = sys.argv[1]
    if infile.upper().endswith(".XLSX"):
        if os.path.isfile(infile):
            pass
        else:
            print(f"{infile} does not exist")
            sys.exit(2)
    else:
        print("usage: sql_from_xlsx.py <infile>.xlsx")
        sys.exit(1)


m = re.match(r"(?P<table>.*)[.]XLSX", os.path.basename(infile).upper())
tablename = None
if m:
    tablename = m.group("table")
else:
    print("match error")
    sys.exit(3)

# print(infile, tablename)
# sys.exit(4)

con = sqlite3.connect('out.db')
df = pd.read_excel(infile, header = 0, usecols = ["SAP", "DESC"])
df = df.fillna("")
df.SAP = df.SAP.str.strip()
df.DESC = df.DESC.str.strip()
df = df[df.DESC != ""]

# print(df)
# print(type(df))
# print(type(df.SAP))
# print(type(df.DESC))
# print(len(df))
# sys.exit()

# for sheet in df:
#     # wb[sheet].to_sql(sheet, con = con, index=False)
#     wb[sheet].to_sql(tablename, con = con, index=False)
# con.commit()
# con.close()

cxn = sqlite3.connect('out.db')
df.to_sql(name = tablename, con = cxn, if_exists = 'replace', index = True)
cxn.commit()
cxn.close()

