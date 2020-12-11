# -*- coding: utf-8 -*-
import os
import sys
import io
import cgi
import cgitb
import re
import xlrd
import pandas as pd
import pyodbc
import json
from datetime import date, datetime, timedelta
from decimal import Decimal
import locale
import codecs
import traceback

def eprint(*args, **kwargs):
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        print(*args, file=sys.stderr, **kwargs)

def main(r):
    eprint("main({})".format(r))

    eprint("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    eprint("ok")

    r["dt1"] = "{}-15".format(r["month"])
    r["dt0"] = "{:%Y-%m}-16".format(datetime.strptime(r["month"], "%Y-%m") - timedelta(days=1))
    #r["load"] = load(conn, r)
    sql = """
select
 s.Post
,a.StaffNo
,s.Name
,a.Dt
,a.Shift
,a.BegTm
,a.FinTm
,a.BegTm_i
,a.FinTm_i
,a.StartTm
,a.FinishTm
,a.StartTm_i
,a.FinishTm_i
,a.Late
,a.Early
,a.PTO
,a.Actual
,a.Extra
,a.Night
,a.Actual_i
,a.Extra_i
,a.Night_i
,a.Memo
,c.CalHoliday
from Atnd a
left outer join Staff s
 on (a.StaffNo = s.StaffNo)
inner join Calendar c
 on (a.Dt = c.CalDate)
where a.Dt between '{}' and '{}'
order by
 s.Post
,a.StaffNo
,a.Dt
""".format(r["dt0"], r["dt1"])
    print(sql)
    df = pd.read_sql(sql, conn)
    #locale.setlocale(locale.LC_TIME, 'ja_JP.UTF-8')
    df['StaffNo'] = df['StaffNo'].str.rstrip()
    df['Name'] = df['Name'].str.rstrip()
    df['Memo'] = df['Memo'].str.rstrip()
    locale.setlocale(locale.LC_ALL, '')
    df["Dt"] = pd.to_datetime(df["Dt"])
    df["strDt"] = df['Dt'].dt.strftime('%Y-%m-%d')
    df["fmtDt"] = df['Dt'].dt.strftime('%m/%d(%a)')
    df["strDay"] = df['Dt'].dt.strftime('%a')
    df["Holiday"] = df['CalHoliday'].str.rstrip()
    print(df["BegTm"])
    print(df["BegTm"].astype(str).str[:5])

    df["BegTm5"] = df["BegTm"].astype(str).str[:5]
    df["FinTm5"] = df["FinTm"].astype(str).str[:5]
    df["BegTm5"] = df["BegTm5"].replace("nan","").replace("None","")
    df["FinTm5"] = df["FinTm5"].replace("nan","").replace("None","")

    df["BegTm_i"] = df["BegTm_i"].astype(str).str[:5]
    df["FinTm_i"] = df["FinTm_i"].astype(str).str[:5]
    df["BegTm_i"] = df["BegTm_i"].replace("nan","").replace("None","")
    df["FinTm_i"] = df["FinTm_i"].replace("nan","").replace("None","")
    print(df)
    
    eprint("conn.close()", end=".")
    conn.close()
    eprint("ok")
    r["df"] = df
    return r

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        cgitb.enable()
        form = cgi.FieldStorage()
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        r["month"] = form.getvalue('month', "{:%Y-%m}".format(date.today()))
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        df = r["df"]
        print(df.to_json(orient='table'))
#        print(r["df"].to_json(orient= 'split', force_ascii= True))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("month", help="", nargs="?", default="{:%Y-%m}".format(date.today()), type=str)
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        r = main(vars(parser.parse_args()))
        print(r["df"].to_json(orient= 'split', force_ascii= True))
