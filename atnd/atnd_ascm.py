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
from datetime import date, datetime, timedelta, time
from dateutil.relativedelta import relativedelta
from decimal import Decimal
import traceback
import zenhan
import atnd
import staff

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    if r["staff"]:
        r["make"] = make_staff(conn, r)
    else:
        r["make"] = make(conn, r)
    conn.commit()

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return r

def make_staff(conn, r):
    print("make_staff({})".format(r))
    sql = """
select
distinct
 a.StaffNo
,a.Name
,'NP' Post
,rtrim(a.Shift) + rtrim(a.ShiftNm) Shift0
,count(*) cnt
,s.StaffNo sStaffNo
from ascm a
left outer join staff s on (a.StaffNo = s.StaffNo)
group by
 a.StaffNo
,a.Name
,Post
,Shift0
,sStaffNo
order by
 a.StaffNo
,cnt desc
"""
    df = pd.read_sql(sql, conn)
    df['StaffNo'] = df['StaffNo'].str.strip()
    df['Name'] = df['Name'].str.strip()
    print(df)
    for i, row in df.iterrows():
        print(i,row)
        if row.sStaffNo:
            continue
        if i > 0 and row.StaffNo == df.iloc[i-1].StaffNo:
            continue
        d = {}
        d["StaffNo"] = row.StaffNo
        d["Name"] = row.Name
        d["Post"] = row.Post
        d["Shift"] = zenhan.z2h(row.Shift0)
        print(d)
        d["insert"] = staff.insert(conn, d)
    return df

def make(conn, r):
    print("make({})".format(r))
    sql = """
select {}
 j.JCode
,t.*
,a.Dt aDt
from JGyobu j, ascm t
left outer join Atnd a
 on (t.StaffNo = a.StaffNo and t.Dt = a.Dt)
where j.JGYOBU='0'
""".format("top {}".format(r["limit"]) if r["limit"] > 0 else "")
    if r["dt"]:
        if r["month"]:
            dt0 = r["dt"]
            dt0 = datetime.strptime(r["dt"], "%Y-%m-%d")
            dt0 -= relativedelta(months=1)
            sql += " and t.dt between '{:%Y-%m-16}' and '{}'".format(dt0, r["dt"])
        else:
            sql += " and t.dt = '{}'".format(r["dt"])
    print(sql)
    df = pd.read_sql(sql, conn)
    print(df)
    for i, row in df.iterrows():
        print(i,row)
        d = {}
        d["JCode"] = row.JCode.rstrip()
        d["StaffNo"] = row.StaffNo.rstrip()
        d["Shift"] = zenhan.z2h(row.Shift.rstrip() + row.ShiftNm.rstrip())
        if d["Shift"] == "9":
            d["Shift"] = "9M"
        d["Dt"] = row.Dt
        d["BegTm"] = row.BegTm
        d["BegTm_i"] = None
        d["FinTm"] = row.FinTM
        d["FinTm_i"] = None
        #d["FinishTm_i"] = None
        d["Late"] = row.Late
        d["Early"] = row.Early
        d["Memo"] = row.Kubun.rstrip() + row.Memo.rstrip()
        print(d)
        d = atnd.calc(d)
        if d["Memo"] == "有休":
            d["Actual_i"] = 0
            d["PTO"] = d["Actual"]
        if row["aDt"]:
            d["update"] = atnd.update(conn, d)
        else:
            d["insert"] = atnd.insert(conn, d)

    return df

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        pass
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("dt", help="", nargs="?", default="{:%Y-%m-%d}".format(date.today()), type=str)
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--limit", help="", default=0, type=int)
        parser.add_argument("--staff", action="store_true", default=False)
        parser.add_argument("--month", action="store_true", default=False)
        r = main(vars(parser.parse_args()))
        print(r["make"])
