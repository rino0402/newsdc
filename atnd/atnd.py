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
from decimal import Decimal
import traceback

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    r["data"] = make(conn, r)
    conn.commit()
    
    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return r

def make(conn, r):
    print("make({})".format(r))
    sql = """
select
 left(convert(d.ts,sql_char),10) dt
,d.ID
,i.TANTO_CODE
,d.Name
,count(*)
,convert(hour(min(d.ts)),sql_char) + ':'
+right('00' + convert(MINUTE(min(d.ts)),sql_char),2)
 minTs
,if(hour(max(d.ts)) < 9,'',
 convert(hour(max(d.ts)),sql_char) + ':'
+right('00' + convert(MINUTE(max(d.ts)),sql_char),2))
 maxTs
,min(d.BT)
,max(d.BT) maxBT
,min(d.ts)
,a.Dt aDt
from DScope d
left outer join DScopeID i
 on (d.ID = i.ID)
left outer join Atnd a
 on (i.TANTO_CODE = a.TANTO_CODE and convert(d.ts,sql_date) = a.Dt)
where dt = '{}'
and d.id<>''
group by
 dt
,d.ID
,i.TANTO_CODE
,d.Name
,a.Dt
order by
 min(d.ts)
""".format(r["dt"] if r["dt"] != "" else date.today())
    sql = """
select
 '{0}' Dt
,s.StaffNo
,s.Name
,s.Post
,s.Shift
,d.Cnt
,d.minTs
,d.maxTs
,a.Dt aDt
,a.BegTm
,a.BegTm_i
,a.FinTm
,a.FinTm_i
,a.StartTm
,a.StartTm_i
,a.FinishTm
,a.FinishTm_i
from Staff s
left outer join (
select
 left(convert(d.ts,sql_char),10) Dt
,d.ID
,i.TANTO_CODE
,count(*) Cnt
,convert(hour(min(d.ts)),sql_char) + ':'
+right('00' + convert(MINUTE(min(d.ts)),sql_char),2)
 minTs
,convert(hour(max(d.ts)),sql_char) + ':'
+right('00' + convert(MINUTE(max(d.ts)),sql_char),2)
 maxTs
from DScope d
left outer join DScopeID i
 on (d.ID = i.ID)
where d.id <> ''
and convert(d.ts, sql_date) = '{0}'
group by
 Dt
,d.ID
,i.TANTO_CODE
) d on (s.StaffNo = d.TANTO_CODE)
left outer join Atnd a
 on (s.StaffNo = a.StaffNo and a.Dt = '{0}')
""".format(r["dt"] if r["dt"] != "" else date.today())
    if r["id"]:
        sql += " where s.StaffNo='{}'".format(r["id"])
    print(sql)
    df = pd.read_sql(sql, conn)
    df['StaffNo'] = df['StaffNo'].str.rstrip()
    print(df)
    data = []
    for i, row in df.iterrows():
        print(i,row)
        d = {}
        d["StaffNo"] = row["StaffNo"]
        d["Shift"] = row["Shift"]
        d["Dt"] = r["dt"]
        d["BegTm"] = row["BegTm"]
        d["BegTm_i"] = row["BegTm_i"]
        if row["minTs"]:
            d["BegTm"] = row["minTs"]
        d["FinTm"] = row["FinTm"]
        d["FinTm_i"] = row["FinTm_i"]
        if row["maxTs"]:
            d["FinTm"] = row["maxTs"]
        d["StartTm"] = row["StartTm"]
        d["StartTm_i"] = row["StartTm_i"]
        d["FinishTm"] = row["FinishTm"]
        d["FinishTm_i"] = row["FinishTm_i"]
        d = calc(d)
        if row["aDt"]:
            d["update"] = update(conn, d)
        else:
            d["insert"] = insert(conn, d)
        print(d)
        data.append(d)
    return data

def get_h_m_s(td):
    m, s = divmod(td.seconds, 60)
    h, m = divmod(m, 60)
    return h, m, s

def calc(d):
    print("calc({})".format(d))
    # 出勤
    beg = None
    try:
        print("BegTm_i={}".format(d["BegTm_i"]))
        beg = d["BegTm_i"]
    except:
        traceback.print_exc()
    if beg == None:
        print("beg={}".format(beg))
        print("BegTm={}".format(d["BegTm"]))
        try:
            beg = datetime.strptime(d["BegTm"], "%H:%M")
        except:
            pass
    if beg == None:
        return d
    # 始業
    if "{:%H:%M}".format(beg) < "07:30":
        beg = time(7,30)
    elif "{:%H:%M}".format(beg) < "08:30":
        beg = time(8,30)
    elif "{:%H:%M}".format(beg) < "09:00":
        beg = time(9,00)
    minute15 = -(-beg.minute // 15) * 15
    if minute15 == 60:
        d["StartTm"] = beg.replace(hour=beg.hour + 1, minute=0)
    else:
        d["StartTm"] = beg.replace(minute=minute15)
    # 始業 入力
    try:
        beg = datetime.strptime(d["StartTm_i"], "%H:%M")
    except:
        pass
    # 10進数 15分単位で切り上げ
    beg = beg.hour + (-(-beg.minute // 15) * 0.25)

    # 退勤
    fin = None
    try:
        fin = d["FinTm_i"]
    except:
        pass
    if fin == None:
        try:
            fin = datetime.strptime(d["FinTm"], "%H:%M")
        except:
            pass
    # 終業
    if fin:
        minute15 = fin.minute // 15 * 15
        d["FinishTm"] = fin.replace(minute=minute15)
    if d["FinishTm_i"]:
        fin = datetime.strptime(d["FinishTm_i"], "%H:%M")
    if fin == None:
        return d
    # 10進数 15分単位で切り捨て
    fin = fin.hour + (fin.minute // 15) * 0.25
    print("{}-{}".format(beg, fin))
    #所定内
    act = fin - beg
    if act < 1:
        fin = 0
        act = 0
        d["FinTm"] = None
        d["FinishTm"] = None
    #昼休み 12:00-12:45
    if beg < 12 and fin > 12.75:
        act -= 0.75
    #休憩 15:00-15:15
    if beg < 15 and fin > 15.25:
        act -= 0.25
    #休憩 19:30-19:45
    if beg < 19.5 and fin > 19.75:
        act -= 0.25
    print(beg, fin, act)
    d["Actual"] = min(act, 7.5)
    #残業
    d["Extra"] = max(0, act - 7.5)
    return d

def update(conn, data):
    sql = "update Atnd"
    st = " set"
    for d in data:
        if d not in ["StaffNo","Dt"]:
            if data[d] == None:
                sql += "{} {} = null".format(st, d)
            else:
                sql += "{} {} = '{}'".format(st, d, data[d])
            st = ", "
    if st != ", ":
        return 0
    sql += " where StaffNo='{}'".format(data["StaffNo"])
    sql += " and Dt='{}'".format(data["Dt"])
    print(sql, end=".")
    try:
        conn.execute(sql)
        ret = conn.execute("select @@rowcount").fetchone()[0]
        print(ret)
        return ret
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise

def insert(conn, data):
    sql = "insert into Atnd ("
    c = ""
    for d in data:
        if data[d]:
            sql += c + d
            c = ","
    sql += ") values ("
    for d in data:
        if data[d]:
            print("{}={} type:{}".format(d, data[d], type(data[d])))
            if type(data[d]) == str:
                sql += "'{0}',".format(data[d].replace("'","''"))
            elif "datetime" in str(type(data[d])):
                sql += "'{0}',".format(data[d])
            else:
                sql += "{0},".format(data[d])
    sql = sql[:-1] + ")"
    print(sql)
    try:
        conn.execute(sql)
        return conn.execute("select @@rowcount").fetchone()[0]
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        cgitb.enable()
        form = cgi.FieldStorage()
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        r["dt"] = form.getvalue('dt', '')
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(r["df"].to_json(orient= 'split', force_ascii= True))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("dt", help="", nargs="?", default="{:%Y-%m-%d}".format(date.today()), type=str)
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--id", help="", default="", type=str)
        r = main(vars(parser.parse_args()))
        print(r["data"])
        for d in r["data"]:
            print("{}".format(d["Dt"]), end=" ")
            print("{}".format(d["StaffNo"]), end=" ")
            print("{}".format(d["BegTm"]), end=" ")
            print("{}".format(d["FinTm"]), end=" ")
            print("{}".format(d["Actual"] if d.get('Actual') else None), end=" ")
            print("{}".format(d["Extra"] if d.get('Extra') else None), end=" ")
            print("{}".format("insert:" + str(d["insert"]) if d.get('insert') else ""), end="")
            print("{}".format("update:" + str(d["update"]) if d.get('update') else ""), end="")
            print("")




            
