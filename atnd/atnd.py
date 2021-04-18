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
    if r["dscopeid"] == "1":
        dscope_sql = """
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
""".format(r["dt"] if r["dt"] != "" else date.today())
    else:
        dscope_sql = """
select
 left(convert(d.ts,sql_char),10) Dt
,d.ID
,d.ID TANTO_CODE
,count(*) Cnt
,convert(hour(min(d.ts)),sql_char) + ':'
+right('00' + convert(MINUTE(min(d.ts)),sql_char),2)
 minTs
,convert(hour(max(d.ts)),sql_char) + ':'
+right('00' + convert(MINUTE(max(d.ts)),sql_char),2)
 maxTs
from DScope d
where d.id <> ''
and convert(d.ts, sql_date) = '{0}'
group by
 Dt
,d.ID
,TANTO_CODE
""".format(r["dt"] if r["dt"] != "" else date.today())

    sql = """
select
 j.JCode
,'{0}' Dt
,s.StaffNo
,s.Name
,s.Post
,ifnull(a.Shift,s.Shift) Shift
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
from JGyobu j, Staff s
left outer join (
{1}
) d on (s.StaffNo = d.TANTO_CODE)
left outer join Atnd a
 on (s.StaffNo = a.StaffNo and a.Dt = '{0}')
where j.JGYOBU='0'
""".format(r["dt"] if r["dt"] != "" else date.today(), dscope_sql)
    if r["id"]:
        sql += " and s.StaffNo='{}'".format(r["id"])
    if r["quit"] == "0":
        sql += " and (s.QuitDt >= '{}' or s.QuitDt is null)".format(r["dt"] if r["dt"] != "" else date.today())
    elif r["quit"] == "1":
        sql += " and s.Quit <> ''"
    print(sql)
    df = pd.read_sql(sql, conn)
    df['JCode'] = df['JCode'].str.rstrip()
    df['StaffNo'] = df['StaffNo'].str.rstrip()
    print(df)
    data = []
    for i, row in df.iterrows():
        print(i,row)
        d = {}
        d["JCode"] = row["JCode"]
        d["StaffNo"] = row["StaffNo"]
        d["Shift"] = row["Shift"]
        d["Dt"] = r["dt"]
        d["BegTm"] = row["BegTm"]
        d["BegTm_i"] = row["BegTm_i"]
        if row["minTs"]:
            d["BegTm"] = row["minTs"]
        d["FinTm"] = row["FinTm"]
        d["FinTm_i"] = row["FinTm_i"]
        if row["maxTs"] and row["minTs"]:
            if row["maxTs"] != row["minTs"]:
                d["FinTm"] = row["maxTs"]
            else:
                d["FinTm"] = None
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

    # 出勤時刻
    beg = d.get("BegTm", None)
    if beg:
        if re.search('^\d+:\d+$', str(beg)):
            beg = datetime.strptime(str(beg), "%H:%M").time()
        elif re.search('^\d+:\d+:\d+$', str(beg)):
            beg = datetime.strptime(str(beg), "%H:%M:%S").time()
    print("calc():beg={}".format(beg))
    if beg and d["Shift"] != "--":
        # 始業時刻 StartTm
        if d["JCode"] == "OSAKA":
            if d["Shift"] == "09":
                # 9:30 - 16:00 5.50
                if "{:%H:%M}".format(beg) <= "09:30":
                    beg = time(9,30)
            elif d["Shift"] == "06":
                if "{:%H:%M}".format(beg) <= "07:30":
                    beg = time(7,30)
                elif "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            elif d["Shift"] == "90":
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            else:
                if "{:%H:%M}".format(beg) <= "08:00":
                    beg = time(8,00)
                elif "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
                elif "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
        elif d["JCode"] == "CARP":
            if "{:%H:%M}".format(beg) <= "08:00":
                beg = time(8,00)
            elif "{:%H:%M}".format(beg) <= "09:00":
                beg = time(9,00)
        elif d["JCode"] == "NARA":
            if d["Shift"] == "1A":
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            elif d["Shift"] == "2B":
                if "{:%H:%M}".format(beg) <= "10:00":
                    beg = time(10,00)
            elif d["Shift"] == "3C":
                if "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
            elif d["Shift"] == "4D":
                if "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
            elif d["Shift"] == "5E":
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            elif d["Shift"] == "6G":
                if "{:%H:%M}".format(beg) <= "09:45":
                    beg = time(9,45)
            elif d["Shift"] == "7L":
                if "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
            elif d["Shift"] == "9M":
                if "{:%H:%M}".format(beg) <= "15:00":
                    beg = time(15,00)
            else:
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
                elif "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
        else:
            if "{:%H:%M}".format(beg) <= "09:00":
                beg = time(9,00)
        minute15 = -(-beg.minute // 15) * 15
        if minute15 == 60:
            d["StartTm"] = beg.replace(hour=beg.hour + 1, minute=0)
        else:
            d["StartTm"] = beg.replace(minute=minute15)
    else:
        d["StartTm"] = None
        beg = None
    print("calc():StartTm={}".format(d["StartTm"]))
    if d.get("StartTm_i", None):
        print("calc():StartTm_i={}".format(d["StartTm_i"]))
        if re.search('^\d+:\d+$', str(d["StartTm_i"])):
            beg = datetime.strptime(str(d["StartTm_i"]), "%H:%M").time()
        elif re.search('^\d+:\d+:\d+$', str(d["StartTm_i"])):
            beg = datetime.strptime(str(d["StartTm_i"]), "%H:%M:%S").time()
    # 10進数 15分単位で切り上げ
    if beg:
        beg = beg.hour + (-(-beg.minute // 15) * 0.25)
    print("calc():beg={}".format(beg))

    # 退勤時刻
    fin = d.get("FinTm", None)
    if fin:
        if re.search('^\d+:\d+$', str(fin)):
            fin = datetime.strptime(str(fin), "%H:%M").time()
        elif re.search('^\d+:\d+:\d+$', str(fin)):
            fin = datetime.strptime(str(fin), "%H:%M:%S").time()
    print("calc():fin={}".format(fin))
    if fin and d["Shift"] != "--":
        # 終業時刻 FinishTm
        if d["JCode"] == "OSAKA":
            if d["Shift"] == "90":
                if "{:%H:%M}".format(fin) >= "17:00":
                    fin = time(17,0)
        minute15 = fin.minute // 15 * 15
        d["FinishTm"] = fin.replace(minute=minute15)
        if d["StartTm"] and d["FinishTm"] < d["StartTm"]:
            d["FinishTm"] = None
    else:
        d["FinishTm"] = None
        fin = None
    print("calc():FinishTm={}".format(d["FinishTm"]))
    if d.get("FinishTm_i", None):
        print("calc():FinishTm_i={}".format(d["FinishTm_i"]))
        if re.search('^\d+:\d+$', str(d["FinishTm_i"])):
            fin = datetime.strptime(str(d["FinishTm_i"]), "%H:%M").time()
        elif re.search('^\d+:\d+:\d+$', str(d["FinishTm_i"])):
            fin = datetime.strptime(str(d["FinishTm_i"]), "%H:%M:%S").time()
    # 10進数 15分単位で切り捨て
    if fin:
        fin = fin.hour + (fin.minute // 15) * 0.25
    print("calc():fin={}".format(fin))
    #所定内
    act = 0
    if fin and beg:
        act = fin - beg
        #昼休み 12:00-12:45
        if beg < 12 and fin > 12.75 and d["Shift"] not in ["6G","7L"]:
            act -= 0.75
        #休憩 15:00-15:15
        if beg < 15 and fin > 15.25:
            act -= 0.25
        #休憩 19:30-19:45
        if beg < 19.5 and fin > 19.75:
            act -= 0.25
        act = max(0, act)
    print("calc():act={}".format(act))
    d["Actual"] = min(act, 7.5)
    #残業
    d["Extra"] = max(0, act - 7.5)
    #休出
    if d["Shift"] == "00":
        #d["Extra"] += d["Actual"]
        #d["Actual"] = 0
        d["Dayoff"] = d["Actual"] + d["Extra"]
        d["Actual"] = 0
        d["Extra"] = 0
    else:
        d["Dayoff"] = 0
    return d

def calc_old(d):
    print("calc({})".format(d))
    d["Actual"] = 0
    d["Extra"] = 0
    # 出勤
    beg = None
    try:
        print("BegTm_i={}".format(d["BegTm_i"]))
        beg = d["BegTm_i"]
    except:
        traceback.print_exc()
    try:
        print("StartTm_i={}".format(d["StartTm_i"]))
        if re.search('^\d+:\d+$', str(d["StartTm_i"])):
            beg = datetime.strptime(str(d["StartTm_i"]), "%H:%M")
        elif re.search('^\d+:\d+:\d+$', str(d["StartTm_i"])):
            beg = datetime.strptime(str(d["StartTm_i"]), "%H:%M:%S")
    except:
        traceback.print_exc()
    if beg == None:
        print("beg={}".format(beg))
        print("BegTm={}".format(d["BegTm"]))
        if d["BegTm"] == None:
            pass
        elif re.search('^\d+:\d+$', str(d["BegTm"])):
            beg = datetime.strptime(str(d["BegTm"]), "%H:%M")
        elif re.search('^\d+:\d+:\d+$', str(d["BegTm"])):
            beg = datetime.strptime(str(d["BegTm"]), "%H:%M:%S")
        print("beg={}".format(beg))
        if beg == None:
            return d
        # 始業
        if d["JCode"] == "OSAKA":
            if d["Shift"] == "09":
                # 9:30 - 16:00 5.50
                if "{:%H:%M}".format(beg) <= "09:30":
                    beg = time(9,30)
            elif d["Shift"] == "06":
                if "{:%H:%M}".format(beg) <= "07:30":
                    beg = time(7,30)
                elif "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            elif d["Shift"] == "90":
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            else:
                if "{:%H:%M}".format(beg) <= "08:00":
                    beg = time(8,00)
                elif "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
                elif "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
        elif d["JCode"] == "CARP":
            if "{:%H:%M}".format(beg) <= "08:00":
                beg = time(8,00)
            elif "{:%H:%M}".format(beg) <= "09:00":
                beg = time(9,00)
        elif d["JCode"] == "NARA":
            if d["Shift"] == "1A":
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            elif d["Shift"] == "2B":
                if "{:%H:%M}".format(beg) <= "10:00":
                    beg = time(10,00)
            elif d["Shift"] == "3C":
                if "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
            elif d["Shift"] == "4D":
                if "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
            elif d["Shift"] == "5E":
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
            elif d["Shift"] == "6G":
                if "{:%H:%M}".format(beg) <= "09:45":
                    beg = time(9,45)
            elif d["Shift"] == "7L":
                if "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
            elif d["Shift"] == "9M":
                if "{:%H:%M}".format(beg) <= "15:00":
                    beg = time(15,00)
            else:
                if "{:%H:%M}".format(beg) <= "08:30":
                    beg = time(8,30)
                elif "{:%H:%M}".format(beg) <= "09:00":
                    beg = time(9,00)
        else:
            if "{:%H:%M}".format(beg) <= "09:00":
                beg = time(9,00)

        minute15 = -(-beg.minute // 15) * 15
        if minute15 == 60:
            d["StartTm"] = beg.replace(hour=beg.hour + 1, minute=0)
        else:
            d["StartTm"] = beg.replace(minute=minute15)
    # 10進数 15分単位で切り上げ
    beg = beg.hour + (-(-beg.minute // 15) * 0.25)

    # 退勤
    fin = None
    try:
        fin = d["FinTm_i"]
    except:
        pass
    if d["FinishTm_i"]:
        print(str(d["FinishTm_i"]))
        if re.search('^\d+:\d+$', str(d["FinishTm_i"])):
            fin = datetime.strptime(str(d["FinishTm_i"]), "%H:%M")
        elif re.search('^\d+:\d+:\d+$', str(d["FinishTm_i"])):
            fin = datetime.strptime(str(d["FinishTm_i"]), "%H:%M:%S")
    if fin == None:
        if d["FinTm"] == None:
            pass
        elif re.search('^\d+:\d+$', str(d["FinTm"])):
            fin = datetime.strptime(str(d["FinTm"]), "%H:%M")
        elif re.search('^\d+:\d+:\d+$', str(d["FinTm"])):
            fin = datetime.strptime(str(d["FinTm"]), "%H:%M:%S")
        """
        try:
            fin = datetime.strptime(d["FinTm"], "%H:%M")
        except:
            pass
        """
        if d["JCode"] == "OSAKA":
            if d["Shift"] == "90":
                if "{:%H:%M}".format(fin) >= "17:00":
                    fin = time(17,0)
    
        # 終業
        if fin:
            minute15 = fin.minute // 15 * 15
            d["FinishTm"] = fin.replace(minute=minute15)
    print("fin={}".format(fin))
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
    if beg < 12 and fin > 12.75 and d["Shift"] not in ["6G","7L"]:
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
    if d["Shift"] == "00":
        d["Extra"] += d["Actual"]
        d["Actual"] = 0
    return d

def update(conn, data):
    sql = "update Atnd"
    st = " set"
    for d in data:
        if d not in ["JCode","StaffNo","Dt"]:
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
        if data[d] and d not in ["JCode"]:
            sql += c + d
            c = ","
    sql += ") values ("
    for d in data:
        if data[d] and d not in ["JCode"]:
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
        parser.add_argument("--dscopeid", help="0/1", default="1", type=str)
        parser.add_argument("--quit", help="0:退職者除く 1:退職者のみ", default="0", type=str)
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
     
