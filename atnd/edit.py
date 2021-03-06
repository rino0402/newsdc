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
import locale
import codecs
import traceback
import sdc

def main(r):
    print("pyodbc.connect({0})".format(r["dsn"]), end=";")
    conn = pyodbc.connect('DSN=' + r["dsn"])
    print("ok")
    r["StaffNo"] = r["id"].split("_")[0]
    r["Dt"] = r["id"].split("_")[1]
    sql = "update Atnd set "
    calc = False
    for k in r.keys():
        if k in ["Shift"]:
            calc = True
            sql += "{}='{}'".format(k, r[k])
        elif k in ["Shift","Memo"]:
            sql += "{}='{}'".format(k, r[k])
        elif k in ["StartTm_i","FinishTm_i"]:
            calc = True
            if r[k]:
                sql += "{} = '{}'".format(k, r[k])
            else:
                sql += "{} = Null".format(k)
        elif k in ["Actual_i","Extra_i","Night_i","Dayoff_i"]:
            if r[k]:
                sql += "{} = {}".format(k, r[k])
            else:
                sql += "{} = Null".format(k)
        elif k in ["Late","Early","PTO","PTO_tm"]:
            if r[k]:
                sql += "{} = {}".format(k, r[k])
            else:
                sql += "{} = 0".format(k, r[k])
        else:
            continue
        break
    if sql.endswith("set "):
        print("更新項目エラー")
        return r
    sql += " where StaffNo='{}'".format(r["StaffNo"])
    sql += " and Dt='{}'".format(r["Dt"])
    print(sql)
    conn.execute(sql)
    print("rowcount=", conn.execute("select @@rowcount").fetchone()[0])

    if calc:
        print("再計算", calc)
        sql = "select j.JCode, a.* from Atnd a, JGyobu j"
        sql += " where a.StaffNo='{}'".format(r["StaffNo"])
        sql += " and a.Dt='{}'".format(r["Dt"])
        sql += " and j.JGYOBU='0'"
        df = pd.read_sql(sql, conn)
        df = df.fillna("")
        print(df)
        for i, row in df.iterrows():
            for index, value in enumerate(row):
                if isinstance(value, str):
                    row[index] = value.rstrip()
            d0= row.to_dict()
            print(d0)
            import atnd
            d1 = d0.copy()
            d1 = atnd.calc(d1)
            print(d1)
            print("StartTm", d0["StartTm"], d1["StartTm"])
            print("FinishTm", d0["FinishTm"], d1["FinishTm"])
            print("Actual", d0["Actual"], d1["Actual"])
            print("Extra", d0["Extra"], d1["Extra"])
            print("Night", d0["Night"], d1["Night"])
            print("Dayoff", d0["Dayoff"], d1["Dayoff"])
            if "StartTm_i" in r.keys():
                r["StartTm"] = "{:%H:%M}".format(d1["StartTm"]) if d1["StartTm"] else ""
            if "FinishTm_i" in r.keys():
                r["FinishTm"] = "{:%H:%M}".format(d1["FinishTm"]) if d1["FinishTm"] else ""
            sql = "update Atnd set"
            if d0["Actual"] != d1["Actual"]:
                sql += ", Actual = {}".format(d1["Actual"])
                r["Actual"] = d1["Actual"]
            if d0["Extra"] != d1["Extra"]:
                sql += ", Extra = {}".format(d1["Extra"])
                r["Extra"] = d1["Extra"]
            if d0["Night"] != d1["Night"]:
                sql += ", Night = {}".format(d1["Night"])
                r["Night"] = d1["Night"]
            if d0["Dayoff"] != d1["Dayoff"]:
                sql += ", Dayoff = {}".format(d1["Dayoff"])
                r["Dayoff"] = d1["Dayoff"]
            if sql.endswith("set"):
                print("変更なし")
            else:
                sql = sql.replace("set,", "set ")
                sql += " where StaffNo='{}'".format(r["StaffNo"])
                sql += " and Dt='{}'".format(r["Dt"])
                print(sql)
                conn.execute(sql)
                print("rowcount=", conn.execute("select @@rowcount").fetchone()[0])
                
    if r["commit"] == "1":
        print("conn.commit()", end=";")
        conn.commit()
        print("ok")
    print("conn.close()", end=";")
    conn.close()
    print("ok")
    return r


def main0(r):
    print("main({})".format(r))

    print("pyodbc.connect({0})".format(r["dsn"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dsn"])
    print("ok")
    r["StaffNo"] = r["id"].split("_")[0]
    r["Dt"] = r["id"].split("_")[1]
    if len(r.keys() & {"Shift","StartTm_i","FinishTm_i"}) > 0:
        sql = "select top 1 j.JCode,a.* from Atnd a,Jgyobu j where j.JGYOBU='0' and a.StaffNo='{}' and a.Dt='{}'".format(r["StaffNo"],r["Dt"])
        print(sql)
        row = conn.execute(sql).fetchone()
        print(row)
        d = {}
        d["JCode"] = row.JCode.rstrip()
        d["Shift"] = r["Shift"].replace("_","") if r.get('Shift') else row.Shift
        d["BegTm"] = "{:%H:%M}".format(row.BegTm) if row.BegTm else ""
        d["BegTm_i"] = datetime.strptime(r["BegTm_i"], "%H:%M") if r.get('BegTm_i') else row.BegTm_i
        
        d["FinTm"] = "{:%H:%M}".format(row.FinTm) if row.FinTm else row.FinTm
        d["FinTm_i"] = datetime.strptime(r["FinTm_i"], "%H:%M") if r.get('FinTm_i') else row.FinTm_i

        d["StartTm"] = "{:%H:%M}".format(row.StartTm) if row.StartTm else ""
        d["StartTm_i"] = r.get('StartTm_i', "{:%H:%M}".format(row.StartTm_i) if row.StartTm_i else "")
        r["StartTm_i"] = d["StartTm_i"]
        d["FinishTm"] = "{:%H:%M}".format(row.FinishTm) if row.FinishTm else ""
        d["FinishTm_i"] = r.get('FinishTm_i', "{:%H:%M}".format(row.FinishTm_i) if row.FinishTm_i else "")
        r["FinishTm_i"] = d["FinishTm_i"]
       
        import atnd
        d = atnd.calc(d)
        print(d)
        r["StartTm"] = "{:%H:%M}".format(d["StartTm"]) if d["StartTm"] else ""
        r["FinishTm"] = "{:%H:%M}".format(d["FinishTm"]) if d["FinishTm"] else ""
        try:
            r["Actual"] = "{}".format(d["Actual"])
        except:
            pass
        try:
            r["Extra"] = "{}".format(d["Extra"])
        except:
            pass
        try:
            r["Night"] = "{}".format(d["Night"])
        except:
            pass
        try:
            r["Dayoff"] = "{}".format(d["Dayoff"])
        except:
            pass

    r["update"] = update(conn, r)
    print("conn.commit()", end=".")
    conn.commit()
    print("ok")
    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return r

def update(conn, data):
    sql = "update Atnd"
    st = " set"
    for d in data:
        if d not in ["JCode","StaffNo","Dt","dsn","id", "action"]:
            if d in ["BegTm_i","FinTm_i","StartTm_i","FinishTm_i"]:
                try:
                    sql += "{} {} = '{}'".format(st, d, datetime.strptime(data[d],"%H:%M"))
                    #tm = datetime.strptime(data[d],"%H:%M")
                    #sql += "{} {} = '{}'".format(st, d, data[d])
                except:
                    sql += "{} {} = null".format(st, d)
            elif d in ["StartTm","FinishTm"]:
                print(d, data[d], type(data[d]))
                if data[d] != "":
                    sql += "{} {} = '{}'".format(st, d, data[d][-8:])
                else:
                    sql += "{} {} = null".format(st, d)
            elif d in ["Actual_i","Extra_i","Night_i","Dayoff_i"]:
                if data[d]:
                    sql += "{} {} = '{}'".format(st, d, data[d])
                else:
                    sql += "{} {} = null".format(st, d)
            elif d in ["Late","Early","PTO"]:
                if data[d]:
                    sql += "{} {} = '{}'".format(st, d, data[d])
                else:
                    sql += "{} {} = 0".format(st, d)
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
        #ret = conn.execute("select @@rowcount").fetchone()[0]
        return sql
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
        form = cgi.FieldStorage(keep_blank_values= True)
        r = {}
        post = ""
        for c in form.keys():
            r[c] = form[c].value
            post += "{}={} ".format(c, form[c].value)
        r["dsn"] = form.getvalue("dsn","newsdc")
        if r["dsn"] == "":
            r["dsn"] = "newsdc"
        r["commit"] = form.getvalue("commit","1")
        sdc.log(post)
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(json.dumps(r, ensure_ascii=False, indent=4))
        """
        print('Content-Type:text/html; charset=UTF-8;\n')
        print(r)
        """
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("edit", help="", nargs="?", default="", type=str)
        parser.add_argument("--dsn", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--id", help="", default="", type=str)
        parser.add_argument("--commit", help="0 1", default="1", type=str)
        r = vars(parser.parse_args())
        print(r)
        print(r["edit"])
        r[r["edit"].split("=")[0]] = r["edit"].split("=")[1]
        print(r)
        r = main(r)
        import pprint
        pprint.pprint(r)

