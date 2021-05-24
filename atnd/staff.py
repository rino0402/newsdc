# -*- coding: utf-8 -*-
import os
import sys
import io
import cgi
import cgitb
import pandas as pd
import pyodbc
import json
from datetime import date, datetime, timedelta
from decimal import Decimal
import traceback
import sdc

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dsn"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dsn"])
    print("ok")

    if r["action"] == "make":
        make(conn, r)
        conn.commit()
    elif r["action"] == "edit":
        edit(conn, r)
        conn.commit()
    else:
        r["df"] = load(conn, r)

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return r

def edit(conn, r):
    sql = "update Staff"
    sql += " set"
    for n in r:
        if n in ["Name","Post","Shift","Quit","QuitDt"]:
            sql += " " if sql.endswith('set') else ","
            if n in ["QuitDt"]:
                if r[n]:
                    sql += "{}='{}'".format(n,r[n])
                else:
                    sql += "{} = null".format(n)
            else:
                sql += "{}='{}'".format(n,r[n])
    sql += " where StaffNo='{}'".format(r["id"])
    r["sql"] = sql
    conn.execute(sql)
    r["rowcount"] = conn.execute("select @@rowcount").fetchone()[0]
    return r

def load(conn, r):
    sql = """
select
*
from Staff
"""
    if sdc.user().post:
        sql += " where Post='{}'".format(sdc.user().post)
    if r.get("quit"):
        sql += " and" if sql.find("where") > 0 else " where"
        sql += " QuitDt is null"
    sql += " order by Post, StaffNo"

    df = pd.read_sql(sql, conn)
    df["StaffNo"] = df["StaffNo"].str.rstrip()
    df["Name"] = df["Name"].str.rstrip()
    df["Post"] = df["Post"].str.rstrip()
    df["Shift"] = df["Shift"].str.rstrip()
    df["Quit"] = df["Quit"].str.rstrip()
    print(df)
    return df

def make(conn, r):
    print("make({})".format(r))
    sql = """
select distinct
 t.StaffNo
,t.Name
,t.Post
,t.Shift
,s.StaffNo sStaffNo 
,s.Name sName
,s.Post sPost
,s.Shift sShift
from timepack t
left outer join Staff s
 on (t.StaffNo = '000' + s.StaffNo)
"""
    cursor = conn.cursor()
    cursor.execute(sql)
    columns = [column[0] for column in cursor.description]
    for row in cursor.fetchall():
        for i, col in enumerate(row):
            if isinstance(col, str):
                row[i] = col.rstrip()
        d = dict(zip(columns, row))
        print(d, end=".")
        d["StaffNo"] = d["StaffNo"][-5:]
        d["Post"] = d["Shift"]
        if d["sStaffNo"]:
            print("update", end=".")
            update(conn, d)
        else:
            print("insert", end=".")
            insert(conn, d)
        print("")

def update(conn, data):
    sql = "update Staff"
    st = " set"
    for d in data:
        if d not in ["StaffNo","sStaffNo","sName","sShift","sPost"]:
            if data[d] == None:
                sql += "{} {} = null".format(st, d)
            else:
                sql += "{} {} = '{}'".format(st, d, data[d])
            st = ", "
    if st != ", ":
        return 0
    sql += " where StaffNo='{}'".format(data["StaffNo"])
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
    sql = "insert into Staff ("
    c = ""
    for d in data:
        if d not in ["sStaffNo","sName","sShift","sPost"]:
            if data[d]:
                sql += c + d
                c = ","
    sql += ") values ("
    for d in data:
        if d not in ["sStaffNo","sName"]:
            if data[d]:
                if type(data[d]) == str:
                    sql += "'{0}',".format(data[d].replace("'","''"))
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
        form = cgi.FieldStorage(keep_blank_values= True)
        r = {}
        r["action"] = ""
        for c in form.keys():
            r[c] = form[c].value
        if r["dsn"]:
            pass
        else:
            r["dsn"] = "newsdc"
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        #print(r["df"].to_json())
        if r["action"] == "edit":
            print(json.dumps(r, ensure_ascii=False, indent=4))
        else:
            print(r["df"].to_json(orient='table'))
        #print(r["df"].to_json(orient= 'split', force_ascii= True))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dsn", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--action", action="store_true", default=False)
        r = main(vars(parser.parse_args()))
        print(sdc)
        print(dir(sdc))
        print(sdc.user().name)
        print(sdc.user().post)
        
