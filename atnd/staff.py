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

def main(r):
    print("main({})".format(r))
    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    make(conn, r)
    conn.commit()

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    
    return r

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
        form = cgi.FieldStorage()
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print('Content-Type:application/json; charset=UTF-8;\n')
        print(r["df"].to_json(orient= 'split', force_ascii= True))
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        r = main(vars(parser.parse_args()))
