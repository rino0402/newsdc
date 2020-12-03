# -*- coding: utf-8 -*-
import os
import sys
import io
import cgi
import re
import xlrd
import pandas as pd
import pyodbc
import json
from datetime import date, datetime, timedelta
from decimal import Decimal
#-------------------------------------------------------------
def main(r):
    print("main({})".format(r))

    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    r["list"] = list1(conn, r)

    print("conn.close()", end=".")
    conn.close()
    print("ok")
    return r
#-------------------------------------------------------------
def sqlwhere(sql, name, cond1 ,cond2):
    if cond1 == "":
        return ""
    w = " and" if "where" in sql else " where"
    if cond1 and cond2:
        w += " {} between '{}' and '{}'".format(name, cond1, cond2)
    elif "%" in cond1:
        w += " {} like '{}'".format(name, cond1)
    else:
        w += " {} = '{}'".format(name, cond1)
    return w
#-------------------------------------------------------------
def list1(conn, r):
    print("list1()")
    sql = "select {}".format(' top {}'.format(r["limit"]) if r["limit"] > 0 else '')
    sql += " * from JGyobu"
    sql += sqlwhere(sql, "JGYOBU", r["jgyobu"], "")
    sql += " order by 1,2,3"
    r = []
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
                elif col is None:
                    row[i] = ""
            d = dict(zip(columns, row))
            r.append(d)
    except pyodbc.ProgrammingError as e:
        print(e)
        raise
    except pyodbc.Error as e:
        print(e)
        raise
    except:
        print('error')
        raise
    return r
#----------------------------------------------------------------
def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError
#----------------------------------------------------------------
def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': ')))
#-------------------------------------------------------------
if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        r = {}
        r["dns"] = form.getvalue('dns', 'newsdc')
        r["jgyobu"] = form.getvalue('jgyobu', '')
        r["limit"] = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(r)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--jgyobu", help="", default="", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = {}
        r["dns"] = args.dns
        r["jgyobu"] = args.jgyobu
        r["limit"] = args.limit
        r = main(r)
        import pprint
        pprint.pprint(r)
