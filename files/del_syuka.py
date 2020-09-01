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

def main(r):
    print("main({})".format(r))

    print("pyodbc.connect({0})".format(r["dns"]), end=".")
    conn = pyodbc.connect('DSN=' + r["dns"])
    print("ok")

    r["acorder"] = acorder(conn, r)

    if r["test"]:
        print("test")
    else:
        print("pyodbc.commit).".format(r["dns"]), end="")
        conn.commit()
        print("ok")
        
    print("conn.close()", end=".")
    conn.close()
    print("ok")

    return r

def acorder(conn, r):
    sql = """
select
*
// KEY_ID_NO
//,ODER_NO
//,KEY_SYUKA_YMD
from del_syuka
where ODER_NO <> ''
and (
ODER_NO in (select IdNo from acorder)
or
ODER_NO in (select Id from aczan)
)
"""
    if r["kenpin"] != 0:
        sql += " or KENPIN_YMD >= '{0:%Y%m%d}'".format(date.today() - timedelta(r["kenpin"]))
    print(sql)
    ret = []
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
            sql = """
insert into y_syuka ({0})
select distinct top 1 {0}
from del_syuka
where KEY_ID_NO='{1}'
""".format(",".join(map(str, columns)), row.KEY_ID_NO)
#            print(sql)
            print("{0} {1} {2} {3}.".format(row.KEY_ID_NO, row.ODER_NO, row.KEY_SYUKA_YMD, row.KENPIN_YMD), end="")
            conn.execute(sql)
            print("y_syuka:{}.".format(conn.execute("select @@rowcount").fetchone()[0]), end="")
            conn.execute("delete from del_syuka where KEY_ID_NO='{}'".format(row.KEY_ID_NO))
            print("del_syuka:{}".format(conn.execute("select @@rowcount").fetchone()[0]))
                
    except pyodbc.ProgrammingError as e:
        print(e)
        print(sql)
        raise
    except pyodbc.Error as e:
        print(e)
        print(sql)
        raise
    except:
        print('error')
        raise
    return ret

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    raise TypeError

def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': ')))

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        dns = form.getvalue('dns', 'newsdc')
        pallet_no = form.getvalue('pallet_no', '')
        id_no = form.getvalue('id_no', '')
        case_qty = form.getvalue('case_qty', '')
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, pallet_no, id_no ,case_qty , limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--kenpin", help="KENPIN_YMD の過去？日分", default=0, type=int)
        parser.add_argument("--test",action="store_true")
        r = main(vars(parser.parse_args()))
#        args= parser.parse_args()
#        r = main(args.dns, args.test)
#        import pprint
#        pprint.pprint(r)
