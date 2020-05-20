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

def main(dns, test):
    print("main({0}, {1})".format(dns, test))
    results = {}
    results["dns"] = dns
    results["test"] = test

    print("pyodbc.connect({0})".format(dns), end=".")
    conn = pyodbc.connect('DSN=' + dns)
    print("ok")

    results["acorder"] = acorder(conn)

    if test:
        print("test")
    else:
        print("pyodbc.commit).".format(dns), end="")
        conn.commit()
        print("ok")
        
    print("conn.close()", end=".")
    conn.close()
    print("ok")

    return results

def acorder(conn):
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
    r = []
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
            print("{0} {1} {2}.".format(row.KEY_ID_NO, row.ODER_NO, row.KEY_SYUKA_YMD), end="")
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
    return r



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
        parser.add_argument("--test",action="store_true")
        args= parser.parse_args()
        r = main(args.dns, args.test)
#        import pprint
#        pprint.pprint(r)
