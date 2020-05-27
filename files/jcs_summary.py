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

def main(dns, dt, limit):
    print("main({0}, {1}, {2})".format(dns, dt, limit))
    results = {}
    results["dns"] = dns
    results["dt"] = dt
    results["limit"] = limit
    print("pyodbc.connect({0})".format(dns), end=".", flush=True)
    conn = pyodbc.connect('DSN=' + dns)
    print(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), flush=True)

    results["data"] = load(conn, dt, limit)

    print("conn.close()", end=".", flush=True)
    conn.close()
    print("ok", flush=True)

    return results

def load(conn, dt, limit):
    print("load({0}, {1}, {2})".format(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), dt, limit))
    ### Pos指図票
    sql = """
select
 p.SHIJI_NO
,left(p.SHIJI_NO,6)		NohinNo
,SUBSTRING(p.SHIJI_NO,7,1)	NohinNo2
,p.BIKOU			DestCode
,p.ORDER_DT			DlvDt
,p.F_CLASS_CODE			Location
,p.HIN_GAI			MazdaPn
,p.SHIJI_QTY 			Qty
,p.S_CLASS_CODE			Pn
,n.EntID
,n.EntTm
,p.HAKKO_DT
,p.HIN_CHECK_TANTO
,p.HIN_CHECK_DATETIME
,p.HIN_CHECK_LABEL_CNT
,p.HIN_CHECK_GENPIN_CNT
,p.KAN_F
,p.KAN_DT
,p.BUNNOU_CNT
,p.UKEIRE_QTY
,p.CANCEL_F
,p.CANCEL_DATETIME
,p.TANTO_CODE
,p.SHONIN_CODE
,p.PRI_SHIJI
,p.N_CLASS_CODE
,p.PRINT_DATETIME
,itm.NameE
,itm.NameJ
,itm.SSpec
,itm.SType
,itm.GPn
from P_SSHIJI_O p
left outer join JcsNohin n
 on ((rtrim(n.NohinNo) + n.NohinNo2) = left(p.SHIJI_NO,7))
left outer join JcsItem itm
 on (p.HIN_GAI = itm.MazdaPn)
where p.ORDER_DT {0} '{1}'
order by
 DlvDt
,n.EntTm
,DestCode
,Location
,p.SHIJI_NO
""".format("=" if dt != "" else ">", dt if dt != "" else datetime.now().strftime('%Y%m%d'))
    r = []
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        prev = ""
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
#                if col is None:
#                    row[i] = ""
            if row.SHIJI_NO != prev:
                r.append(dict(zip(columns, row)))
            prev = row.SHIJI_NO
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

def decimal_default_proc(obj):
    if isinstance(obj, Decimal):
        return float(obj)
    print("decimal_default_proc:" + str(obj))
    return obj
    raise TypeError

class DatetimeEncoder(json.JSONEncoder):
    def default(self, obj):
        if isinstance(obj, str):
            return 'a'
        if isinstance(obj, datetime ):
            return obj.strftime("%Y-%m-%d %H:%M:%S")
        if isinstance(obj, date ):
            return obj.strftime("%Y-%m-%d")
        if isinstance(obj, Decimal):
            return float(obj)
        return json.JSONEncoder.default(self, obj)

def print_json(results):
    print('Content-Type:application/json; charset=UTF-8;\n')
    print(json.dumps(results,
#                     default=decimal_default_proc,
                     ensure_ascii=True,
                     indent=4,
                     sort_keys=False,
                     separators=(',', ': '),
                     cls=DatetimeEncoder))

if __name__ == "__main__":
    if 'REQUEST_METHOD' in os.environ:
        form = cgi.FieldStorage()
        dns = form.getvalue('dns', 'newsdc')
        dt = form.getvalue('dt', '')
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, dt, limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--dt", help="納入日", default="", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = main(args.dns, args.dt, args.limit)
        print_json(r)
        import pprint
        pprint.pprint(r)
