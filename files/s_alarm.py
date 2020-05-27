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

def main(dns, limit):
    print("main({0}, {1})".format(dns, limit))
    results = {}
    results["dns"] = dns
    results["limit"] = limit
    print("pyodbc.connect({0})".format(dns), end=".", flush=True)
    conn = pyodbc.connect('DSN=' + dns)
    print(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), flush=True)

    results["data"] = load(conn, limit)

    print("conn.close()", end=".", flush=True)
    conn.close()
    print("ok", flush=True)

    return results

def load(conn, limit):
    print("load({0}, {1})".format(conn.getinfo(pyodbc.SQL_DATA_SOURCE_NAME), limit))
    sql = """
select
 i.G_SYUSHI
,p.C_NAME
,i.HIN_GAI
,i.HIN_NAME
,z.qty zqty
,convert(i.HOJYU_P,sql_decimal) HOJYU_P
,ifnull(z.qty,0) - convert(i.HOJYU_P,sql_decimal) FitQty
,o.zanQty zanQty
,ifnull(z.qty,0) - convert(i.HOJYU_P,sql_decimal) + ifnull(o.zanQty,0) Fit2Qty
,i.LAST_NYU_DT
,i.LAST_SYU_DT
,o.cnt
,o.minDt
,o.maxDt
,o.maxKAN_DT
from item i
left outer join p_code p on (p.C_Code = i.G_SYUSHI and p.DATA_KBN = '03')
left outer join (
	select
	 HIN_GAI
	,sum(convert(YUKO_Z_QTY,SQL_DECIMAL)) qty
	from zaiko
	where jgyobu='S'
	and naigai='1'
	group by
	 HIN_GAI
) z on (i.HIN_GAI = z.HIN_GAI)
left outer join (
	select
	 HIN_GAI
	,sum(if(KAN_F = '0' and CANCEL_F='0',convert(ORDER_QTY,sql_decimal),0)) zanQty
	,count(*) cnt
	,max(KAN_DT)	maxKAN_DT
	,min(Y_NOUKI_DT) minDt
	,max(Y_NOUKI_DT) maxDt
	from P_SHORDER
	where JGYOBU='S'
	and NAIGAI='1'
//	and KAN_F = '0'
	and CANCEL_F = '0'
	group by
	 HIN_GAI
) o on (i.HIN_GAI = o.HIN_GAI)
where i.jgyobu='S'
and i.naigai='1'
and i.HIN_GAI <> 'TEST'
and convert(i.HOJYU_P,sql_decimal) > 0
order by
 Fit2Qty
"""
    r = []
    try:
        cursor = conn.cursor()
        cursor.execute(sql)
        columns = [column[0] for column in cursor.description]
        for row in cursor.fetchall():
            for i, col in enumerate(row):
                if isinstance(col, str):
                    row[i] = col.rstrip()
            r.append(dict(zip(columns, row)))
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
        limit = int(form.getvalue('limit', 0))
        sys.stdout = None
        r = main(dns, limit)
        sys.stdout = sys.__stdout__
        print_json(r)
    else:
        import argparse
        parser = argparse.ArgumentParser()
        parser.add_argument("--dns", help="default: newsdc", default="newsdc", type=str)
        parser.add_argument("--limit", help="default: 0", default=0, type=int)
        args= parser.parse_args()
        r = main(args.dns, args.limit)
        print_json(r)
#        import pprint
#        pprint.pprint(r)
